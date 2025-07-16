# app_openpyxl.py
"""
Streamlit CO₂ Estimator – cloud-ready
• openpyxl  → loads workbook (.xlsx)
• xlcalculator → live recalculation when possible (fallback to static values)
"""

# ── Imports ────────────────────────────────────────────────────────────────
from pathlib import Path
import streamlit as st
from openpyxl import load_workbook
from xlcalculator import ModelCompiler, Evaluator
import requests
from bs4 import BeautifulSoup

# ── Streamlit config ───────────────────────────────────────────────────────
st.set_page_config(page_title="FuelTrust CO₂ Ship Estimator", layout="wide")
st.title("🚢 Ship Estimator – Powered by FuelTrust")

# ── Excel workbook path ────────────────────────────────────────────────────
EXCEL_PATH = Path("CO2EmissionsEstimator3.xlsx")
if not EXCEL_PATH.exists():
    st.error(f"❌ {EXCEL_PATH} not found.")
    st.stop()

# ── Load workbook + evaluator (cached) ─────────────────────────────────────
@st.cache_resource(show_spinner="Loading Excel model…")
def load_model(path: Path):
    wb = load_workbook(path, data_only=False, keep_links=False)  # Changed to data_only=False to access formulas if needed
    mc = ModelCompiler()
    model = mc.read_and_parse_archive(str(path))
    ev = Evaluator(model)
    return wb, ev

wb, ev = load_model(EXCEL_PATH)
ship_sheet   = wb["Ship Estimator"]
lookup_sheet = wb["LookupTables"]

# ── Helper functions ───────────────────────────────────────────────────────
xl_addr = lambda sheet, cell: f"'{sheet}'!{cell.upper()}"

def _flatten(val):
    """Convert [220] or [[220]] → 220 so widgets get a scalar."""
    while isinstance(val, list) and len(val) == 1:
        val = val[0]
    return val

def set_value(cell, value):
    ship_sheet[cell].value = value
    try:
        ev.set_cell_value(xl_addr("Ship Estimator", cell), value)
    except Exception:
        pass  # evaluator may not handle ranges; ignore

def get_value(cell):
    """
    1) Try xlcalculator (live).
    2) Flatten 1-element lists.
    3) If fail or non-numeric, use Python fallback calc.
    4) Else cached value.
    """
    try:
        val = ev.evaluate(xl_addr("Ship Estimator", cell))
        val = _flatten(val)
        if isinstance(val, (int, float)):
            return val
    except Exception:
        pass
    # Fallback to Python calc for key cells
    val = calculate_fallback(cell)
    if val is not None:
        return val
    # Final fallback to cached
    cached = ship_sheet[cell].value
    return cached if isinstance(cached, (int, float)) else None

def calculate_fallback(cell):
    """Python mimic of Excel formulas for outputs if evaluator fails."""
    sea_days = get_value("B16") or 0
    port_days = get_value("B17") or 0
    total_days = sea_days + port_days
    if total_days == 0:
        return None
    if cell == "B18":  # Annual AVG NM = SEA DAYS * Avg nm / SEA Day + PORT DAYS * Avg nm / PORT day
        return (sea_days * (get_value("B7") or 0)) + (port_days * (get_value("B8") or 0))
    if cell == "E6":  # Average Daily Fuel Use (MT) = (SEA Fuel * SEA DAYS + PORT Fuel * PORT DAYS) / total_days
        return ((get_value("B10") or 0) * sea_days + (get_value("B11") or 0) * port_days) / total_days
    if cell == "E7":  # Annex II Emissions CO2 = Total fuel * Cf (from fuel type)
        fuel_type = ship_sheet["B19"].value
        cf_row = fuel_options.index(fuel_type) if fuel_type in fuel_options else 0
        cf = lookup_sheet.cell(row=43 + cf_row, column=2).value  # Cf from A43:B64
        total_fuel = (get_value("B10") or 0) * sea_days + (get_value("B11") or 0) * port_days
        return total_fuel * cf
    if cell == "E8":  # Measured CO2 Estimate = E7 * (1 - B21)
        return (get_value("E7") or 0) * (1 - (get_value("B21") or 0))
    if cell == "E9":  # CO2 Reduction = E7 - E8
        return (get_value("E7") or 0) - (get_value("E8") or 0)
    if cell == "E10":  # EU CO2 = E7 * (B12 + B13 * 0.5)
        return (get_value("E7") or 0) * ((get_value("B12") or 0) + (get_value("B13") or 0) * 0.5)
    if cell == "E11":  # EU ETS (2024) Liability = E10 * B26 * 0.4  (assuming 40% for 2024)
        return (get_value("E10") or 0) * (get_value("B26") or 0) * 0.4
    if cell == "E12":  # EU Eligible CO2 Reductions = E9 * (B12 + B13 * 0.5)
        return (get_value("E9") or 0) * ((get_value("B12") or 0) + (get_value("B13") or 0) * 0.5)
    if cell == "E13":  # Annex-II CO2 (2025→) = E7 * 1.50419  (CO2e factor from B69)
        return (get_value("E7") or 0) * 1.50419
    if cell == "E14":  # Measured CO2e Estimate = E13 * (1 - 0.0412)  (Avg CO2E reduction from B67)
        return (get_value("E13") or 0) * (1 - 0.0412)
    if cell == "E15":  # Measured CO2e Reduction = E13 - E14
        return (get_value("E13") or 0) - (get_value("E14") or 0)
    if cell in ["E16", "E17", "E18", "E19"]:  # Savings € 2025-2028 = E15 * EUA price * liability %
        year = {"E16": 2025, "E17": 2026, "E18": 2027, "E19": 2028}[cell]
        liability_pct = [0.4, 0.7, 1.0, 1.0][year - 2025]  # From Price Models tab
        eua = get_value("B26") or 67.6  # Fallback to default
        return (get_value("E15") or 0) * eua * liability_pct
    if cell == "E21":  # Avg Fraud Savings / yr = E7 * B23 * B26
        return (get_value("E7") or 0) * (get_value("B23") or 0) * (get_value("B26") or 0)
    return None

def safe_metric(label, value, prefix=""):
    if isinstance(value, (int, float)):
        st.metric(label, f"{prefix}{value:,.2f}")
    else:
        st.metric(label, "–")

def get_range_values(range_name):
    """Fetch values from a named range as a flat list (assuming column/row)."""
    if range_name not in wb.defined_names:
        return None
    defined = wb.defined_names[range_name]
    values = []
    for sheet_title, coord in defined.destinations:
        sheet = wb[sheet_title]
        for row in sheet[coord]:
            for cell in row:
                values.append(cell.value)
    return values

# ── Sidebar – user inputs ──────────────────────────────────────────────────
st.sidebar.header("Adjust Estimator Inputs")

# Ship type dropdown
ship_type_list = [c.value for c in lookup_sheet["A"][1:39] if c.value]  # Up to row 39 from Excel
current_ship_type = ship_sheet["B6"].value
selected_ship_type = st.sidebar.selectbox(
    "Ship Type", ship_type_list,
    index=ship_type_list.index(current_ship_type) if current_ship_type in ship_type_list else 0,
)
if selected_ship_type != current_ship_type:
    set_value("B6", selected_ship_type)
    # Update dependent cells via Python lookup (mimics INDEX/MATCH)
    dependents = {
        "B7": "SeaDayNM",   # Average nm / SEA Day
        "B8": "PortDayNM",  # Average nm / PORT Day
        "B10": "SeaDayMT",  # Avg SEA Fuel use (MT)
        "B11": "PortDayMT", # Avg PORT Fuel use (MT)
        "B16": "SeaDays",   # SEA DAYS
        "B17": "PortDays",  # PORT DAYS
    }
    ship_index = ship_type_list.index(selected_ship_type)  # 0-based
    for cell, range_name in dependents.items():
        values = get_range_values(range_name)
        if values and ship_index < len(values):
            new_val = values[ship_index]
            if isinstance(new_val, (int, float)):
                set_value(cell, new_val)
    # Trigger fallback for B18
    set_value("B18", calculate_fallback("B18"))

# Numeric inputs - add on_change to trigger recalcs
num_inputs = {
    "B7": "Average nm / SEA Day",
    "B8": "Average nm / PORT Day",
    "B10": "Avg SEA Fuel use (MT)",
    "B11": "Avg PORT Fuel use (MT)",
    "B16": "SEA DAYS",
    "B17": "PORT DAYS",
    "B18": "Annual AVG NM",
}
for cell, label in num_inputs.items():
    default_val = get_value(cell) or 0.0
    new_val = st.sidebar.number_input(label, value=float(default_val))
    if new_val != default_val:
        set_value(cell, new_val)
        if cell in ["B7", "B8", "B16", "B17"]:  # Recalc B18 if these change
            set_value("B18", calculate_fallback("B18"))

# Percentage sliders
slider_inputs = {
    "B12": "% Voyages EU-EU",
    "B13": "% In/Out EU",
    "B14": "Non-EU %",
}
for cell, label in slider_inputs.items():
    pct_default = int((get_value(cell) or 0) * 100)
    new_pct = st.sidebar.slider(label, 0, 100, pct_default)
    if new_pct != pct_default:
        set_value(cell, new_pct / 100)

# Fuel type dropdown
fuel_options = [row[0].value for row in lookup_sheet["A43:A64"] if row[0].value]
current_fuel = ship_sheet["B19"].value
fuel_type = st.sidebar.selectbox(
    "Default SEA Fuel", fuel_options,
    index=fuel_options.index(current_fuel) if current_fuel in fuel_options else 0,
)
if fuel_type != current_fuel:
    set_value("B19", fuel_type)

# CO₂ overage and fraud
co2_over_pct = int((get_value("B21") or 0) * 100)
new_co2_over = st.sidebar.number_input("Avg CO₂ Overage (%)", value=co2_over_pct, min_value=0)
if new_co2_over != co2_over_pct:
    set_value("B21", new_co2_over / 100)

fraud_pct = int((get_value("B23") or 0) * 100)
new_fraud = st.sidebar.number_input("Avg Fraud (%)", value=fraud_pct, min_value=0)
if new_fraud != fraud_pct:
    set_value("B23", new_fraud / 100)

# ── Live EUA price ─────────────────────────────────────────────────────────
def get_live_eua_price():
    try:
        url = "https://www.eex.com/en/market-data/environmental-markets/spot-market/european-emission-allowances"
        soup = BeautifulSoup(requests.get(url, timeout=8).text, "html.parser")
        price_text = soup.find("td", string="2021-2030").find_next("td").text.strip()
        return float(price_text.replace(",", "."))
    except Exception:
        return None

live_price  = get_live_eua_price()
sidebar_val = float(live_price or get_value("B26") or 67.6)  # Fallback to 67.6 if None
sidebar_price = st.sidebar.number_input("Current EUA Price (€)", value=sidebar_val)
if sidebar_price != get_value("B26"):
    set_value("B26", sidebar_price)

# ── Estimator output ───────────────────────────────────────────────────────
st.subheader("📊 Estimator Results")

col1, col2 = st.columns(2)
metrics_col1 = {
    "Average Daily Fuel Use (MT)": "E6",
    "Annex II Emissions CO₂": "E7",
    "Measured CO₂ Estimate": "E8",
    "CO₂ Reduction": "E9",
    "EU CO₂": "E10",
    "EU ETS (2024) Liability": "E11",
    "EU Eligible CO₂ Reductions": "E12",
}
metrics_col2 = {
    "Annex-II CO₂ (2025→)": "E13",
    "Measured CO₂e Estimate": "E14",
    "Measured CO₂e Reduction": "E15",
    "SAVINGS € 2025": "E16",
    "SAVINGS € 2026": "E17",
    "SAVINGS € 2027": "E18",
    "SAVINGS € 2028": "E19",
    "Avg Fraud Savings / yr": "E21",
}

with col1:
    for lbl, adr in metrics_col1.items():
        safe_metric(lbl, get_value(adr), "€ " if "€" in lbl else "")

with col2:
    for lbl, adr in metrics_col2.items():
        safe_metric(lbl, get_value(adr), "€ " if "€" in lbl else "")

st.info("📌 Excel charts are removed in this version. Replace with Streamlit charts if needed.")
