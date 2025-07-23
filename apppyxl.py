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
st.title("Ship Estimator – Powered by FuelTrust")

# ── Excel workbook path ────────────────────────────────────────────────────
EXCEL_PATH = Path("CO2EmissionsEstimator3.xlsx")
if not EXCEL_PATH.exists():
    st.error(f"❌ {EXCEL_PATH} not found.")
    st.stop()

# ── Load workbook + evaluator (cached) ─────────────────────────────────────
@st.cache_resource(show_spinner="Loading Excel model…")
def load_model(path: Path):
    wb = load_workbook(path, data_only=True, keep_links=False)
    mc = ModelCompiler()
    model = mc.read_and_parse_archive(str(path))
    ev = Evaluator(model)
    return wb, ev

wb, ev = load_model(EXCEL_PATH)
ship_sheet   = wb["Ship Estimator"]
lookup_sheet = wb["LookupTables"]

# Define fuel_options globally here, as it's static and needed in calculate_fallback
fuel_options = [row[0].value for row in lookup_sheet["A43:A64"] if row[0].value]

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
        set_value(cell, val)  # Update sheet with calculated value for chaining
        return val
    # Final fallback to cached
    cached = ship_sheet[cell].value
    return cached if isinstance(cached, (int, float)) else None

def calculate_fallback(cell):
    """Python mimic of Excel formulas for outputs if evaluator fails."""
    # Fetch raw values directly to avoid recursion, with safe float conversion
    def safe_float(cell_key, sheet=ship_sheet):
        val = sheet[cell_key].value
        try:
            return float(val) if val is not None else 0.0
        except (ValueError, TypeError):
            return 0.0

    sea_days = safe_float("B16")
    port_days = safe_float("B17")
    total_days = sea_days + port_days
    if total_days == 0:
        return None
    if cell == "B18":  # Annual AVG NM = SEA DAYS * Avg nm / SEA Day + PORT DAYS * Avg nm / PORT day
        nm_sea = safe_float("B7")
        nm_port = safe_float("B8")
        val = (sea_days * nm_sea) + (port_days * nm_port)
        return val
    if cell == "E6":  # Average Daily Fuel Use (MT) = (SEA Fuel * SEA DAYS + PORT Fuel * PORT DAYS) / total_days
        fuel_sea = safe_float("B10")
        fuel_port = safe_float("B11")
        val = (fuel_sea * sea_days + fuel_port * port_days) / total_days
        return val
    if cell == "E7":  # Annex II Emissions CO2 = Total fuel * Cf (from fuel type)
        fuel_type = ship_sheet["B19"].value
        cf_row = fuel_options.index(fuel_type) if fuel_type in fuel_options else 0
        cf = safe_float(f"B{43 + cf_row}", sheet=lookup_sheet)  # Cf from LookupTables B43+
        total_fuel = (safe_float("B10") * sea_days + safe_float("B11") * port_days)
        val = total_fuel * cf
        return val
    if cell == "E8":  # Measured CO2 Estimate = E7 * (1 - B21)
        co2_over = safe_float("B21")
        val = (get_value("E7") or 0) * (1 - co2_over)
        return val
    if cell == "E9":  # CO2 Reduction = E7 - E8
        val = (get_value("E7") or 0) - (get_value("E8") or 0)
        return val
    if cell == "E10":  # EU CO2 = E7 * (B12 + B13 * 0.5)
        eu_eu_pct = safe_float("B12")
        in_out_pct = safe_float("B13")
        val = (get_value("E7") or 0) * (eu_eu_pct + in_out_pct * 0.5)
        return val
    if cell == "E11":  # EU ETS (2024) Liability = E10 * B26 * 0.4
        eua_price = safe_float("B26")
        val = (get_value("E10") or 0) * eua_price * 0.4
        return val
    if cell == "E12":  # EU Eligible CO2 Reductions = E9 * (B12 + B13 * 0.5)
        eu_eu_pct = safe_float("B12")
        in_out_pct = safe_float("B13")
        val = (get_value("E9") or 0) * (eu_eu_pct + in_out_pct * 0.5)
        return val
    if cell == "E13":  # Annex-II CO2 (2025→) = E7 * 1.50419
        val = (get_value("E7") or 0) * 1.50419
        return val
    if cell == "E14":  # Measured CO2e Estimate = E13 * (1 - 0.0412)
        val = (get_value("E13") or 0) * (1 - 0.0412)
        return val
    if cell == "E15":  # Measured CO2e Reduction = E13 - E14
        val = (get_value("E13") or 0) - (get_value("E14") or 0)
        return val
    if cell in ["E16", "E17", "E18", "E19"]:  # Savings € 2025-2028 = E15 * EUA price * liability %
        year = {"E16": 2025, "E17": 2026, "E18": 2027, "E19": 2028}[cell]
        liability_pct = [0.4, 0.7, 1.0, 1.0][year - 2025]
        eua = safe_float("B26")
        val = (get_value("E15") or 0) * eua * liability_pct
        return val
    if cell == "E21":  # Avg Fraud Savings / yr = E7 * B23 * B26
        fraud_pct = safe_float("B23")
        eua = safe_float("B26")
        val = (get_value("E7") or 0) * fraud_pct * eua
        return val
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

# Use a form to batch inputs and update only on refresh button
with st.sidebar.form(key="estimator_form"):
    # Ship type dropdown
    ship_type_list = [c.value for c in lookup_sheet["A"][1:39] if c.value]
    current_ship_type = ship_sheet["B6"].value
    selected_ship_type = st.selectbox(
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
        ship_index = ship_type_list.index(selected_ship_type)
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
        new_val = st.number_input(label, value=float(default_val))
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
        new_pct = st.slider(label, 0, 100, pct_default)
        if new_pct != pct_default:
            set_value(cell, new_pct / 100)

    # Fuel type dropdown
    current_fuel = ship_sheet["B19"].value
    fuel_type = st.selectbox(
        "Default SEA Fuel", fuel_options,
        index=fuel_options.index(current_fuel) if current_fuel in fuel_options else 0,
    )
    if fuel_type != current_fuel:
        set_value("B19", fuel_type)

    # CO₂ overage and fraud
    co2_over_pct = int((get_value("B21") or 0) * 100)
    new_co2_over = st.number_input("Avg CO₂ Overage (%)", value=co2_over_pct, min_value=0)
    if new_co2_over != co2_over_pct:
        set_value("B21", new_co2_over / 100)

    fraud_pct = int((get_value("B23") or 0) * 100)
    new_fraud = st.number_input("Avg Fraud (%)", value=fraud_pct, min_value=0)
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
    sidebar_val = float(live_price or get_value("B26") or 67.6)
    sidebar_price = st.number_input("Current EUA Price (€)", value=sidebar_val)
    if sidebar_price != get_value("B26"):
        set_value("B26", sidebar_price)

    # Refresh button as form submit
    refresh_button = st.form_submit_button("Refresh")

# ── Estimator output ───────────────────────────────────────────────────────
st.subheader("Estimator Results:")

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

# Add the dynamic CTA at the bottom of the results
st.markdown("---")
co2e_reduction = get_value("E15") or 0.0
co2e_reduction_estimate = get_value("E14") or 0.0
st.markdown("""
<div style="text-align: center; background-color: #e6f7ff; padding: 20px; border-radius: 10px; border: 1px solid #91d5ff;">
<h2 style="color: #1A8C1A;">Estimated FuelTrust CO₂e Reduction Per Vessel Type: <strong>{co2e_reduction:,.2f} MT</strong></h2>
<h2 style="color: #1A8C1A;">Measured CO2e Estimate For Floras CO2e-Offset Calculator Below: <strong>{co2e_reduction_estimate:,.2f} MT</strong></h2>
<p style="font-size: 18px;">Unlock the exact decarb and in-depth insights tailored for your vessels.</p>
<a href="https://dk2advisor.com/getintouch" style="background-color: #1890ff; color: white; font-size: 16px; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;">Get in Touch Today</a>
</div>
""".format(co2e_reduction=co2e_reduction, co2e_reduction_estimate=co2e_reduction_estimate), unsafe_allow_html=True)

#st.info("📌 Excel charts are removed in this version. Replace with Streamlit charts if needed.")
