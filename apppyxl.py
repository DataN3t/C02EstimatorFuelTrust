# app_openpyxl.py
"""
Streamlit CO₂ Estimator – cloud‑ready
• openpyxl          → loads workbook (.xlsx)
• xlcalculator      → live recalculation when possible (fallback to static values)
"""

# ── Imports ────────────────────────────────────────────────────────────────
from pathlib import Path
import streamlit as st
from openpyxl import load_workbook
from xlcalculator import ModelCompiler, Evaluator
import requests
from bs4 import BeautifulSoup

# ── Streamlit config ───────────────────────────────────────────────────────
st.set_page_config(
    page_title="FuelTrust CO₂ Ship Estimator",
    layout="wide"
)
st.title("🚢 Ship Estimator – Powered by FuelTrust")

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

# ── Helper functions ───────────────────────────────────────────────────────
xl_addr = lambda sheet, cell: f"'{sheet}'!{cell.upper()}"

def set_value(cell, value):
    ship_sheet[cell].value = value
    try:
        ev.set_cell_value(xl_addr("Ship Estimator", cell), value)
    except Exception:
        pass

def get_value(cell):
    try:
        val = ev.evaluate(xl_addr("Ship Estimator", cell))
    except Exception:
        val = None
    if isinstance(val, (int, float)):
        return val
    cached = ship_sheet[cell].value
    return cached if isinstance(cached, (int, float)) else None

def safe_metric(label, value, prefix=""):
    if isinstance(value, (int, float)):
        st.metric(label, f"{prefix}{value:,.2f}")
    else:
        st.metric(label, "–")

# ── Sidebar – user inputs ──────────────────────────────────────────────────
st.sidebar.header("Adjust Estimator Inputs")

# Ship type dropdown
ship_type_list = [cell.value for cell in lookup_sheet["A"][1:] if cell.value]
current_ship_type = ship_sheet["B6"].value
selected_ship_type = st.sidebar.selectbox(
    "Ship Type", ship_type_list,
    index=ship_type_list.index(current_ship_type) if current_ship_type in ship_type_list else 0,
)
if selected_ship_type != current_ship_type:
    set_value("B6", selected_ship_type)

# Numeric inputs
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

# Percentage sliders
slider_inputs = {
    "B12": "% Voyages EU‑EU",
    "B13": "% In/Out EU",
    "B14": "Non‑EU %",
}
for cell, label in slider_inputs.items():
    pct_default = int(get_value(cell) or 0)
    new_pct = st.sidebar.slider(label, 0, 100, pct_default)
    if new_pct != pct_default:
        set_value(cell, new_pct)

# Fuel type dropdown
fuel_options = [r[0].value for r in lookup_sheet["A43:A64"] if r[0].value]
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

live_price = get_live_eua_price()
sidebar_price = st.sidebar.number_input("Current EUA Price (€)", value=float(live_price or get_value("B26") or 0))
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
    "Annex‑II CO₂ (2025→)": "E13",
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
