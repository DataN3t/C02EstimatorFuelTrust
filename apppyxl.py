# app_openpyxl.py
"""
Streamlit CO₂ Estimator – cloud‑ready
Replaces xlwings/Excel‑COM dependency with pure‑Python stack:
    • openpyxl         – load the workbook
    • xlcalculator     – evaluate all worksheet formulas
    • pandas           – convenience for tabular ops / dropdown lists

‼️ Path note ‼️
For local testing you asked to keep the absolute Windows path you’re used to.  
When you later deploy to Streamlit Cloud you must switch this back to a repo‑relative path, otherwise the file will not exist on the server.
"""

# ---- Imports --------------------------------------------------------------
import os
from pathlib import Path
import streamlit as st
from openpyxl import load_workbook
from xlcalculator import ModelCompiler, Evaluator
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------------------------------------------------------
# Configuration
# ----------------------------------------------------------------------------
st.set_page_config(
    page_title="FuelTrust CO₂ Ship Estimator",
    layout="wide",
    menu_items={"Report a bug": None, "About": None},
)

st.title("🚢 Ship Estimator – Powered by FuelTrust")

# ----------------------------------------------------------------------------
# Excel path – using your original absolute Windows path for local runs
# ----------------------------------------------------------------------------
EXCEL_PATH = Path("CO2EmissionsEstimator3.xlsx")
if not EXCEL_PATH.exists():
    st.error("❌ Excel file not found in repo directory.")
    st.stop()

# ----------------------------------------------------------------------------
# ----------------------------------------------------------------------------
# Helper – load workbook + build xlcalculator evaluator (cached per session)
# ----------------------------------------------------------------------------
@st.cache_resource(show_spinner="Loading Excel model…")
def load_model(path: Path):
    wb = load_workbook(path, data_only=False, keep_links=False)
    mc = ModelCompiler()
    model = mc.read_and_parse_archive(str(path))
    ev = Evaluator(model)
    return wb, ev

# ----------------------------------------------------------------------------
# Load workbook + evaluator
# ----------------------------------------------------------------------------
wb, ev = load_model(EXCEL_PATH)

# Kick-start evaluation (force one cell to compute)
ev.evaluate("'Ship Estimator'!E6")  # trigger calculation once

ship_sheet   = wb["Ship Estimator"]
lookup_sheet = wb["LookupTables"]

# ----------------------------------------------------------------------------
# Excel <--> xlcalculator helpers
# ----------------------------------------------------------------------------
xl_addr = lambda sheet, cell: f"'{sheet}'!{cell.upper()}"

def set_value(cell, value):
    ship_sheet[cell].value = value
    ev.set_cell_value(xl_addr("Ship Estimator", cell), value)

get_value = lambda cell: ev.evaluate(xl_addr("Ship Estimator", cell))

# # --- DEBUG probe -------------------------------------------------
# probe_cells = ["E6", "E7", "E11"]
# for c in probe_cells:
#     try:
#         val = get_value(c)
#         st.write(f"🔍 DEBUG {c} →", val)
#     except Exception as e:
#         st.error(f"⚠️ Error while evaluating {c}: {e}")
# -----------------------------------------------------------------


# ----------------------------------------------------------------------------
# Safe metric helper
# ----------------------------------------------------------------------------

def safe_metric(label, value, prefix=""):
    if value is None:
        st.metric(label, "–")
    else:
        try:
            st.metric(label, f"{prefix}{value:,.2f}")
        except Exception:
            st.metric(label, f"{prefix}{value}")

# ----------------------------------------------------------------------------
# Sidebar – user inputs (mirrors original)
# ----------------------------------------------------------------------------

st.sidebar.header("Adjust Estimator Inputs")

# Ship Type dropdown ---------------------------------------------------------
# ship_type_list = [c.value for c in lookup_sheet["A2":"A"] if c[0].value]
ship_type_list = [cell.value for cell in lookup_sheet["A"][1:] if cell.value]
current_ship_type = ship_sheet["B6"].value
selected_ship_type = st.sidebar.selectbox(
    "Ship Type",
    ship_type_list,
    index=ship_type_list.index(current_ship_type) if current_ship_type in ship_type_list else 0,
)
if selected_ship_type != current_ship_type:
    set_value("B6", selected_ship_type)

# Numeric inputs -------------------------------------------------------------
# num_inputs = {
#     "B7": "Average nm / SEA Day",
#     "B8": "Average nm / PORT Day",
#     "B10": "Avg SEA Fuel use (MT)",
#     "B11": "Avg PORT Fuel use (MT)",
#     "B16": "SEA DAYS",
#     "B17": "PORT DAYS",
#     "B18": "Annual AVG NM",
# }
# for cell, label in num_inputs.items():
#     default = ship_sheet[cell].value or 0.0
#     new_val = st.sidebar.number_input(label, value=float(default))
#     if new_val != default:
#         set_value(cell, new_val)

# Numeric inputs -------------------------------------------------------------
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
    default_val = get_value(cell)            # evaluated numeric result
    if default_val is None:
        default_val = 0.0
    new_val = st.sidebar.number_input(label, value=float(default_val))
    if new_val != default_val:
        set_value(cell, new_val)



# Percentage sliders ---------------------------------------------------------
# slider_inputs = {
#     "B12": "% of Voyages EU to EU",
#     "B13": "% of Voyages in/out of EU",
#     "B14": "Non‑EU % of Voyages",
# }
# for cell, label in slider_inputs.items():
#     pct_default = ship_sheet[cell].value or 0
#     new_pct = st.sidebar.slider(label, 0, 100, int(pct_default))
#     if new_pct != pct_default:
#         set_value(cell, new_pct)

slider_inputs = {
    "B12": "% of Voyages EU to EU",
    "B13": "% of Voyages in/out of EU",
    "B14": "Non-EU % of Voyages",
}

for cell, label in slider_inputs.items():
    pct_default = get_value(cell) or 0
    new_pct = st.sidebar.slider(label, 0, 100, int(pct_default))
    if new_pct != pct_default:
        set_value(cell, new_pct)


# Fuel type dropdown ---------------------------------------------------------
# fuel_options = [c.value for c in lookup_sheet["A43":"A64"] if c[0].value]
fuel_options = [row[0].value for row in lookup_sheet["A43:A64"] if row[0].value]

current_fuel = ship_sheet["B19"].value
fuel_type = st.sidebar.selectbox(
    "Default SEA Fuel Type",
    fuel_options,
    index=fuel_options.index(current_fuel) if current_fuel in fuel_options else 0,
)
if fuel_type != current_fuel:
    set_value("B19", fuel_type)

# Average CO₂ Overage & Fraud -----------------------------------------------
# co2_over_pct = (ship_sheet["B21"].value or 0) * 100
# new_co2_over = st.sidebar.number_input("Average CO₂ Overage (%)", value=co2_over_pct)
# if new_co2_over != co2_over_pct:
#     set_value("B21", new_co2_over / 100)

# fraud_pct = (ship_sheet["B23"].value or 0) * 100
# new_fraud = st.sidebar.number_input("Average Fraud (Conservative) (%)", value=fraud_pct)
# if new_fraud != fraud_pct:
#     set_value("B23", new_fraud / 100)

# Average CO₂ Overage & Fraud -----------------------------------------------
co2_over_pct = (get_value("B21") or 0) * 100
new_co2_over = st.sidebar.number_input(
    "Average CO₂ Overage (%)", value=float(co2_over_pct), min_value=0.0
)
if new_co2_over != co2_over_pct:
    set_value("B21", new_co2_over / 100)

fraud_pct = (get_value("B23") or 0) * 100
new_fraud = st.sidebar.number_input(
    "Average Fraud (Conservative) (%)", value=float(fraud_pct), min_value=0.0
)
if new_fraud != fraud_pct:
    set_value("B23", new_fraud / 100)


# ----------------------------------------------------------------------------
# Live EUA price fetch -------------------------------------------------------
# ----------------------------------------------------------------------------

def get_live_eua_price():
    try:
        url = (
            "https://www.eex.com/en/market-data/environmental-markets/spot-market/"
            "european-emission-allowances"
        )
        headers = {"User-Agent": "Mozilla/5.0"}
        soup = BeautifulSoup(requests.get(url, headers=headers, timeout=10).text, "html.parser")
        # price_text = soup.find("td", text="2021-2030").find_next("td").text.strip()
        price_text = soup.find("td", string="2021-2030").find_next("td").text.strip()

        return float(price_text.replace(",", "."))
    except Exception:
        return None

live_price = get_live_eua_price()
def_price = live_price if live_price else ship_sheet["B26"].value or 0
new_price = st.sidebar.number_input("Current EUA Price (€)", value=float(def_price))
if new_price != def_price:
    set_value("B26", new_price)

# ----------------------------------------------------------------------------
# Results --------------------------------------------------------------------
# ----------------------------------------------------------------------------
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
    "Annex‑II CO₂ Emissions (2025 onward)": "E13",
    "Measured CO₂e Estimate": "E14",
    "Measured CO₂e Reduction": "E15",
    "SAVINGS € in 2025": "E16",
    "SAVINGS € in 2026": "E17",
    "SAVINGS € in 2027": "E18",
    "SAVINGS € in 2028": "E19",
    "Average Fraud Savings / yr": "E21",
}

with col1:
    for label, addr in metrics_col1.items():
        prefix = "€ " if "€" in label or "Liability" in label else ""
        safe_metric(label, get_value(addr), prefix)

with col2:
    for label, addr in metrics_col2.items():
        prefix = "€ " if "€" in label or "SAVINGS" in label else ""
        safe_metric(label, get_value(addr), prefix)

# ----------------------------------------------------------------------------
# Placeholder info (chart removed) -------------------------------------------
# ----------------------------------------------------------------------------
st.info("📝 Excel charts aren’t displayed in this cloud‑version. Replace with a Python chart if needed.")
