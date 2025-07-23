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
    if cell == "E6":  # Average Daily Fuel Use (MT) = (SEA Fuel * SEA DAYS + PORT
