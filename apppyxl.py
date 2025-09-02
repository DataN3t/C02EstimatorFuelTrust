"""
Streamlit CO₂ Estimator – cloud-ready
• openpyxl  → loads workbook (.xlsx)
• xlcalculator → live recalculation when possible (fallback to static values)
"""

# ── Imports ────────────────────────────────────────────────────────────────
from pathlib import Path
import os, re, json, html
import streamlit as st
import streamlit.components.v1 as components
from openpyxl import load_workbook
from xlcalculator import ModelCompiler, Evaluator
import requests
from bs4 import BeautifulSoup
from decimal import Decimal, InvalidOperation
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
from typing import List, Dict, Optional

# ── Streamlit config ───────────────────────────────────────────────────────
st.set_page_config(page_title="FuelTrust CO₂ Ship Estimator", layout="wide")
st.title("Ship Estimator – Powered by FuelTrust")

# ── Excel workbook path ────────────────────────────────────────────────────
EXCEL_PATH = Path("CO2EmissionsEstimator3.xlsx")
if not EXCEL_PATH.exists():
    st.error(f"❌ {EXCEL_PATH} not found.")
    st.stop()

# ── Load workbook + evaluator (cached) ─────────────────────────────────────
@st.cache_resource(show_spinner="Loading C02 model…")
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

# ── Helpers ────────────────────────────────────────────────────────────────
xl_addr = lambda sheet, cell: f"'{sheet}'!{cell.upper()}"

def _flatten(val):
    while isinstance(val, list) and len(val) == 1:
        val = val[0]
    return val

def set_value(cell, value):
    ship_sheet[cell].value = value
    try:
        ev.set_cell_value(xl_addr("Ship Estimator", cell), value)
    except Exception:
        pass

def get_value(cell):
    try:
        val = ev.evaluate(xl_addr("Ship Estimator", cell))
        val = _flatten(val)
        if isinstance(val, (int, float)):
            return val
    except Exception:
        pass
    val = calculate_fallback(cell)
    if val is not None:
        set_value(cell, val)
        return val
    cached = ship_sheet[cell].value
    return cached if isinstance(cached, (int, float)) else None

def calculate_fallback(cell):
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
    if cell == "B18":
        nm_sea = safe_float("B7")
        nm_port = safe_float("B8")
        return (sea_days * nm_sea) + (port_days * nm_port)
    if cell == "E6":
        fuel_sea = safe_float("B10")
        fuel_port = safe_float("B11")
        return (fuel_sea * sea_days + fuel_port * port_days) / total_days
    if cell == "E7":
        fuel_type = ship_sheet["B19"].value
        cf_row = fuel_options.index(fuel_type) if fuel_type in fuel_options else 0
        cf = safe_float(f"B{43 + cf_row}", sheet=lookup_sheet)
        total_fuel = (safe_float("B10") * sea_days + safe_float("B11") * port_days)
        return total_fuel * cf
    if cell == "E8":
        co2_over = safe_float("B21")
        return (get_value("E7") or 0) * (1 - co2_over)
    if cell == "E9":
        return (get_value("E7") or 0) - (get_value("E8") or 0)
    if cell == "E10":
        eu_eu_pct = safe_float("B12")
        in_out_pct = safe_float("B13")
        return (get_value("E7") or 0) * (eu_eu_pct + in_out_pct * 0.5)
    if cell == "E11":
        eua_price = safe_float("B26")
        return (get_value("E10") or 0) * eua_price * 0.4
    if cell == "E12":
        eu_eu_pct = safe_float("B12")
        in_out_pct = safe_float("B13")
        return (get_value("E9") or 0) * (eu_eu_pct + in_out_pct * 0.5)
    if cell == "E13":
        return (get_value("E7") or 0) * 1.50419
    if cell == "E14":
        return (get_value("E13") or 0) * (1 - 0.0412)
    if cell == "E15":
        return (get_value("E13") or 0) - (get_value("E14") or 0)
    if cell in ["E16", "E17", "E18", "E19"]:
        year = {"E16": 2025, "E17": 2026, "E18": 2027, "E19": 2028}[cell]
        liability_pct = [0.4, 0.7, 1.0, 1.0][year - 2025]
        eua = safe_float("B26")
        return (get_value("E15") or 0) * eua * liability_pct
    if cell == "E21":
        fraud_pct = safe_float("B23")
        eua = safe_float("B26")
        return (get_value("E7") or 0) * fraud_pct * eua
    return None

def safe_metric(label, value, prefix=""):
    if isinstance(value, (int, float)):
        st.metric(label, f"{prefix}{value:,.2f}")
    else:
        st.metric(label, "–")

def get_range_values(range_name):
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

# ───────────────────────────────────────────────────────────────────────────
# EUA 3‑MONTH FORECAST (Vertis) — fetch + render
# ───────────────────────────────────────────────────────────────────────────
VERTIS_API_URL = "https://myvertis.com/mvapi/prices/"

# keep-it-simple token style
def get_vertis_token() -> str:
    return os.getenv("VERTIS_API_TOKEN", "95ddbff0db89ee2fc7561899847eec35561a8651")

EUA_3M_PATTERNS = [
    r"\beua\b.*\b3m\b",
    r"\beua-?3m\b",
    r"\beua\b.*\b3[-\s]?month(s)?\b",
    r"\beua\b.*\bq\+?1\b",
]
EUA_FALLBACK_PATTERNS = [r"^\s*eua\s*$", r"^\s*eua\b"]
CURRENCY_SYMBOLS = {"EUR": "€", "GBP": "£", "USD": "$"}

def fetch_vertis_prices(token: str, timeout: int = 15) -> List[Dict]:
    headers = {"Authorization": f"Token {token}"}
    params = {"format": "json"}
    resp = requests.get(VERTIS_API_URL, headers=headers, params=params, timeout=timeout)
    resp.raise_for_status()
    data = resp.json()
    if not isinstance(data, list):
        raise ValueError(f"Unexpected response shape: {type(data)}")
    return data

def _match_first(items: List[Dict], patterns: List[str]) -> Optional[Dict]:
    for pat in patterns:
        rx = re.compile(pat, flags=re.IGNORECASE)
        for it in items:
            name = str(it.get("product_name", "")).strip()
            if rx.search(name):
                return it
    return None

def pick_eua_3m_item(items: List[Dict]) -> Optional[Dict]:
    it = _match_first(items, EUA_3M_PATTERNS)
    return it or _match_first(items, EUA_FALLBACK_PATTERNS)

def _fmt_price(p, currency_code: str) -> str:
    try:
        d = Decimal(str(p))
        q = d.quantize(Decimal("0.01"))
    except (InvalidOperation, TypeError):
        return f"{p} {currency_code or ''}".strip()
    sym = CURRENCY_SYMBOLS.get(currency_code or "", "")
    return f"{sym}{q}" if sym else f"{q} {currency_code}".strip()

def build_eua_ticker_html(item: Optional[Dict], title: str = "EUA 3‑Month (Forward)") -> str:
    if not item:
        return """
        <div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial;
                    width: 100%; max-width: 520px; padding: 14px 16px; border: 1px solid #e5e7eb;
                    border-radius: 12px; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
          <div style="font-weight: 600; font-size: 15px; color: #111827;">EUA 3‑Month (Forward)</div>
          <div style="margin-top: 8px; font-size: 13px; color: #6b7280;">Not found in API response.</div>
        </div>
        """
    name = str(item.get("product_name", "EUA 3M")).strip()
    price = item.get("price")
    currency = str(item.get("currency", "")).strip()
    updated_at_raw = str(item.get("updated_at", "")).strip()
    pretty_time = updated_at_raw
    try:
        dt = datetime.fromisoformat(updated_at_raw)
        if dt.tzinfo is None:
            dt = dt.replace(tzinfo=timezone.utc)
        pretty_time = dt.astimezone(ZoneInfo("Europe/Berlin")).strftime("%d %B %Y %H:%M %Z")
    except Exception:
        pass
    price_str = _fmt_price(price, currency)
    return f"""
    <div style="font-family: system-ui, -apple-system, Segoe UI, Roboto, Arial;
                width: 100%; max-width: 520px; padding: 14px 16px; border: 1px solid #e5e7eb;
                border-radius: 12px; background: #fff; box-shadow: 0 1px 2px rgba(0,0,0,0.05);">
      <div style="display:flex; align-items:center; justify-content:space-between;">
        <div style="font-weight: 600; font-size: 15px; color: #111827;">{html.escape(title)}</div>
        <div style="font-size: 12px; color: #6b7280;">{html.escape(pretty_time)}</div>
      </div>
      <div style="margin-top: 6px; font-size: 28px; font-weight: 700; letter-spacing:-0.01em; color: #111827;">
        {html.escape(price_str)}
      </div>
      <div style="margin-top: 4px; font-size: 12px; color: #6b7280;">
        {html.escape(name)} • Currency: {html.escape(currency or '—')}
      </div>
    </div>
    """

# ── EEX spot (fallback) ─────────────────────────────────────────────────────
def get_live_eua_price():
    try:
        url = "https://www.eex.com/en/market-data/environmental-markets/spot-market/european-emission-allowances"
        soup = BeautifulSoup(requests.get(url, timeout=8).text, "html.parser")
        price_text = soup.find("td", string="2021-2030").find_next("td").text.strip()
        return float(price_text.replace(",", "."))
    except Exception:
        return None

# ── Auto-apply EUA price BEFORE showing the sidebar ─────────────────────────
autofill_source = None
vertis_item = None
ticker_price = None

token = get_vertis_token()
if token and token != "YOUR_VERTIS_TOKEN_HERE":
    try:
        prices = fetch_vertis_prices(token)
        vertis_item = pick_eua_3m_item(prices)
        if vertis_item and vertis_item.get("price") is not None:
            ticker_price = float(Decimal(str(vertis_item["price"])))
            set_value("B26", ticker_price)
            autofill_source = "Vertis 3‑Month"
    except Exception:
        pass

if ticker_price is None:
    lp = get_live_eua_price()
    if lp is not None:
        set_value("B26", float(lp))
        autofill_source = "EEX Spot"

if get_value("B26") is None:
    set_value("B26", 67.6)

# ── Sidebar – user inputs ──────────────────────────────────────────────────
st.sidebar.header("Adjust Estimator Inputs")

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
        dependents = {
            "B7": "SeaDayNM", "B8": "PortDayNM",
            "B10": "SeaDayMT", "B11": "PortDayMT",
            "B16": "SeaDays",  "B17": "PortDays",
        }
        ship_index = ship_type_list.index(selected_ship_type)
        for cell, range_name in dependents.items():
            values = get_range_values(range_name)
            if values and ship_index < len(values):
                new_val = values[ship_index]
                if isinstance(new_val, (int, float)):
                    set_value(cell, new_val)
        set_value("B18", calculate_fallback("B18"))

    # Numeric inputs — now with step=1.0
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
        new_val = st.number_input(label, value=float(default_val), step=1.0)
        if new_val != default_val:
            set_value(cell, new_val)
            if cell in ["B7", "B8", "B16", "B17"]:
                set_value("B18", calculate_fallback("B18"))

    # Percentage sliders (unchanged)
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
        "Default SEA Fuel", [*fuel_options],
        index=fuel_options.index(current_fuel) if current_fuel in fuel_options else 0,
    )
    if fuel_type != current_fuel:
        set_value("B19", fuel_type)

    # Avg CO₂ overage & fraud — explicit step=1
    co2_over_pct = int((get_value("B21") or 0) * 100)
    new_co2_over = st.number_input("Avg CO₂ Overage (%)", value=co2_over_pct, min_value=0, step=1)
    if new_co2_over != co2_over_pct:
        set_value("B21", new_co2_over / 100)

    fraud_pct = int((get_value("B23") or 0) * 100)
    new_fraud = st.number_input("Avg Fraud (%)", value=fraud_pct, min_value=0, step=1)
    if new_fraud != fraud_pct:
        set_value("B23", new_fraud / 100)

    # Current EUA Price (€) — explicit step=1.0 so +/- changes the integer part
    sidebar_default = float(get_value("B26") or 0.0)
    sidebar_price = st.number_input("Current EUA Price (€)", value=sidebar_default, step=1.0)
    if sidebar_price != get_value("B26"):
        set_value("B26", sidebar_price)

    if autofill_source:
        st.caption(f"Auto-filled from {autofill_source}")

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
    # Ticker card (no button)
    card_html = build_eua_ticker_html(vertis_item, title="EUA 3‑Month (Forward)")
    components.html(card_html, height=140)

with col2:
    for lbl, adr in metrics_col2.items():
        safe_metric(lbl, get_value(adr), "€ " if "€" in lbl else "")

# ── Bottom CTA ─────────────────────────────────────────────────────────────
def safe_html(s):
    return html.escape(str(s))

st.markdown("---")
co2e_reduction = get_value("E15") or 0.0
co2e_reduction_estimate = get_value("E14") or 0.0

st.markdown("""
<div style="text-align: center; background-color: #e6f7ff; padding: 20px; border-radius: 10px; border: 1px solid #91d5ff;">
<h2 style="color: #008000;">Estimated FuelTrust CO₂e Reduction Per Vessel Type: <strong>{co2e_reduction} MT</strong></h2>
<h2 style="color: #008000;">Measured CO2e Estimate For Floras CO2e-Offset Calculator Below: <strong>{co2e_reduction_estimate} MT</strong></h2>
<p style="font-size: 18px;">Unlock the exact decarb and in-depth insights tailored for your vessels.</p>
<a href="https://www.dk2advisor.com/get-in-touch" style="background-color: #1890ff; color: white; font-size: 16px; padding: 12px 24px; text-decoration: none; border-radius: 5px; font-weight: bold;">Get in Touch Today</a>
</div>
""".format(
    co2e_reduction=safe_html(f"{co2e_reduction:,.2f}"),
    co2e_reduction_estimate=safe_html(f"{co2e_reduction_estimate:,.2f}")
), unsafe_allow_html=True)
