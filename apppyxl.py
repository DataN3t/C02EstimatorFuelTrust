# ----------------------------------------------------------------------------
# Load workbook + evaluator
# ----------------------------------------------------------------------------
wb, ev = load_model(EXCEL_PATH)

# Kick-start evaluation (force one cell to compute)
ev.evaluate("'Ship Estimator'!E6")  # any known formula cell will do

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

# ----------------------------------------------------------------------------
# DEBUG (optional: remove once working)
# ----------------------------------------------------------------------------
probe_cells = ["E6", "E7", "E11"]
for c in probe_cells:
    st.write(f"🔍 DEBUG {c} →", get_value(c))
