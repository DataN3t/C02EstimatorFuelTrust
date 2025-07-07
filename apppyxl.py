# --- Formula debug (temporary) ---------------------------------------------
st.subheader("🔎 Formula debug (temporary)")
check_cells = ["E6", "E7", "E11", "E13"]
for addr in check_cells:
    try:
        val = get_value(addr)
        st.write(addr, "=", val)
    except Exception as e:
        st.error(f"{addr} ➜ {e}")

# --- Raw evaluation test ----------------------------------------------------
st.subheader("⚙️ Raw Evaluation Test")
try:
    raw = ev.evaluate("'Ship Estimator'!B6")  # simple cell to probe
    st.write("B6 =", raw)
except Exception as e:
    st.error(f"Failed evaluating B6: {e}")

# ----------------------------------------------------------------------------
# Results --------------------------------------------------------------------
st.subheader("📊 Estimator Results")
