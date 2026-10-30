import streamlit as st

from landed_cost_engine import (
    BUCKETS,
    ParseError,
    compute_landed_cost,
    compute_net_avg_cost,
    load_worksheet,
    parse_obligation_file,
    parse_schedule_file,
)
from settings_store import is_streamlit_cloud, load_settings, save_settings

st.set_page_config(page_title="IEX Landed Cost per Unit", page_icon="\u26a1", layout="centered")

cloud_deployment = is_streamlit_cloud()
if cloud_deployment:
    saved = st.session_state.setdefault("app_settings", load_settings())
else:
    saved = load_settings()

st.title("SJCL \u2013 IEX Landed Cost per Unit")
st.caption(
    "Upload the Scheduling Request (SCH) file and the Daily Obligation "
    "Summary Report (DOR) file for a delivery day. This computes the "
    "Landed Cost per Unit \u2013 RTC / Peak / Normal / Off-Peak, the same way "
    "as the SJCL - IEX - Landed Cost workbook."
)

with st.sidebar:
    st.header("Settings")
    if cloud_deployment:
        st.caption("Settings are private to this browser session and reset when it ends.")
    else:
        st.caption("Saved automatically \u2014 these stay the same next time you open the app.")
    margin_paisa = st.number_input(
        "Trading margin (paisa / unit)",
        min_value=0.0,
        value=saved["margin_paisa"],
        step=0.5,
        format="%.2f",
        help="\u201cTwo (2) Paisa Trading Margin on Net Purchase of power\u201d in the "
        "reference sheet. Change this if your margin policy differs.",
    )
    css_rate = st.number_input(
        "CSS \u2013 Cross Subsidy Surcharge (Rs. / unit)",
        min_value=0.0,
        value=saved["css_rate"],
        step=0.01,
        format="%.2f",
        help="Used to compute Net Avg Cost. This tends to change periodically "
        "(e.g. per notification), so update it here when it changes.",
    )
    st.divider()
    st.markdown(
        "**Time-of-day buckets (fixed)**\n\n"
        "- Peak: 06:00\u201310:00 & 18:00\u201322:00\n"
        "- Normal: 15:00\u201318:00 & 22:00\u201324:00\n"
        "- Off-Peak: 00:00\u201306:00 & 10:00\u201315:00"
    )

col1, col2 = st.columns(2)
with col1:
    sch_file = st.file_uploader(
        "Scheduling Request file (SCH)",
        type=["xlsx", "xls"],
        help="The report with 'Time Period' rows and Injection/Drawal at "
        "Regional periphery / Interface point columns.",
    )
with col2:
    dor_file = st.file_uploader(
        "Daily Obligation Summary Report (DOR)",
        type=["xlsx", "xls"],
        help="The report with the Funds Payin / charges block and the "
        "Daily Trade Report (Period / Qty / Rate / Amount) table.",
    )

st.divider()
st.subheader("Andhra Pradesh Energy Cost")
st.caption(
    "Enter these yourself \u2013 they aren't read from any file. Saved "
    "automatically, so you only need to update them when the tariff changes "
    "(e.g. monthly/yearly), not every time you open the app."
)

ap_col1, ap_col2, ap_col3 = st.columns(3)
with ap_col1:
    ap_peak = st.number_input(
        "Peak (Rs./unit)", min_value=0.0, value=saved["ap_peak"], step=0.05, format="%.2f"
    )
with ap_col2:
    ap_normal = st.number_input(
        "Normal (Rs./unit)", min_value=0.0, value=saved["ap_normal"], step=0.05, format="%.2f"
    )
with ap_col3:
    ap_offpeak = st.number_input(
        "Off-Peak (Rs./unit)", min_value=0.0, value=saved["ap_offpeak"], step=0.05, format="%.2f"
    )

# Persist settings every time they change, so the app remembers them next
# time it's opened - no need to click a separate "Save" button.
current_settings = {
    "margin_paisa": margin_paisa,
    "css_rate": css_rate,
    "ap_peak": ap_peak,
    "ap_normal": ap_normal,
    "ap_offpeak": ap_offpeak,
}
if current_settings != {k: saved[k] for k in current_settings}:
    if cloud_deployment:
        st.session_state["app_settings"] = current_settings
    else:
        save_settings(current_settings)

st.divider()

run = st.button("Calculate Landed Cost", type="primary", use_container_width=True)

if run:
    if not sch_file or not dor_file:
        st.error("Please upload both the SCH and DOR files.")
        st.stop()

    try:
        with st.spinner("Reading scheduling file..."):
            sch_ws = load_worksheet(sch_file.getvalue(), sch_file.name)
            schedule = parse_schedule_file(sch_ws)

        with st.spinner("Reading obligation / charges file..."):
            dor_ws = load_worksheet(dor_file.getvalue(), dor_file.name)
            obligation = parse_obligation_file(dor_ws)

        results = compute_landed_cost(
            schedule, obligation, margin_rs_per_kwh=margin_paisa / 100.0
        )
    except ParseError as e:
        st.error(str(e))
        st.stop()
    except Exception as e:  # pragma: no cover - safety net for unexpected files
        st.error(f"Something went wrong while reading the files: {e}")
        st.stop()

    ap_energy_cost = {"Peak": ap_peak, "Normal": ap_normal, "OffPeak": ap_offpeak}
    net_avg = compute_net_avg_cost(results, ap_energy_cost, css_rate)

    st.success("Done.")

    order = ["RTC", "Peak", "Normal", "OffPeak"]
    labels = {"RTC": "RTC", "Peak": "Peak", "Normal": "Normal", "OffPeak": "Off-Peak"}

    main_col, side_col = st.columns([2, 1])

    with main_col:
        st.subheader("Landed Cost per Unit \u2013 RTC / PEAK / NORMAL / OFF-PEAK")
        cols = st.columns(4)
        for c, key in zip(cols, order):
            r = results[key]
            with c:
                st.metric(labels[key], f"\u20b9 {r.landed_cost_per_unit:.2f}")

    with side_col:
        rtc_net_avg = net_avg.get("RTC")
        if rtc_net_avg is not None:
            color = "#2ecc71" if rtc_net_avg >= 0 else "#e74c3c"
            st.markdown(
                f"""
                <div style="border:1px solid #333;border-radius:10px;
                            padding:16px 18px;text-align:center;">
                    <div style="font-size:1rem;opacity:0.75;margin-bottom:4px;">
                        Net Avg Cost
                    </div>
                    <div style="font-size:2.4rem;font-weight:800;color:{color};">
                        \u20b9 {rtc_net_avg:.2f}
                    </div>
                </div>
                """,
                unsafe_allow_html=True,
            )
        else:
            st.info("Enter Andhra Pradesh Energy Cost to see Net Avg Cost.")

    with st.expander("Calculation details (for verification)", expanded=False):
        import pandas as pd

        rows = []
        for key in order:
            r = results[key]
            rows.append(
                {
                    "Bucket": labels[key],
                    "Final Qty after losses, kWh (A)": round(r.final_qty_kwh, 2),
                    "Net Purchase Qty, kWh (C)": round(r.net_qty_kwh, 2),
                    "Net Purchase Amount, Rs. (C)": round(r.net_amount_rs, 2),
                    "NLDC/CTU/STU/ALDC charges, Rs. (D)": round(r.other_charges_rs, 2),
                    "Total Amount, Rs. (E = C+D)": round(r.total_amount_rs, 2),
                    "Trading Margin, Rs. (F)": round(r.trading_margin_rs, 2),
                    "Net Amount for Purchase, Rs. (G = E+F)": round(r.net_purchase_amount_rs, 2),
                    "Landed Cost / Unit, Rs. (H = G/A)": round(r.landed_cost_per_unit, 2),
                    "AP Energy Cost / Unit, Rs.": round(ap_energy_cost[key], 2) if ap_energy_cost.get(key) is not None else None,
                    "Net Avg Cost, Rs.": round(net_avg[key], 2) if net_avg.get(key) is not None else None,
                }
            )
        st.dataframe(pd.DataFrame(rows).set_index("Bucket").T, use_container_width=True)

        st.caption("Charge line-items read from the DOR file:")
        charge_rows = [(label, round(amount, 2)) for label, amount in obligation.charge_lines]
        st.dataframe(
            pd.DataFrame(charge_rows, columns=["Charge", "Amount (Rs.)"]),
            hide_index=True,
            use_container_width=True,
        )
