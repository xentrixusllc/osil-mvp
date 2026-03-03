import streamlit as st
import pandas as pd
import plotly.express as px
from osil_engine import run_osil

st.set_page_config(page_title="Xentrixus OSIL MVP", layout="wide")

st.title("Xentrixus OSIL — Stability Intelligence MVP")
st.caption("Upload ITSM exports → get BVSI, Structural Risk Debt™, SIP priorities, and executive interpretation.")

with st.expander("CSV Template (Required Columns)", expanded=False):
    st.code(
        "Service,Service_Tier,Opened_Date,Closed_Date,Priority,Reopened_Flag,Change_Related_Flag,Category\n"
        "Customer Portal,Tier 1,2026-01-05,2026-01-06,P2,0,1,Performance",
        language="text"
    )

uploaded = st.file_uploader("Upload your ITSM CSV export", type=["csv"])

if uploaded:
    df = pd.read_csv(uploaded)
    try:
        results = run_osil(df)
    except Exception as e:
        st.error(f"OSIL could not run: {e}")
        st.stop()

    bvsi = results["overall"]["BVSI"]
    posture = results["posture"]

    col1, col2, col3 = st.columns(3)
    col1.metric("BVSI", f"{bvsi:.1f}")
    col2.metric("Operating Posture", posture)
    col3.metric("As-of Date", results["as_of"])

    # Interpretation blocks
    st.subheader("Executive Interpretation")
    if bvsi < 40:
        st.write("Your organization operates in a reactive stability posture. Executive confidence and customer trust exposure may be materially constrained during disruption events.")
    elif bvsi < 60:
        st.write("Your organization maintains operational control, but structural instability remains active. Executive confidence may fluctuate under peak demand conditions.")
    elif bvsi < 80:
        st.write("Your organization demonstrates controlled stability. Governance supports growth, though eliminating recurring instability will strengthen executive confidence.")
    else:
        st.write("Your organization operates within a predictable stability envelope. Governance maturity supports scalable change and sustained customer trust.")

    st.divider()

    svc = results["service_table"].copy()

    # Radar (diamond) via long-form + polar
    st.subheader("Stability Domain Profile (Radar)")
    radar = svc[["Service","Service_Resilience","Change_Governance","Structural_Risk_Debt","Reliability_Momentum"]].copy()
    # show overall radar instead of per service
    overall = results["overall"]
    radar_overall = pd.DataFrame({
        "Domain": ["Service Resilience","Change Governance","Structural Risk Debt","Reliability Momentum"],
        "Score": [
            overall["Overall Service Resilience"],
            overall["Overall Change Governance"],
            overall["Overall Structural Risk Debt"],
            overall["Overall Reliability Momentum"],
        ]
    })
    fig_radar = px.line_polar(radar_overall, r="Score", theta="Domain", line_close=True)
    st.plotly_chart(fig_radar, use_container_width=True)

    st.subheader("Risk Concentration Heatmap (Service × Domain)")
    heat = svc.set_index(["Service","Service_Tier"])[["Service_Resilience","Change_Governance","Structural_Risk_Debt","Reliability_Momentum"]]
    heat_reset = heat.reset_index()
    heat_melt = heat_reset.melt(id_vars=["Service","Service_Tier"], var_name="Domain", value_name="Score")
    fig_heat = px.density_heatmap(
        heat_melt, x="Domain", y="Service", z="Score",
        color_continuous_scale="RdYlGn"
    )
    st.plotly_chart(fig_heat, use_container_width=True)

    st.subheader("Top SIP Candidates (Next 30 Days)")
    st.dataframe(results["sip_table"], use_container_width=True)

    with st.expander("Full Service Table (for analysts)", expanded=False):
        st.dataframe(svc, use_container_width=True)
else:
    st.info("Upload a CSV to run OSIL.")
