import streamlit as st
import pandas as pd
from pathlib import Path
import os
import plotly.express as px

# IMPORTANT: your repo has osil_engine.py (not osil_engine.py)
from osil_engine import run_osil

APP_TITLE = "Xentrixus OSIL — Stability Intelligence MVP"

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Upload ITSM exports → get BVSI, Structural Risk Debt™, SIP priorities, and executive interpretation.")

with st.expander("CSV Template (Required Columns)", expanded=False):
    st.code(
        "Service,Service_Tier,Opened_Date,Closed_Date,Priority,Reopened_Flag,Change_Related_Flag,Category\n"
        "Customer Portal,Tier 1,2026-01-05,2026-01-06,P2,0,1,Performance",
        language="text"
    )

with st.expander("Demo Data Diagnostics (temporary)", expanded=False):
    st.write("Working directory:", os.getcwd())
    st.write("Repo root files:", [p.name for p in Path(".").iterdir()])
    if Path("data").exists():
        st.write("./data files:", [p.name for p in Path("data").iterdir()])
    else:
        st.error("Folder './data' NOT found.")

def load_demo_single():
    """
    Your current engine expects ONE dataframe with Change_Related_Flag column.
    So we will load demo_incidents.csv only (it already contains the required columns).
    """
    p = Path("data") / "demo_incidents.csv"
    if not p.exists():
        raise FileNotFoundError(f"Missing {p}. Please ensure it exists in the repo.")
    return pd.read_csv(p)

def render_results(results: dict):
    bvsi = results["overall"]["BVSI"]
    posture = results["posture"]

    c1, c2, c3 = st.columns(3)
    c1.metric("BVSI", f"{bvsi:.1f}")
    c2.metric("Operating Posture", posture)
    c3.metric("As-of Date", results["as_of"])

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

    st.subheader("Stability Domain Profile (Radar)")
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

    st.subheader("Top SIP Candidates (Next 30 Days)")
    st.dataframe(results["sip_table"], use_container_width=True)

    with st.expander("Analyst View (Service Table)", expanded=False):
        st.dataframe(results["service_table"], use_container_width=True)

st.subheader("Run Options")

colA, colB = st.columns(2)
with colA:
    if st.button("Run with Demo Data"):
        try:
            demo_df = load_demo_single()
            results = run_osil(demo_df)
            render_results(results)
        except Exception as e:
            st.error(f"Demo run failed: {e}")

with colB:
    uploaded = st.file_uploader("Upload your ITSM CSV export", type=["csv"])
    if uploaded and st.button("Run OSIL with Uploaded CSV"):
        try:
            df = pd.read_csv(uploaded)
            results = run_osil(df)
            render_results(results)
        except Exception as e:
            st.error(f"Run failed: {e}")
