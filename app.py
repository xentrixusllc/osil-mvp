import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px

from osil_engine import run_osil

APP_TITLE = "OSIL by Xentrixus"

st.set_page_config(page_title=APP_TITLE, layout="wide")
st.title(APP_TITLE)
st.caption("Upload ITSM exports → get BVSI, Structural Risk Debt™, SIP priorities, and executive interpretation.")

# ---------------------------
# Helpers
# ---------------------------
def posture_from_bvsi(bvsi: float) -> str:
    if bvsi < 40:
        return "Reactive Instability"
    if bvsi < 60:
        return "Fragile Stability"
    if bvsi < 80:
        return "Controlled but Exposed"
    return "High Confidence Operations"

def posture_exec_meaning(posture: str) -> str:
    meanings = {
        "Reactive Instability": "Operational performance is unpredictable. Executive confidence and customer trust are exposed during disruption events.",
        "Fragile Stability": "Basic control exists, but recurring instability remains active. Confidence is vulnerable under peak demand or change periods.",
        "Controlled but Exposed": "Operations are stable enough to grow, but persistent structural drivers still create avoidable risk. Targeted SIPs will raise confidence and trust.",
        "High Confidence Operations": "Stability is predictable and governed. The organization can scale change while preserving service experience and trust.",
    }
    return meanings.get(posture, "")

def bvsi_legend_markdown(bvsi: float) -> str:
    # Simple indicator line to make the scale intuitive
    return f"""
**BVSI Scale (Business Value Stability Index)**  
- **80–100:** High Confidence Operations  
- **60–79:** Controlled but Exposed  
- **40–59:** Fragile Stability  
- **0–39:** Reactive Instability  

**Current BVSI:** **{bvsi:.1f}**
"""

def recommended_actions(overall: dict) -> list[str]:
    """
    Generates 3 exec-ready actions based on lowest-scoring domain.
    """
    domain_map = {
        "Overall Service Resilience": "Service Resilience",
        "Overall Change Governance": "Change Governance",
        "Overall Structural Risk Debt": "Structural Risk Debt™",
        "Overall Reliability Momentum": "Reliability Momentum",
    }

    # Identify weakest domain
    candidates = {k: overall.get(k, np.nan) for k in domain_map.keys()}
    candidates = {k: v for k, v in candidates.items() if pd.notna(v)}
    weakest_key = min(candidates, key=candidates.get) if candidates else "Overall Structural Risk Debt"
    weakest = domain_map.get(weakest_key, "Structural Risk Debt™")

    # Action sets
    if weakest == "Change Governance":
        return [
            "Launch a 30-day SIP on Tier 1 services with high change-induced instability (tighten risk checks and deployment discipline).",
            "Establish a focused change review for repeat offenders: enforce rollback readiness, test coverage, and owner accountability.",
            "Reduce recovery time by pre-defining runbooks and escalation paths for the highest-impact failure modes."
        ]
    if weakest == "Service Resilience":
        return [
            "Launch a Tier 1 resilience SIP: reduce MTTR drivers via runbooks, alert quality, and ownership clarity.",
            "Cut reopen churn by strengthening closure quality (verification steps, defect linkage, and prevention tasks).",
            "Stabilize the highest-volume failure theme (e.g., performance/login) with targeted engineering fixes and monitoring."
        ]
    if weakest == "Reliability Momentum":
        return [
            "Treat recurring instability as a managed backlog: build a weekly reliability cadence tied to Tier 1 outcomes.",
            "Prioritize prevention over speed: address the top recurrence drivers behind repeat incidents and reopens.",
            "Create visible stability commitments (top SIPs, owners, due dates) to restore executive confidence through predictability."
        ]
    # Default: Structural Risk Debt™
    return [
        "Launch SIPs against recurring instability clusters in Tier 1 services (reduce repeat incidents and reopen churn).",
        "Strengthen problem-to-prevention flow: convert repeat incident themes into corrective actions with owners and deadlines.",
        "Improve change discipline for high-risk services to reduce instability introduced by releases and configuration shifts."
    ]

# ---------------------------
# Data Safety Note
# ---------------------------
with st.expander("Data Safety (MVP)", expanded=False):
    st.write(
        "This MVP is designed for early demos and pilots. It processes data in-memory for scoring and visualization. "
        "Avoid uploading sensitive personal data. Use sanitized exports during pilots."
    )

# ---------------------------
# CSV Template
# ---------------------------
with st.expander("CSV Template (Required Columns)", expanded=False):
    st.markdown("Minimum required columns:")
    st.code(
        "Service,Service_Tier,Opened_Date,Closed_Date,Priority,Reopened_Flag,Category\n"
        "Customer Portal,Tier 1,2026-01-05 08:00,2026-01-05 12:30,P2,0,Performance",
        language="text"
    )
    st.markdown("Optional columns supported:")
    st.code(
        "Resolved_Date,Change_Related_Flag\n"
        "Resolved_Date is preferred for MTTR when present. Change_Related_Flag defaults to 0 if missing.",
        language="text"
    )

# ---------------------------
# Demo loader
# ---------------------------
def load_demo():
    # Uses your repo demo file
    return pd.read_csv("data/demo_incidents.csv")

# ---------------------------
# Render Results
# ---------------------------
def render_results(results: dict):
    overall = results["overall"]
    bvsi = float(overall["BVSI"])
    posture = results.get("posture") or posture_from_bvsi(bvsi)

    # Summary metrics
    c1, c2, c3 = st.columns(3)
    c1.metric("BVSI", f"{bvsi:.1f}")
    c2.metric("Operating Posture", posture)
    c3.metric("As-of Date", results.get("as_of", ""))

    # BVSI meaning (exec-facing)
    st.info(bvsi_legend_markdown(bvsi))
    st.write(f"**What this means for leadership:** {posture_exec_meaning(posture)}")

    # Executive Interpretation
    st.subheader("Executive Interpretation")
    st.write(
        "Your organization demonstrates operational stability measured across resilience, governance, structural risk, and momentum. "
        "The stability posture reflects how consistently services deliver outcomes with minimal disruption to customers and internal teams."
    )

    # Recommended next actions (auto)
    st.subheader("Recommended Next Actions (Next 30 Days)")
    for i, action in enumerate(recommended_actions(overall), start=1):
        st.write(f"{i}. {action}")

    st.divider()

    # Radar chart
    st.subheader("Operational Stability Profile (Radar)")
    radar_overall = pd.DataFrame({
        "Domain": ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"],
        "Score": [
            overall.get("Overall Service Resilience", 0),
            overall.get("Overall Change Governance", 0),
            overall.get("Overall Structural Risk Debt", 0),
            overall.get("Overall Reliability Momentum", 0),
        ]
    })
    fig_radar = px.line_polar(radar_overall, r="Score", theta="Domain", line_close=True)
    st.plotly_chart(fig_radar, use_container_width=True)
    st.caption("How to read: a balanced shape suggests aligned governance; a collapsed axis indicates a concentrated stability gap.")

    # SIP candidates
    st.subheader("Top SIP Candidates (Next 30 Days)")
    st.dataframe(results["sip_table"], use_container_width=True)

    # Analyst view
    with st.expander("Analyst View (Service Table)", expanded=False):
        st.dataframe(results["service_table"], use_container_width=True)

# ---------------------------
# Run Options (clear two-path UX)
# ---------------------------
st.subheader("Run Options")

left, right = st.columns(2)

with left:
    st.markdown("### Explore with Demo Data")
    st.caption("One click to see OSIL outputs using sample ITSM data.")
    if st.button("Run Demo Analysis"):
        try:
            demo_df = load_demo()
            results = run_osil(demo_df)
            render_results(results)
        except Exception as e:
            st.error(f"Demo run failed: {e}")

with right:
    st.markdown("### Analyze Your Data")
    st.caption("Upload a CSV export and generate your stability profile.")
    uploaded = st.file_uploader("Upload your ITSM CSV export", type=["csv"])
    if uploaded and st.button("Run OSIL with Uploaded CSV"):
        try:
            df = pd.read_csv(uploaded)
            results = run_osil(df)
            render_results(results)
        except Exception as e:
            st.error(f"Run failed: {e}")
