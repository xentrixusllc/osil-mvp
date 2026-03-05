import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from osil_engine import run_osil, REQUIRED_COLUMNS
from report_generator import build_osil_pdf_report


st.set_page_config(
    page_title="OSIL™ by Xentrixus",
    layout="wide",
)

APP_TITLE = "OSIL™ by Xentrixus"
APP_SUB = "Operational Stability Intelligence"


DEMO_CSV_PATH = "data/demo_incidents.csv"


# --------------------------------------------------
# EXECUTIVE SNAPSHOT GENERATOR
# --------------------------------------------------

def build_operational_snapshot(results):

    overall = results.get("overall", {})
    sip_df = results.get("sip_table")

    bvsi = overall.get("BVSI", 0)
    posture = results.get("posture", "Unknown")

    interpretation = results.get("analyst_review", "")

    top_risks = []
    if sip_df is not None and len(sip_df) > 0:
        for _, row in sip_df.head(3).iterrows():
            svc = row.get("Service", "Unknown Service")
            reason = row.get("Reason", "Operational instability")
            top_risks.append(f"{svc} — {reason}")

    actions = []
    if sip_df is not None and len(sip_df) > 0:
        for _, row in sip_df.head(3).iterrows():
            svc = row.get("Service", "Service")
            actions.append(f"Launch Service Improvement Program for {svc}")

    return {
        "bvsi": bvsi,
        "posture": posture,
        "interpretation": interpretation,
        "risks": top_risks,
        "actions": actions
    }


# --------------------------------------------------
# HEATMAP
# --------------------------------------------------

def render_heatmap(service_risk_df):

    if service_risk_df is None or service_risk_df.empty:
        return

    metric_cols = [
        "Recurrence_Risk",
        "MTTR_Drag_Risk",
        "Reopen_Churn_Risk",
        "Change_Collision_Risk"
    ]

    show = service_risk_df.head(10)

    services = show["Service"].tolist()
    tiers = show["Service_Tier"].tolist()
    matrix = show[metric_cols].to_numpy(dtype=float)

    fig = plt.figure(figsize=(10,5), dpi=160)
    ax = plt.gca()

    im = ax.imshow(matrix, aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(np.arange(len(metric_cols)))
    ax.set_xticklabels(
        ["Recurrence","MTTR Drag","Reopen Churn","Change Collision"]
    )

    ax.set_yticks(np.arange(len(services)))
    ax.set_yticklabels(
        [f"{s} ({t})" for s,t in zip(services,tiers)]
    )

    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            ax.text(j,i,f"{matrix[i,j]:.0f}",
                    ha="center",va="center",fontsize=9)

    plt.tight_layout()

    st.pyplot(fig)


# --------------------------------------------------
# APP
# --------------------------------------------------

def main():

    if "results" not in st.session_state:
        st.session_state.results = None

    if "pdf" not in st.session_state:
        st.session_state.pdf = None


    st.title(APP_TITLE)
    st.caption(APP_SUB)

    mode = st.radio(
        "Choose data source",
        ["Run Demo Data","Upload CSV"]
    )

    df = None

    if mode == "Run Demo Data":

        if st.button("Run Demo"):

            df = pd.read_csv(DEMO_CSV_PATH)

    else:

        uploaded = st.file_uploader("Upload CSV")

        if uploaded:
            df = pd.read_csv(uploaded)


    if df is not None:

        try:
            results = run_osil(df)
            st.session_state.results = results
        except Exception as e:
            st.error(e)


    results = st.session_state.results

    if results is None:
        return


    # --------------------------------------------------
    # SNAPSHOT
    # --------------------------------------------------

    snapshot = build_operational_snapshot(results)

    st.markdown("---")

    st.subheader("Operational Stability Snapshot")

    st.markdown(
        f"""
BVSI™: **{snapshot['bvsi']:.1f}**

Operating Posture: **{snapshot['posture']}**
"""
    )

    st.markdown("**What This Means**")

    st.markdown(snapshot["interpretation"], unsafe_allow_html=True)

    if snapshot["risks"]:

        st.markdown("**Top Stability Risks**")

        for r in snapshot["risks"]:
            st.markdown(f"• {r}")

    if snapshot["actions"]:

        st.markdown("**Recommended Actions**")

        for a in snapshot["actions"]:
            st.markdown(f"• {a}")


    # --------------------------------------------------
    # METRICS
    # --------------------------------------------------

    overall = results.get("overall", {})

    col1,col2,col3 = st.columns(3)

    col1.metric("BVSI™", f"{overall.get('BVSI',0):.1f}")
    col2.metric("Operating Posture", results.get("posture",""))
    col3.metric("As of", results.get("as_of",""))


    # --------------------------------------------------
    # RADAR
    # --------------------------------------------------

    st.markdown("---")

    st.subheader("Operational Stability Profile")

    labels = [
        "Service Resilience",
        "Change Governance",
        "Structural Risk Debt",
        "Reliability Momentum"
    ]

    values = [
        float(overall.get("Overall Service Resilience",0)),
        float(overall.get("Overall Change Governance",0)),
        float(overall.get("Overall Structural Risk Debt",0)),
        float(overall.get("Overall Reliability Momentum",0))
    ]

    angles = np.linspace(0,2*np.pi,len(labels),endpoint=False).tolist()
    values_loop = values + values[:1]
    angles_loop = angles + angles[:1]

    fig = plt.figure(figsize=(6,5),dpi=160)
    ax = plt.subplot(111,polar=True)

    ax.plot(angles_loop,values_loop,linewidth=2)
    ax.fill(angles_loop,values_loop,alpha=0.1)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels)

    ax.set_ylim(0,100)

    st.pyplot(fig)


    # --------------------------------------------------
    # HEATMAP
    # --------------------------------------------------

    st.markdown("---")

    st.subheader("Service Stability Heatmap")

    render_heatmap(results.get("service_risk_df"))


    # --------------------------------------------------
    # SIP TABLE
    # --------------------------------------------------

    st.markdown("---")

    st.subheader("Top SIP Candidates")

    sip_df = results.get("sip_table")

    if sip_df is not None:
        st.dataframe(sip_df)


    # --------------------------------------------------
    # PDF
    # --------------------------------------------------

    st.markdown("---")

    if st.button("Generate Executive Report PDF"):

        pdf = build_osil_pdf_report(results)

        st.session_state.pdf = pdf


    if st.session_state.pdf:

        st.download_button(
            "Download OSIL Executive Report",
            st.session_state.pdf,
            "OSIL_Executive_Report.pdf"
        )


if __name__ == "__main__":
    main()
