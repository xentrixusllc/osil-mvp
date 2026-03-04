import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from io import BytesIO

from osil_engine import run_osil, REQUIRED_COLUMNS
from report_generator import build_osil_pdf_report


st.set_page_config(
    page_title="OSIL™ by Xentrixus",
    layout="wide",
)


APP_TITLE = "Xentrixus OSIL™ — Stability Intelligence MVP"
APP_SUB = "Upload ITSM exports → get BVSI™, Structural Risk Debt™, SIP priorities, and executive interpretation."


def _load_demo_csv() -> pd.DataFrame:
    # Demo incidents data path in repo
    # If you later add other demo files, we can extend this.
    return pd.read_csv("data/demo_incidents.csv")


def _required_template_text() -> str:
    cols = ",".join(REQUIRED_COLUMNS)
    example = "Customer Portal,Tier 1,2026-01-05,2026-01-06,P2,0,1,Performance"
    return f"{cols}\n{example}"


def _render_heatmap(service_risk_df: pd.DataFrame) -> None:
    st.subheader("Service Stability Heatmap (Top 10 Services by Risk)")
    st.caption("Executive view: Service × Stability Risk (Recurrence, MTTR Drag, Reopen Churn, Change Collision).")

    if service_risk_df is None or service_risk_df.empty:
        st.info("No service risk data available yet.")
        return

    # Heatmap matrix
    metric_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_cols = ["Service", "Service_Tier"] + metric_cols + ["Total_Service_Risk"]

    show = service_risk_df.copy()
    show = show[display_cols]

    # Prepare matrix values
    services = show["Service"].tolist()
    tiers = show["Service_Tier"].tolist()

    matrix = show[metric_cols].to_numpy(dtype=float)

    # Plot
    fig = plt.figure(figsize=(10, 5), dpi=160)
    ax = plt.gca()

    # Use a built-in colormap; values 0-100
    im = ax.imshow(matrix, aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(np.arange(len(metric_cols)))
    ax.set_xticklabels(["Recurrence", "MTTR Drag", "Reopen Churn", "Change Collision"], rotation=0)

    ax.set_yticks(np.arange(len(services)))
    ax.set_yticklabels([f"{s} ({t})" for s, t in zip(services, tiers)])

    ax.set_title("Service × Stability Risk (0–100)")

    # Annotate values
    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            ax.text(j, i, f"{matrix[i, j]:.0f}", ha="center", va="center", fontsize=9)

    cbar = plt.colorbar(im, ax=ax, fraction=0.025, pad=0.02)
    cbar.set_label("Risk Score (0–100)")

    plt.tight_layout()
    st.pyplot(fig)

    # Show the table under the heatmap
    st.markdown("**Top 10 Services — Risk Breakdown**")
    st.dataframe(show, use_container_width=True)

    # Download CSV
    csv_bytes = show.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Service Risk Table (CSV)",
        data=csv_bytes,
        file_name="osil_service_risk_top10.csv",
        mime="text/csv",
    )


def main():
    st.title(APP_TITLE)
    st.write(APP_SUB)

    with st.expander("CSV Template (Required Columns)", expanded=False):
        st.code(_required_template_text())

    st.markdown("### Run Options")
    mode = st.radio("Choose a run mode", ["Run with Demo Data", "Upload a CSV"], horizontal=True)

    df = None

    if mode == "Run with Demo Data":
        if st.button("Run with Demo Data"):
            try:
                df = _load_demo_csv()
                st.success("Demo data loaded. Running OSIL…")
            except Exception as e:
                st.error(f"Failed to load demo data: {e}")
                return

    else:
        uploaded = st.file_uploader("Upload your ITSM CSV export", type=["csv"])
        if uploaded is not None:
            try:
                df = pd.read_csv(uploaded)
                st.success("Upload received. Running OSIL…")
            except Exception as e:
                st.error(f"Could not read CSV: {e}")
                return

    # Run only if df is present
    if df is None:
        st.info("Choose Demo Data or Upload a CSV to run OSIL.")
        return

    # Run OSIL
    try:
        results = run_osil(df)
    except Exception as e:
        st.error(f"Run failed: {e}")
        return

    overall = results.get("overall", {})
    posture = results.get("posture", "")
    as_of = results.get("as_of", "")

    # KPI Row
    c1, c2, c3 = st.columns([1, 2, 1])

    with c1:
        st.metric("BVSI™", f"{overall.get('BVSI', 0):.1f}" if "BVSI" in overall else "—")

    with c2:
        st.metric("Operating Posture", posture if posture else "—")

    with c3:
        st.metric("As-of Date", as_of if as_of else "—")

    st.markdown("---")

    # Radar + Domains
    st.subheader("Operational Stability Profile (Radar)")

    # Build a simple radar in-app as well (optional but useful)
    labels = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
    values = [
        float(overall.get("Overall Service Resilience", 0)),
        float(overall.get("Overall Change Governance", 0)),
        float(overall.get("Overall Structural Risk Debt", 0)),
        float(overall.get("Overall Reliability Momentum", 0)),
    ]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + values[:1]
    angles_loop = angles + angles[:1]

    fig = plt.figure(figsize=(6.8, 5), dpi=160)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=10)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=9)
    ax.set_ylim(0, 100)

    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.10)

    st.pyplot(fig)

    # NEW: Service Stability Heatmap (Option 1: Top 10)
    st.markdown("---")
    _render_heatmap(results.get("service_risk_df"))

    # SIP Candidates
    st.markdown("---")
    st.subheader("Top SIP Candidates (Next 30 Days)")
    sip_df = results.get("sip_table")
    if sip_df is None or len(sip_df) == 0:
        st.info("No SIP candidates generated.")
    else:
        st.dataframe(sip_df, use_container_width=True)

    # Analyst Review
    st.markdown("---")
    st.subheader("Executive Interpretation")
    st.write(results.get("analyst_review", ""))

    # Report Download
    st.markdown("---")
    st.subheader("Executive Report (PDF)")
    if st.button("Generate Executive Report (PDF)"):
        try:
            pdf_bytes = build_osil_pdf_report(results)
            st.download_button(
                label="Download OSIL Executive Report (PDF)",
                data=pdf_bytes,
                file_name=f"OSIL_Executive_Report_{as_of or 'latest'}.pdf",
                mime="application/pdf",
            )
            st.success("Report generated.")
        except Exception as e:
            st.error(f"Report generation failed: {e}")


if __name__ == "__main__":
    main()
