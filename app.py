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

APP_TITLE = "Xentrixus OSIL™ — Stability Intelligence MVP"
APP_SUB = "Upload ITSM exports → get BVSI™, Structural Risk Debt™, SIP priorities, and executive interpretation."


DEMO_CSV_PATH = "data/demo_incidents.csv"  # demo file in your repo


def _required_template_text() -> str:
    cols = ",".join(REQUIRED_COLUMNS)
    example = "Customer Portal,Tier 1,2026-01-05,2026-01-06,P2,0,1,Performance"
    return f"{cols}\n{example}"


def _find_column_case_insensitive(df: pd.DataFrame, target: str):
    """Return actual column name if found ignoring case/spaces, else None."""
    norm_target = target.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == norm_target:
            return c
    return None


def _ensure_required_columns_for_demo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Demo datasets are allowed to be imperfect.
    We auto-add any missing REQUIRED_COLUMNS with safe defaults so the demo always runs.
    """
    d = df.copy()

    # If demo columns have wrong case, normalize them by renaming to required names
    rename_map = {}
    for req in REQUIRED_COLUMNS:
        found = _find_column_case_insensitive(d, req)
        if found and found != req:
            rename_map[found] = req
    if rename_map:
        d = d.rename(columns=rename_map)

    # Add missing required columns with safe defaults
    for req in REQUIRED_COLUMNS:
        if req not in d.columns:
            if req in ["Reopened_Flag", "Change_Related_Flag"]:
                d[req] = 0
            elif req == "Priority":
                d[req] = "P3"
            elif req == "Service_Tier":
                d[req] = "Tier 3"
            elif req == "Category":
                d[req] = "General"
            elif req in ["Opened_Date", "Closed_Date"]:
                # If dates are missing (rare), create plausible placeholders
                d[req] = pd.Timestamp("2026-01-01")
            else:
                d[req] = "Unknown"

    # Force correct ordering (not required, but nice)
    d = d[REQUIRED_COLUMNS + [c for c in d.columns if c not in REQUIRED_COLUMNS]]
    return d


def _load_demo_csv() -> pd.DataFrame:
    df = pd.read_csv(DEMO_CSV_PATH)
    df = _ensure_required_columns_for_demo(df)
    return df


def _render_heatmap(service_risk_df: pd.DataFrame) -> None:
    st.subheader("Service Stability Heatmap (Top 10 Services by Risk)")
    st.caption("Executive view: Service × Stability Risk (Recurrence, MTTR Drag, Reopen Churn, Change Collision).")

    if service_risk_df is None or service_risk_df.empty:
        st.info("No service risk data available yet.")
        return

    metric_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_cols = ["Service", "Service_Tier"] + metric_cols + ["Total_Service_Risk"]

    show = service_risk_df.copy()
    show = show[display_cols]

    services = show["Service"].tolist()
    tiers = show["Service_Tier"].tolist()
    matrix = show[metric_cols].to_numpy(dtype=float)

    fig = plt.figure(figsize=(10, 5), dpi=160)
    ax = plt.gca()

    im = ax.imshow(matrix, aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(np.arange(len(metric_cols)))
    ax.set_xticklabels(["Recurrence", "MTTR Drag", "Reopen Churn", "Change Collision"], rotation=0)

    ax.set_yticks(np.arange(len(services)))
    ax.set_yticklabels([f"{s} ({t})" for s, t in zip(services, tiers)])

    ax.set_title("Service × Stability Risk (0–100)")

    for i in range(matrix.shape[0]):
        for j in range(matrix.shape[1]):
            ax.text(j, i, f"{matrix[i, j]:.0f}", ha="center", va="center", fontsize=9)

    cbar = plt.colorbar(im, ax=ax, fraction=0.025, pad=0.02)
    cbar.set_label("Risk Score (0–100)")

    plt.tight_layout()
    st.pyplot(fig)

    st.markdown("**Top 10 Services — Risk Breakdown**")
    st.dataframe(show, use_container_width=True)

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

    if df is None:
        st.info("Choose Demo Data or Upload a CSV to run OSIL.")
        return

    # Run OSIL (strict validation remains in osil_engine for uploaded CSVs)
    try:
        results = run_osil(df)
    except Exception as e:
        st.error(f"Run failed: {e}")
        return

    overall = results.get("overall", {})
    posture = results.get("posture", "")
    as_of = results.get("as_of", "")

    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        st.metric("BVSI™", f"{overall.get('BVSI', 0):.1f}" if "BVSI" in overall else "—")
    with c2:
        st.metric("Operating Posture", posture if posture else "—")
    with c3:
        st.metric("As-of Date", as_of if as_of else "—")

    st.markdown("---")

    st.subheader("Operational Stability Profile (Radar)")

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

    st.markdown("---")
    _render_heatmap(results.get("service_risk_df"))

    st.markdown("---")
    st.subheader("Top SIP Candidates (Next 30 Days)")
    sip_df = results.get("sip_table")
    if sip_df is None or len(sip_df) == 0:
        st.info("No SIP candidates generated.")
    else:
        st.dataframe(sip_df, use_container_width=True)

    st.markdown("---")
    st.subheader("Executive Interpretation")
    st.write(results.get("analyst_review", ""))

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
