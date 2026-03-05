import streamlit as st
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from osil_engine import run_osil, REQUIRED_COLUMNS
from report_generator import build_osil_pdf_report


st.set_page_config(page_title="OSIL™ by Xentrixus", layout="wide")

APP_TITLE = "OSIL™ by Xentrixus"
APP_SUB = "Operational Stability Intelligence"

DEMO_CSV_PATH = "data/demo_incidents.csv"


# -----------------------------
# Demo helper: tolerate imperfect demo data
# -----------------------------
def _find_column_case_insensitive(df: pd.DataFrame, target: str):
    norm_target = target.strip().lower()
    for c in df.columns:
        if str(c).strip().lower() == norm_target:
            return c
    return None


def _ensure_required_columns_for_demo(df: pd.DataFrame) -> pd.DataFrame:
    """
    Demo datasets are allowed to be imperfect.
    Auto-add missing REQUIRED_COLUMNS with safe defaults so demo always runs.
    """
    d = df.copy()

    # rename case-insensitive matches to the required names
    rename_map = {}
    for req in REQUIRED_COLUMNS:
        found = _find_column_case_insensitive(d, req)
        if found and found != req:
            rename_map[found] = req
    if rename_map:
        d = d.rename(columns=rename_map)

    # add missing required columns
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
                d[req] = pd.Timestamp("2026-01-01")
            else:
                d[req] = "Unknown"

    # keep required columns first
    d = d[REQUIRED_COLUMNS + [c for c in d.columns if c not in REQUIRED_COLUMNS]]
    return d


def _load_demo_csv() -> pd.DataFrame:
    df = pd.read_csv(DEMO_CSV_PATH)
    return _ensure_required_columns_for_demo(df)


# -----------------------------
# Executive Snapshot (Phase 7)
# -----------------------------
def build_operational_snapshot(results: dict) -> dict:
    overall = results.get("overall", {})
    posture = results.get("posture", "Unknown")
    bvsi = float(overall.get("BVSI", 0))

    # Use analyst review as the short narrative (already executive-toned)
    interpretation_html = results.get("analyst_review", "")

    # Build top risks and recommended actions from SIP table (preferred)
    sip_df = results.get("sip_table")
    top_risks = []
    actions = []

    if sip_df is not None and isinstance(sip_df, pd.DataFrame) and not sip_df.empty:
        # We will use the columns that actually exist in your SIP output
        # (Service, Service_Tier, Suggested_Theme, Why_Flagged, Priority_Label)
        for _, row in sip_df.head(3).iterrows():
            svc = str(row.get("Service", "Unknown Service"))
            tier = str(row.get("Service_Tier", ""))
            why = str(row.get("Why_Flagged", "Operational instability pattern"))
            theme = str(row.get("Suggested_Theme", ""))
            tier_txt = f" ({tier})" if tier else ""
            theme_txt = f" — {theme}" if theme and theme != "nan" else ""
            top_risks.append(f"{svc}{tier_txt}{theme_txt} — {why}")

        # Actions map directly to SIP initiation
        for _, row in sip_df.head(3).iterrows():
            svc = str(row.get("Service", "that service"))
            actions.append(f"Launch a Service Improvement Program (SIP) for {svc} with an accountable owner and 30-day outcomes.")

    else:
        # fallback if SIP is empty
        top_risks = ["No SIP candidates generated from this dataset yet."]
        actions = ["Run OSIL on a larger rolling window (e.g., 6–12 months) to strengthen signal quality."]

    return {
        "bvsi": bvsi,
        "posture": posture,
        "interpretation_html": interpretation_html,
        "top_risks": top_risks,
        "actions": actions,
    }


def _render_snapshot_box(snapshot: dict):
    st.subheader("Operational Stability Snapshot")

    st.markdown(
        f"""
<div style="border:1px solid #D1D5DB;background:#F5F7FA;padding:14px 16px;border-radius:10px;line-height:1.55;">
  <div style="font-size:18px;">
    <b>BVSI™:</b> {snapshot["bvsi"]:.1f}<br/>
    <b>Operating Posture:</b> {snapshot["posture"]}
  </div>
  <div style="margin-top:10px;">
    <b>What This Means</b><br/>
    {snapshot["interpretation_html"]}
  </div>
  <div style="margin-top:10px;">
    <b>Top Stability Risks</b>
    <ul style="margin-top:6px;">
      {''.join([f"<li>{r}</li>" for r in snapshot["top_risks"]])}
    </ul>
  </div>
  <div style="margin-top:10px;">
    <b>Recommended Actions</b>
    <ul style="margin-top:6px;">
      {''.join([f"<li>{a}</li>" for a in snapshot["actions"]])}
    </ul>
  </div>
</div>
""",
        unsafe_allow_html=True,
    )


# -----------------------------
# Heatmap
# -----------------------------
def _render_heatmap(service_risk_df: pd.DataFrame) -> None:
    st.subheader("Service Stability Heatmap (Top 10 Services by Risk)")
    st.caption("Executive view: Service × Stability Risk (Recurrence, MTTR Drag, Reopen Churn, Change Collision).")

    if service_risk_df is None or service_risk_df.empty:
        st.info("No service risk data available yet.")
        return

    metric_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_cols = ["Service", "Service_Tier"] + metric_cols + ["Total_Service_Risk"]

    show = service_risk_df.copy()[display_cols]

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


# -----------------------------
# Main App
# -----------------------------
def main():
    # Stable state (prevents rerun problems for downloads)
    if "osil_results" not in st.session_state:
        st.session_state.osil_results = None
    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
    if "pdf_filename" not in st.session_state:
        st.session_state.pdf_filename = "OSIL_Executive_Report_latest.pdf"

    st.title(APP_TITLE)
    st.caption(APP_SUB)

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

    if df is not None:
        try:
            results = run_osil(df)
            st.session_state.osil_results = results
            st.session_state.pdf_bytes = None  # reset report when new run occurs
        except Exception as e:
            st.error(f"Run failed: {e}")
            return

    if st.session_state.osil_results is None:
        st.info("Choose Demo Data or Upload a CSV to run OSIL.")
        return

    results = st.session_state.osil_results
    overall = results.get("overall", {})
    posture = results.get("posture", "")
    as_of = results.get("as_of", "")

    # Snapshot first (Signal → Meaning → Risk → Action)
    st.markdown("---")
    snapshot = build_operational_snapshot(results)
    _render_snapshot_box(snapshot)

    # KPI row
    st.markdown("---")
    c1, c2, c3 = st.columns([1, 2, 1])
    with c1:
        st.metric("BVSI™", f"{overall.get('BVSI', 0):.1f}" if "BVSI" in overall else "—")
    with c2:
        st.metric("Operating Posture", posture if posture else "—")
    with c3:
        st.metric("As-of Date", as_of if as_of else "—")

    # Radar
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

    # Heatmap
    st.markdown("---")
    _render_heatmap(results.get("service_risk_df"))

    # SIPs
    st.markdown("---")
    st.subheader("Top SIP Candidates (Next 30 Days)")
    sip_df = results.get("sip_table")
    if sip_df is None or sip_df.empty:
        st.info("No SIP candidates generated.")
    else:
        st.dataframe(sip_df, use_container_width=True)

    # PDF
    st.markdown("---")
    st.subheader("Executive Report (PDF)")

    colA, colB = st.columns([1, 2])

    with colA:
        if st.button("Generate / Refresh PDF"):
            try:
                pdf_bytes = build_osil_pdf_report(results)
                st.session_state.pdf_bytes = pdf_bytes
                st.session_state.pdf_filename = f"OSIL_Executive_Report_{as_of or 'latest'}.pdf"
                st.success("PDF generated. Use the download button on the right.")
            except Exception as e:
                st.error(f"Report generation failed: {e}")

    with colB:
        if st.session_state.pdf_bytes:
            st.download_button(
                label="⬇️ Download OSIL Executive Report (PDF)",
                data=st.session_state.pdf_bytes,
                file_name=st.session_state.pdf_filename,
                mime="application/pdf",
                key="download_pdf_btn",
            )
        else:
            st.info("Click **Generate / Refresh PDF** to create the report, then download it here.")


if __name__ == "__main__":
    main()
