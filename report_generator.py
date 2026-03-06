import io
from typing import Any, Dict

import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    PageBreak,
)


# ============================================================
# Helpers
# ============================================================
def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None:
            return default
        return float(val)
    except Exception:
        return default


def _safe_df(val: Any) -> pd.DataFrame:
    if isinstance(val, pd.DataFrame):
        return val.copy()
    return pd.DataFrame()


def _styles():
    styles = getSampleStyleSheet()

    styles.add(
        ParagraphStyle(
            name="SmallMuted",
            parent=styles["BodyText"],
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#555555"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="SectionHeader",
            parent=styles["Heading2"],
            fontSize=14,
            leading=16,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=8,
        )
    )

    styles.add(
        ParagraphStyle(
            name="Metric",
            parent=styles["BodyText"],
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#111111"),
        )
    )

    return styles


def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    labels = list(domain_scores.keys())
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]

    if not labels:
        labels = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
        values = [0, 0, 0, 0]

    import numpy as np

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]

    fig = plt.figure(figsize=(4.6, 4.3), dpi=180)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.12)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 100)

    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)

    ax.set_title("Operational Stability Profile", pad=18, fontsize=11)

    img_bytes = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img_bytes, format="png", bbox_inches="tight", transparent=False)
    plt.close(fig)
    img_bytes.seek(0)
    return img_bytes


def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> io.BytesIO | None:
    df = _safe_df(service_risk_top10)
    if df.empty:
        return None

    needed = [
        "Service",
        "Service_Tier",
        "Recurrence_Risk",
        "MTTR_Drag_Risk",
        "Reopen_Churn_Risk",
        "Change_Collision_Risk",
    ]
    missing = [c for c in needed if c not in df.columns]
    if missing:
        return None

    hm = df.head(10).copy()
    hm.index = hm["Service"].astype(str) + " (" + hm["Service_Tier"].astype(str) + ")"
    hm = hm[
        ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    ].rename(
        columns={
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision",
        }
    )

    hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)

    fig = plt.figure(figsize=(7.0, 4.4), dpi=180)
    ax = plt.gca()

    im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(hm.columns)))
    ax.set_xticklabels(list(hm.columns), fontsize=8)
    ax.set_yticks(range(len(hm.index)))
    ax.set_yticklabels(list(hm.index), fontsize=8)
    ax.set_title("Service Stability Heatmap (Top 10 Services)", fontsize=11)

    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            ax.text(j, i, f"{int(round(float(hm.iat[i, j]), 0))}", ha="center", va="center", fontsize=8)

    cbar = plt.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cbar.set_label("Risk Score", fontsize=8)
    cbar.ax.tick_params(labelsize=7)

    img_bytes = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img_bytes, format="png", bbox_inches="tight", transparent=False)
    plt.close(fig)
    img_bytes.seek(0)
    return img_bytes


def _make_table(df: pd.DataFrame, font_size: int = 8) -> Table:
    df = _safe_df(df)
    if df.empty:
        data = [["No data available"]]
        tbl = Table(data)
        tbl.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("FONTNAME", (0, 0), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ]
            )
        )
        return tbl

    data = [list(df.columns)] + [[str(x) for x in row] for row in df.values.tolist()]
    tbl = Table(data, repeatRows=1)

    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9EDF3")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.black),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), font_size),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return tbl


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#444444"))
    canvas.drawString(0.6 * inch, 0.35 * inch, "Operational Stability Intelligence (OSIL™) • BVSI™ • Structural Risk Debt™")
    canvas.drawRightString(7.9 * inch, 0.35 * inch, "© Xentrixus • Confidential")
    canvas.restoreState()


# ============================================================
# Public API expected by app.py
# ============================================================
def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    styles = _styles()
    out = io.BytesIO()
    doc = SimpleDocTemplate(
        out,
        pagesize=LETTER,
        rightMargin=0.55 * inch,
        leftMargin=0.55 * inch,
        topMargin=0.55 * inch,
        bottomMargin=0.55 * inch,
    )

    tenant_name = str(payload.get("tenant_name", "Default"))
    as_of = str(payload.get("as_of", ""))
    bvsi = _safe_float(payload.get("bvsi", 0.0))
    posture = str(payload.get("posture", "Unknown"))
    executive_interpretation = str(payload.get("executive_interpretation", ""))

    domain_scores = payload.get("domain_scores", {}) or {}
    service_risk_top10 = _safe_df(payload.get("service_risk_top10"))
    sip_candidates = _safe_df(payload.get("sip_candidates"))

    detected_dataset = str(payload.get("detected_dataset", "unknown")).upper()
    service_anchor_used = str(payload.get("service_anchor_used", "None"))
    data_readiness_score = _safe_float(payload.get("data_readiness_score", 0.0))

    story = []

    # --------------------------------------------------------
    # Page 1
    # --------------------------------------------------------
    story.append(Paragraph("OSIL™ by Xentrixus", styles["Title"]))
    story.append(Paragraph(f"Operational Stability Intelligence Report  |  As-of: {as_of}", styles["SmallMuted"]))
    story.append(Spacer(1, 10))

    summary_rows = [
        ["BVSI™ (Business Value Stability Index)", f"{bvsi:.1f}"],
        ["Operating Posture", posture],
        ["Detected Dataset", detected_dataset],
        ["Service Anchor Used", service_anchor_used],
        ["Data Readiness Score", f"{data_readiness_score:.1f}%"],
        ["Organization", tenant_name],
    ]
    summary_tbl = Table(summary_rows, colWidths=[3.4 * inch, 3.2 * inch])
    summary_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F7FA")),
                ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                ("FONTNAME", (0, 0), (0, -1), "Helvetica-Bold"),
                ("FONTNAME", (1, 0), (1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )
    story.append(summary_tbl)

    story.append(Spacer(1, 12))
    story.append(Paragraph("Executive Interpretation", styles["SectionHeader"]))
    story.append(Paragraph(executive_interpretation, styles["BodyText"]))

    # --------------------------------------------------------
    # Page 2
    # --------------------------------------------------------
    story.append(PageBreak())
    story.append(Paragraph("Operational Stability Profile & Diagnostics", styles["SectionHeader"]))

    radar_img = _build_radar_image(domain_scores)
    story.append(Image(radar_img, width=5.4 * inch, height=4.8 * inch))
    story.append(Spacer(1, 10))

    ds_df = pd.DataFrame(
        {
            "Domain": list(domain_scores.keys()),
            "Score (0–100)": [round(_safe_float(v), 1) for v in domain_scores.values()],
        }
    )
    story.append(Paragraph("Domain Scores (0–100)", styles["Heading3"]))
    story.append(_make_table(ds_df, font_size=9))

    # --------------------------------------------------------
    # Page 3
    # --------------------------------------------------------
    story.append(PageBreak())
    story.append(Paragraph("Service Improvement Priorities (SIPs)", styles["SectionHeader"]))
    story.append(
        Paragraph(
            "SIPs below are prioritized by Service + Theme and tier exposure. "
            "These represent the most actionable stability initiatives to improve BVSI™ and reduce avoidable disruption.",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 8))
    story.append(_make_table(sip_candidates.head(10), font_size=8))

    # --------------------------------------------------------
    # Page 4
    # --------------------------------------------------------
    story.append(PageBreak())
    story.append(Paragraph("Service Stability Heatmap (Top 10)", styles["SectionHeader"]))
    story.append(
        Paragraph(
            "This heatmap highlights where stability risk is concentrated by service across four executive-relevant risk drivers. "
            "Scores are normalized 0–100 (higher = more risk).",
            styles["BodyText"],
        )
    )
    story.append(Spacer(1, 8))

    hm_img = _build_heatmap_image(service_risk_top10)
    if hm_img is not None:
        story.append(Image(hm_img, width=6.6 * inch, height=4.2 * inch))
    else:
        story.append(Paragraph("No Service Risk data available.", styles["BodyText"]))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    out.seek(0)
    return out.getvalue()


# ============================================================
# Charts
# ============================================================
def radar_chart(domain_scores: dict):
    labels = list(domain_scores.keys())
    values = list(domain_scores.values())

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]

    fig = plt.figure(figsize=(4.3, 4.3), dpi=200)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.12)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
    ax.set_title("Operational Stability Profile (Radar)", pad=18, fontsize=11)

    plt.tight_layout()
    return fig


def heatmap_chart(heatmap_df: pd.DataFrame):
    safe = heatmap_df.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)
    values = safe.to_numpy()

    fig = plt.figure(figsize=(7.6, 4.6), dpi=200)
    ax = plt.gca()
    im = ax.imshow(np.nan_to_num(values, nan=0.0), aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(range(len(safe.columns)))
    ax.set_xticklabels(list(safe.columns), fontsize=8)
    ax.set_yticks(range(len(safe.index)))
    ax.set_yticklabels(list(safe.index), fontsize=8)
    ax.set_title("Service × Stability Risk (0–100)", fontsize=11)

    for i in range(values.shape[0]):
        for j in range(values.shape[1]):
            v = values[i, j]
            txt = "—" if (pd.isna(v) or np.isinf(v)) else f"{int(round(float(v), 0))}"
            ax.text(j, i, txt, ha="center", va="center", fontsize=8)

    cbar = plt.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cbar.set_label("Risk Score", fontsize=8)
    cbar.ax.tick_params(labelsize=7)

    plt.tight_layout()
    return fig


# ============================================================
# Narrative / Trend
# ============================================================
def render_service_instability_leaders(service_risk_df: pd.DataFrame) -> None:
    st.subheader("Service Instability Leaders (Top 5)")
    st.caption("Narrative view of the services currently driving the highest operational instability risk.")

    if service_risk_df is None or service_risk_df.empty:
        st.info("No service instability signals available.")
        return

    df = service_risk_df.sort_values("Total_Service_Risk", ascending=False).head(5).copy()

    for rank, (_, row) in enumerate(df.iterrows(), start=1):
        service = str(row.get("Service", "Unknown Service"))
        tier = str(row.get("Service_Tier", "Unknown Tier"))
        score = float(row.get("Total_Service_Risk", 0.0))

        risks = {
            "Recurrence": float(row.get("Recurrence_Risk", 0.0)),
            "MTTR Drag": float(row.get("MTTR_Drag_Risk", 0.0)),
            "Reopen Churn": float(row.get("Reopen_Churn_Risk", 0.0)),
            "Change Collision": float(row.get("Change_Collision_Risk", 0.0)),
        }
        primary = max(risks, key=risks.get)
        primary_score = risks.get(primary, 0.0)

        if primary == "Recurrence":
            meaning = "Recurring incidents suggest unresolved structural issues and repeat operational friction."
            action = "Start a SIP focused on recurrence elimination: clear problem statements, root cause pathways, and preventive controls."
        elif primary == "MTTR Drag":
            meaning = "Recovery times are longer than expected, indicating response coordination gaps, unclear ownership, or weak runbooks."
            action = "Start a SIP focused on recovery execution: playbooks, escalation pathways, and targeted automation to reduce recovery time."
        elif primary == "Reopen Churn":
            meaning = "High reopen rates suggest incomplete resolution or fixes that do not hold under operational load."
            action = "Start a SIP focused on fix quality: improve validation, tighten closure criteria, and drive problem investigations for repeat patterns."
        else:
            meaning = "Instability patterns frequently occur near change windows, suggesting governance gaps or insufficient pre-release validation."
            action = "Start a SIP focused on change governance: Tier-1 controls, stronger validation, and post-change monitoring."

        st.markdown(
            f"""
<div style="border:1px solid #D1D5DB;background:#F5F7FA;padding:14px 16px;border-radius:10px;margin-bottom:10px;">
  <div style="font-size:16px;"><b>#{rank} {service}</b></div>
  <div style="margin-top:6px;">
    <b>Tier:</b> {tier} &nbsp; | &nbsp; <b>Total Risk Score:</b> {score:.1f}
  </div>
  <div style="margin-top:10px;">
    <b>Primary Instability Driver:</b> {primary} ({primary_score:.0f}/100)<br/>
    {meaning}
  </div>
  <div style="margin-top:10px;">
    <b>Recommended Action:</b><br/>
    {action}
  </div>
</div>
""",
            unsafe_allow_html=True,
        )


def _momentum_arrow(series: pd.Series) -> str:
    s = series.dropna()
    if len(s) < 2:
        return "→"
    last = float(s.iloc[-1])
    prev = float(s.iloc[-2])
    if last >= prev + 1.0:
        return "↑"
    if last <= prev - 1.0:
        return "↓"
    return "→"


def _momentum_label(arrow: str) -> str:
    return {"↑": "Improving", "→": "Stable", "↓": "Declining"}.get(arrow, "Stable")


def _render_bvsi_trend(tenant_name: str, admin_enabled: bool):
    st.subheader("BVSI Trend (Saved Runs)")
    trend_df = _db_read_runs_for_tenant(tenant_name, limit=500)

    if trend_df.empty or trend_df["bvsi"].dropna().empty:
        st.info("No saved runs yet for this organization. Run OSIL once to begin trend tracking.")
        return

    view = trend_df.dropna(subset=["run_ts", "bvsi"]).tail(12).copy()
    arrow = _momentum_arrow(view["bvsi"])
    label = _momentum_label(arrow)

    c1, c2, c3 = st.columns([1, 1, 3])
    with c1:
        st.metric("Momentum", f"{arrow} {label}")
    with c2:
        st.metric("Latest BVSI™", f"{float(view['bvsi'].iloc[-1]):.1f}")
    with c3:
        st.write(f"Trend is filtered to **{tenant_name}** (SQLite-backed).")

    if admin_enabled:
        with st.expander("Admin Controls (Protected)"):
            st.warning(f"Reset permanently deletes saved run history for **{tenant_name}**. This impacts BVSI trends.")
            confirm = st.text_input("Type RESET to confirm", key="reset_confirm_text")
            if st.button(f"Clear History for {tenant_name}", key="reset_clear_btn"):
                if confirm.strip().upper() == "RESET":
                    _db_clear_runs_for_tenant(tenant_name)
                    st.success(f"History cleared for {tenant_name}.")
                    st.rerun()
                else:
                    st.error("Confirmation not valid. Type RESET exactly to proceed.")

    fig = plt.figure(figsize=(7, 3.2), dpi=140)
    ax = plt.gca()
    ax.plot(view["run_ts"], view["bvsi"], marker="o")
    ax.set_ylim(0, 100)
    ax.set_ylabel("BVSI™ (0–100)")
    ax.set_xlabel("Run Timestamp (UTC)")
    ax.set_title(f"BVSI™ Trend — Last 12 Runs ({tenant_name})")
    plt.xticks(rotation=30, ha="right")
    plt.tight_layout()

    tc1, tc2, tc3 = st.columns([1, 2, 1])
    with tc2:
        st.pyplot(fig)

    export = trend_df.copy()
    export["run_ts"] = export["run_ts"].astype(str)
    csv_bytes = export.to_csv(index=False).encode("utf-8")
    st.download_button(
        "Download BVSI Trend History (CSV)",
        data=csv_bytes,
        file_name=f"osil_bvsi_trend_history_{tenant_name}.csv".replace(" ", "_"),
        mime="text/csv",
    )


# ============================================================
# Main
# ============================================================
def main():
    _db_init_and_migrate()

    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
    if "pdf_filename" not in st.session_state:
        st.session_state.pdf_filename = "OSIL_Executive_Report_latest.pdf"
    if "analysis_results" not in st.session_state:
        st.session_state.analysis_results = None
    if "last_run_signature" not in st.session_state:
        st.session_state.last_run_signature = None

    admin_enabled = _admin_mode_panel()
    tenant_name = _tenant_selector()

    st.title(APP_TITLE)
    st.caption(APP_SUB)

    st.subheader("Run Options")
    mode = st.radio("Choose a run mode", ["Run with Demo Data", "Upload a CSV"], horizontal=True)

    df = None
    source_label = None
    run_requested = False
    run_signature = None

    if mode == "Run with Demo Data":
        if st.button("Run Demo Analysis", use_container_width=True):
            demo_path = DEMO_CSV_PATH if os.path.exists(DEMO_CSV_PATH) else "demo_incidents.csv"
            try:
                df = pd.read_csv(demo_path)
                source_label = f"Demo ({os.path.basename(demo_path)})"
                run_requested = True
                run_signature = f"demo::{demo_path}"
            except Exception as e:
                st.error(f"Demo load failed: {e}")
                return
    else:
        uploaded_file = st.file_uploader("Upload operational CSV export (any practice)", type=["csv"])
        if uploaded_file is not None:
            if st.button("Run Uploaded Analysis", use_container_width=True):
                try:
                    df = pd.read_csv(uploaded_file)
                    source_label = f"Upload ({uploaded_file.name})"
                    run_requested = True
                    run_signature = f"upload::{uploaded_file.name}::{uploaded_file.size}"
                except Exception as e:
                    st.error(f"Upload load failed: {e}")
                    return

    if run_requested:
        st.success(f"Loaded: {source_label}")

        practice_type = detect_practice_type(df)
        df, anchor_used = normalize_service_anchor(df)
        df = ensure_minimum_columns(df)
        readiness_score = calculate_data_readiness(df)

        open_col, close_col = infer_open_close_columns(df)
        df["__mttr_hours"] = compute_mttr_hours(df, open_col, close_col)

        roll = df.groupby("Service_Anchor", dropna=False).agg(
            recurrence=("Service_Anchor", "count"),
            reopen_rate=("Reopened_Flag", "mean"),
            change_collision_rate=("Change_Related_Flag", "mean"),
            mttr_hours=("__mttr_hours", "mean"),
            tier=("Service_Tier", lambda x: x.value_counts().index[0] if len(x) else "Tier 3"),
            category=("Category", lambda x: x.value_counts().index[0] if len(x) else "General"),
        ).reset_index()

        domain_scores = build_domain_scores(roll)
        bvsi = compute_bvsi(domain_scores)
        posture = operating_posture(bvsi)
        gap = dominant_gap(domain_scores)
        as_of = date.today().isoformat()
        exec_text = executive_interpretation(bvsi, posture, gap)

        service_risk_df = build_service_risk_df(roll)
        top10 = service_risk_df.head(10).copy()

        hm = top10.set_index(top10["Service"] + " (" + top10["Service_Tier"] + ")")[
            ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
        ].rename(
            columns={
                "Recurrence_Risk": "Recurrence",
                "MTTR_Drag_Risk": "MTTR Drag",
                "Reopen_Churn_Risk": "Reopen Churn",
                "Change_Collision_Risk": "Change Collision",
            }
        ).apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)

        sip = service_risk_df.copy()
        sip["SIP_Priority_Score"] = (
            0.45 * sip["Recurrence_Risk"] +
            0.25 * sip["MTTR_Drag_Risk"] +
            0.20 * sip["Reopen_Churn_Risk"] +
            0.10 * sip["Change_Collision_Risk"]
        ).round(1)

        def label_sip(score: float) -> str:
            if score >= 70:
                return "Next SIP"
            if score >= 45:
                return "Monitor"
            return "Backlog"

        sip["Priority_Label"] = sip["SIP_Priority_Score"].apply(label_sip)
        sip["Suggested_Theme"] = "Stability Improvement"
        sip["Why_Flagged"] = "Tier exposure + recurrence/MTTR/reopen/change risk"

        sip_view = sip.head(5)[
            ["Service", "Service_Tier", "Suggested_Theme", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"]
        ]

        results = {
            "source_label": source_label,
            "practice_type": practice_type,
            "anchor_used": anchor_used,
            "readiness_score": readiness_score,
            "bvsi": bvsi,
            "posture": posture,
            "gap": gap,
            "as_of": as_of,
            "exec_text": exec_text,
            "domain_scores": domain_scores,
            "service_risk_df": service_risk_df,
            "top10": top10,
            "heatmap_df": hm,
            "sip_view": sip_view,
            "tenant_name": tenant_name,
            "preview_df": df.head(20),
        }
        st.session_state.analysis_results = results

        if run_signature != st.session_state.last_run_signature:
            top_services = roll["Service_Anchor"].astype(str).head(3).tolist()
            _db_insert_run(
                tenant_name=tenant_name,
                run_ts=_utc_now_iso(),
                bvsi=bvsi,
                posture=posture,
                top_services=top_services,
            )
            st.session_state.last_run_signature = run_signature

    results = st.session_state.analysis_results

    if results is None:
        st.info("Choose a mode, provide a file if needed, and click the run button.")
        return

    # Diagnostics
    st.markdown("### Data Diagnostics")
    d1, d2, d3 = st.columns(3)
    d1.metric("Detected Dataset", results["practice_type"].upper())
    d2.metric("Service Anchor Used", results["anchor_used"])
    d3.metric("Data Readiness Score", f'{results["readiness_score"]:.1f}%')

    with st.expander("Preview (first 20 rows)"):
        st.dataframe(results["preview_df"], use_container_width=True)

    st.divider()

    # Snapshot
    st.subheader("Operational Stability Snapshot")
    st.caption(f'Organization: {results["tenant_name"]}')
    st.markdown(
        f"""
<div style="border:1px solid #D1D5DB;background:#F5F7FA;padding:14px 16px;border-radius:10px;line-height:1.55;">
  <div style="font-size:18px;">
    <b>BVSI™:</b> {results["bvsi"]:.1f}<br/>
    <b>Operating Posture:</b> {results["posture"]}
  </div>
  <div style="margin-top:10px;">
    <b>What This Means</b><br/>
    {results["exec_text"]}
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.divider()
    _render_bvsi_trend(tenant_name=results["tenant_name"], admin_enabled=admin_enabled)

    st.divider()
    m1, m2, m3 = st.columns([1, 2, 1])
    m1.metric("BVSI™", f'{results["bvsi"]:.1f}')
    m2.metric("Operating Posture", results["posture"])
    m3.metric("As-of Date", results["as_of"])

    st.divider()

    st.markdown("### Operational Stability Profile (Radar)")
    rad_fig = radar_chart(results["domain_scores"])
    rc1, rc2, rc3 = st.columns([1, 2, 1])
    with rc2:
        st.pyplot(rad_fig, use_container_width=False)

    st.markdown("### Stability Domain Scores (0–100)")
    st.dataframe(
        pd.DataFrame({"Domain": list(results["domain_scores"].keys()), "Score": list(results["domain_scores"].values())}),
        use_container_width=True,
    )

    st.divider()

    st.markdown("### Service Stability Heatmap (Top 10 Services by Risk)")
    hm_fig = heatmap_chart(results["heatmap_df"])
    hc1, hc2, hc3 = st.columns([1, 2, 1])
    with hc2:
        st.pyplot(hm_fig, use_container_width=False)

    st.markdown("**Top 10 Services — Risk Breakdown**")
    st.dataframe(results["top10"], use_container_width=True)

    risk_csv = results["top10"].to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Service Risk Table (CSV)",
        data=risk_csv,
        file_name=f'osil_service_risk_top10_{results["tenant_name"]}.csv'.replace(" ", "_"),
        mime="text/csv",
    )

    st.divider()
    render_service_instability_leaders(results["service_risk_df"])

    st.divider()

    st.markdown("### Top SIP Candidates (Next 30 Days)")
    st.dataframe(results["sip_view"], use_container_width=True)

    sip_csv = results["sip_view"].to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download SIP Candidates (CSV)",
        data=sip_csv,
        file_name=f'osil_sip_candidates_{results["tenant_name"]}.csv'.replace(" ", "_"),
        mime="text/csv",
    )

    st.divider()

    st.markdown("### Executive PDF Report")

    payload = {
        "bvsi": results["bvsi"],
        "posture": results["posture"],
        "as_of": results["as_of"],
        "executive_interpretation": results["exec_text"],
        "domain_scores": results["domain_scores"],
        "service_risk_top10": results["top10"],
        "sip_candidates": results["sip_view"],
        "data_readiness_score": results["readiness_score"],
        "service_anchor_used": results["anchor_used"],
        "detected_dataset": results["practice_type"],
        "tenant_name": results["tenant_name"],
    }

    colA, colB = st.columns([1, 2])

    with colA:
        if st.button("Generate / Refresh PDF", use_container_width=True):
            try:
                st.session_state.pdf_bytes = build_osil_pdf_report(payload)
                st.session_state.pdf_filename = (
                    f'OSIL_Executive_Report_{results["as_of"]}_{results["tenant_name"]}.pdf'.replace(" ", "_")
                )
                st.success("PDF generated. Use the download button on the right.")
            except Exception as e:
                st.error(f"Report generation failed: {e}")

    with colB:
        if st.session_state.pdf_bytes:
            st.download_button(
                label="⬇️ Download OSIL™ Executive Report (PDF)",
                data=st.session_state.pdf_bytes,
                file_name=st.session_state.pdf_filename,
                mime="application/pdf",
                key="download_pdf_btn",
                use_container_width=True,
            )
        else:
            st.info("Click **Generate / Refresh PDF** first, then download it here.")


if __name__ == "__main__":
    main()
