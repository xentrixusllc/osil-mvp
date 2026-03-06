import io
import os
import json
import sqlite3
from datetime import datetime, timezone, date

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

from data_classifier import (
    detect_practice_type,
    normalize_service_anchor,
    calculate_data_readiness,
)

# Optional external report generator
try:
    from report_generator import build_osil_pdf_report as external_build_osil_pdf_report
    EXTERNAL_REPORT_GENERATOR_AVAILABLE = True
except Exception:
    EXTERNAL_REPORT_GENERATOR_AVAILABLE = False

# Internal PDF fallback
try:
    from reportlab.lib.pagesizes import LETTER
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    REPORTLAB_AVAILABLE = True
except Exception:
    REPORTLAB_AVAILABLE = False


# ============================================================
# Config
# ============================================================
st.set_page_config(page_title="OSIL™ by Xentrixus", page_icon="📊", layout="wide")

APP_TITLE = "OSIL™ by Xentrixus"
APP_SUB = "Operational Stability Intelligence"

DEMO_CSV_PATH = "data/demo_incidents.csv"
SQLITE_PATH = "data/osil.db"


# ============================================================
# DB / Admin / Tenant
# ============================================================
def _safe_mkdir_for_file(path: str) -> None:
    folder = os.path.dirname(path)
    if folder and not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)


def _db_conn():
    _safe_mkdir_for_file(SQLITE_PATH)
    return sqlite3.connect(SQLITE_PATH, check_same_thread=False)


def _db_has_column(conn, table: str, col: str) -> bool:
    cur = conn.cursor()
    cur.execute(f"PRAGMA table_info({table})")
    cols = [r[1] for r in cur.fetchall()]
    return col in cols


def _db_init_and_migrate():
    conn = _db_conn()
    cur = conn.cursor()

    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS osil_runs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            tenant_name TEXT NOT NULL DEFAULT 'Default',
            run_ts TEXT NOT NULL,
            bvsi REAL NOT NULL,
            posture TEXT,
            top_services_json TEXT
        )
        """
    )
    conn.commit()

    if not _db_has_column(conn, "osil_runs", "tenant_name"):
        cur.execute("ALTER TABLE osil_runs ADD COLUMN tenant_name TEXT NOT NULL DEFAULT 'Default'")
        conn.commit()

    conn.close()


def _db_insert_run(tenant_name: str, run_ts: str, bvsi: float, posture: str, top_services: list):
    _db_init_and_migrate()
    tenant = (tenant_name or "Default").strip()

    conn = _db_conn()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO osil_runs (tenant_name, run_ts, bvsi, posture, top_services_json)
        VALUES (?, ?, ?, ?, ?)
        """,
        (tenant, run_ts, float(bvsi), posture or "", json.dumps(top_services or [])),
    )
    conn.commit()
    conn.close()


def _db_list_tenants(limit: int = 200) -> list:
    _db_init_and_migrate()
    conn = _db_conn()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT DISTINCT tenant_name
        FROM osil_runs
        WHERE tenant_name IS NOT NULL AND TRIM(tenant_name) <> ''
        ORDER BY tenant_name ASC
        LIMIT ?
        """,
        (limit,),
    )
    rows = cur.fetchall()
    conn.close()
    return [r[0] for r in rows] if rows else []


def _db_read_runs_for_tenant(tenant_name: str, limit: int = 500) -> pd.DataFrame:
    _db_init_and_migrate()
    tenant = (tenant_name or "Default").strip()

    conn = _db_conn()
    df = pd.read_sql_query(
        """
        SELECT tenant_name, run_ts, bvsi, posture, top_services_json
        FROM osil_runs
        WHERE tenant_name = ?
        ORDER BY datetime(run_ts) ASC
        LIMIT ?
        """,
        conn,
        params=(tenant, limit),
    )
    conn.close()

    if df.empty:
        return pd.DataFrame(columns=["tenant_name", "run_ts", "bvsi", "posture", "top_services"])

    df["run_ts"] = pd.to_datetime(df["run_ts"], errors="coerce")
    df["bvsi"] = pd.to_numeric(df["bvsi"], errors="coerce")
    df["top_services"] = df["top_services_json"].apply(
        lambda x: json.loads(x) if isinstance(x, str) and x.strip() else []
    )
    df = df.drop(columns=["top_services_json"], errors="ignore")
    df = df.sort_values("run_ts")
    return df


def _db_clear_runs_for_tenant(tenant_name: str):
    _db_init_and_migrate()
    tenant = (tenant_name or "Default").strip()

    conn = _db_conn()
    cur = conn.cursor()
    cur.execute("DELETE FROM osil_runs WHERE tenant_name = ?", (tenant,))
    conn.commit()
    conn.close()


def _utc_now_iso() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat()


def _get_admin_passcode() -> str:
    try:
        return str(st.secrets.get("OSIL_ADMIN_PASSCODE", "")).strip()
    except Exception:
        return ""


def _admin_mode_panel() -> bool:
    st.sidebar.markdown("### Admin Mode")

    if "admin_ok" not in st.session_state:
        st.session_state.admin_ok = False

    admin_pass = _get_admin_passcode()
    if not admin_pass:
        st.sidebar.info("Admin passcode not configured in Secrets.")
        return False

    with st.sidebar.expander("Unlock Admin Mode"):
        entered = st.text_input("Admin Passcode", type="password", key="admin_pass_input")
        if st.button("Unlock", key="admin_unlock_btn"):
            if entered.strip() == admin_pass:
                st.session_state.admin_ok = True
                st.success("Admin Mode enabled for this session.")
            else:
                st.session_state.admin_ok = False
                st.error("Invalid passcode.")

    return bool(st.session_state.admin_ok)


def _tenant_selector() -> str:
    st.sidebar.markdown("### Organization")
    tenants = _db_list_tenants()

    if "tenant_name" not in st.session_state:
        st.session_state.tenant_name = tenants[0] if tenants else "Default"

    mode = st.sidebar.radio("Choose", ["Select existing", "Enter new"], horizontal=False, key="tenant_mode")

    if mode == "Select existing":
        options = tenants if tenants else ["Default"]
        idx = options.index(st.session_state.tenant_name) if st.session_state.tenant_name in options else 0
        sel = st.sidebar.selectbox("Organization", options=options, index=idx, key="tenant_select")
        st.session_state.tenant_name = sel
    else:
        typed = st.sidebar.text_input("Organization Name", value=st.session_state.tenant_name, key="tenant_input")
        st.session_state.tenant_name = typed.strip() if typed.strip() else "Default"

    return st.session_state.tenant_name


# ============================================================
# Data helpers
# ============================================================
def _safe_parse_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")


def ensure_minimum_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()

    defaults = {
        "Service_Tier": "Tier 3",
        "Priority": "P3",
        "Reopened_Flag": 0,
        "Change_Related_Flag": 0,
        "Category": "General",
    }
    for col, default_val in defaults.items():
        if col not in out.columns:
            out[col] = default_val

    out["Service_Tier"] = out["Service_Tier"].astype(str).str.strip().replace({"": "Tier 3", "nan": "Tier 3"})
    for flag_col in ["Reopened_Flag", "Change_Related_Flag"]:
        out[flag_col] = pd.to_numeric(out[flag_col], errors="coerce").fillna(0).astype(int)

    return out


def infer_open_close_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    open_candidates = [
        "Opened_Date", "Opened_At", "opened_at", "open_date",
        "created_at", "Created", "Created_At", "Opened"
    ]
    close_candidates = [
        "Closed_Date", "Closed_At", "resolved_at", "Resolved_At",
        "close_date", "closed_at", "Resolved", "Closed", "Updated"
    ]

    open_col = next((c for c in open_candidates if c in df.columns), None)
    close_col = next((c for c in close_candidates if c in df.columns), None)
    return open_col, close_col


def compute_mttr_hours(df: pd.DataFrame, open_col: str | None, close_col: str | None) -> pd.Series:
    if not open_col or not close_col:
        return pd.Series([np.nan] * len(df), index=df.index)
    opened = _safe_parse_datetime(df[open_col])
    closed = _safe_parse_datetime(df[close_col])
    return (closed - opened).dt.total_seconds() / 3600.0


def normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    if len(s) == 0 or float(s.max()) == float(s.min()):
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - s.min()) / (s.max() - s.min()) * 100.0).clip(0, 100)


def build_domain_scores(service_rollup: pd.DataFrame) -> dict:
    if service_rollup is None or service_rollup.empty:
        return {
            "Service Resilience": 0.0,
            "Change Governance": 0.0,
            "Structural Risk Debt™": 0.0,
            "Reliability Momentum": 0.0,
        }

    rec = normalize_0_100(service_rollup["recurrence"])
    mttr = normalize_0_100(service_rollup["mttr_hours"])
    reopen = normalize_0_100(service_rollup["reopen_rate"])
    chg = normalize_0_100(service_rollup["change_collision_rate"])

    service_resilience = float((100 - (0.55 * rec + 0.45 * mttr)).clip(0, 100).mean())
    change_governance = float((100 - chg).clip(0, 100).mean())
    structural_risk_debt = float((100 - (0.65 * rec + 0.35 * reopen)).clip(0, 100).mean())
    reliability_momentum = float((100 - (0.50 * rec + 0.25 * reopen + 0.25 * mttr)).clip(0, 100).mean())

    return {
        "Service Resilience": round(service_resilience, 1),
        "Change Governance": round(change_governance, 1),
        "Structural Risk Debt™": round(structural_risk_debt, 1),
        "Reliability Momentum": round(reliability_momentum, 1),
    }


def compute_bvsi(domain_scores: dict) -> float:
    weights = {
        "Service Resilience": 0.30,
        "Change Governance": 0.25,
        "Structural Risk Debt™": 0.25,
        "Reliability Momentum": 0.20,
    }
    return round(sum(float(domain_scores.get(k, 0.0)) * w for k, w in weights.items()), 1)


def operating_posture(bvsi: float) -> str:
    if bvsi >= 85:
        return "High Confidence Operations"
    if bvsi >= 70:
        return "Controlled and Improving"
    if bvsi >= 55:
        return "Controlled but Exposed"
    if bvsi >= 40:
        return "Reactive and Exposed"
    return "Fragile Operations"


def dominant_gap(domain_scores: dict) -> str:
    if not domain_scores:
        return "Stability signal coverage"
    return min(domain_scores.items(), key=lambda x: x[1])[0]


def executive_interpretation(bvsi: float, posture: str, gap: str) -> str:
    if bvsi >= 85:
        return (
            f"Your organization demonstrates a **{posture}** operating posture (BVSI™ {bvsi:.1f}). "
            f"Stability signals are balanced overall, and targeted prevention work can further strengthen executive confidence and preserve customer trust. "
            f"Primary focus: **{gap}**."
        )
    if bvsi >= 70:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Control exists, but improvement is uneven across stability domains. Prioritize SIPs where Tier-1 exposure and recurrence patterns concentrate. "
            f"Primary focus: **{gap}**."
        )
    if bvsi >= 55:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Operational control exists, but recurring instability patterns still create exposure—especially across higher-impact services. "
            f"Focused SIP execution over the next 30–60 days will increase executive confidence and protect customer trust. "
            f"Primary focus: **{gap}**."
        )
    if bvsi >= 40:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Instability is material and likely visible to business stakeholders. Immediate stabilization actions and tighter change controls are recommended. "
            f"Primary focus: **{gap}**."
        )
    return (
        f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
        f"Operational risk is elevated and may threaten reliability commitments. Stabilization should be treated as an executive priority. "
        f"Primary focus: **{gap}**."
    )


def build_service_risk_df(roll: pd.DataFrame) -> pd.DataFrame:
    rec = normalize_0_100(roll["recurrence"])
    mttr = normalize_0_100(roll["mttr_hours"])
    reopen = normalize_0_100(roll["reopen_rate"])
    chg = normalize_0_100(roll["change_collision_rate"])

    df = pd.DataFrame({
        "Service": roll["Service_Anchor"].astype(str),
        "Service_Tier": roll["tier"].astype(str),
        "Recurrence_Risk": rec.round(1),
        "MTTR_Drag_Risk": mttr.round(1),
        "Reopen_Churn_Risk": reopen.round(1),
        "Change_Collision_Risk": chg.round(1),
    })

    df["Total_Service_Risk"] = (
        0.35 * df["Recurrence_Risk"] +
        0.25 * df["MTTR_Drag_Risk"] +
        0.25 * df["Reopen_Churn_Risk"] +
        0.15 * df["Change_Collision_Risk"]
    ).round(1)

    return df.sort_values("Total_Service_Risk", ascending=False).reset_index(drop=True)


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
# Narrative cards / trend
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
# PDF generation
# ============================================================
def _internal_build_pdf(payload: dict) -> bytes:
    if not REPORTLAB_AVAILABLE:
        raise RuntimeError("ReportLab is not available for internal PDF generation.")

    out = io.BytesIO()
    doc = SimpleDocTemplate(out, pagesize=LETTER)
    styles = getSampleStyleSheet()
    story = []

    title = Paragraph("OSIL™ Executive Report", styles["Title"])
    story.append(title)
    story.append(Spacer(1, 12))

    summary_lines = [
        f"Organization: {payload.get('tenant_name', 'Default')}",
        f"As-of: {payload.get('as_of', '')}",
        f"BVSI™: {payload.get('bvsi', 0)}",
        f"Operating Posture: {payload.get('posture', '')}",
        f"Detected Dataset: {payload.get('detected_dataset', '')}",
        f"Service Anchor Used: {payload.get('service_anchor_used', '')}",
        f"Data Readiness Score: {payload.get('data_readiness_score', 0)}%",
    ]
    for line in summary_lines:
        story.append(Paragraph(line, styles["BodyText"]))

    story.append(Spacer(1, 12))
    story.append(Paragraph("Executive Interpretation", styles["Heading2"]))
    story.append(Paragraph(payload.get("executive_interpretation", ""), styles["BodyText"]))

    story.append(Spacer(1, 12))
    story.append(Paragraph("Domain Scores", styles["Heading2"]))
    ds = payload.get("domain_scores", {})
    ds_rows = [["Domain", "Score"]]
    for k, v in ds.items():
        ds_rows.append([str(k), str(v)])
    ds_tbl = Table(ds_rows)
    ds_tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
        ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
    ]))
    story.append(ds_tbl)

    story.append(Spacer(1, 12))
    story.append(Paragraph("Top Service Risks", styles["Heading2"]))
    top10 = payload.get("service_risk_top10")
    if isinstance(top10, pd.DataFrame) and not top10.empty:
        rows = [list(top10.columns)]
        for _, r in top10.head(10).iterrows():
            rows.append([str(x) for x in r.tolist()])
        tbl = Table(rows, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
        ]))
        story.append(tbl)

    story.append(PageBreak())
    story.append(Paragraph("Top SIP Candidates", styles["Heading2"]))
    sip = payload.get("sip_candidates")
    if isinstance(sip, pd.DataFrame) and not sip.empty:
        rows = [list(sip.columns)]
        for _, r in sip.head(10).iterrows():
            rows.append([str(x) for x in r.tolist()])
        tbl = Table(rows, repeatRows=1)
        tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), colors.lightgrey),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
            ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
            ("FONTSIZE", (0, 0), (-1, -1), 7),
        ]))
        story.append(tbl)

    doc.build(story)
    out.seek(0)
    return out.getvalue()


def generate_pdf_bytes(payload: dict) -> bytes:
    # Prefer external generator if available
    if EXTERNAL_REPORT_GENERATOR_AVAILABLE:
        try:
            pdf_obj = external_build_osil_pdf_report(payload)
            if isinstance(pdf_obj, (bytes, bytearray)):
                return bytes(pdf_obj)
            if isinstance(pdf_obj, io.BytesIO):
                return pdf_obj.getvalue()
            if isinstance(pdf_obj, str) and os.path.exists(pdf_obj):
                with open(pdf_obj, "rb") as f:
                    return f.read()
        except Exception:
            pass

    # Fallback internal generator
    return _internal_build_pdf(payload)


# ============================================================
# Main
# ============================================================
def main():
    _db_init_and_migrate()

    if "pdf_bytes" not in st.session_state:
        st.session_state.pdf_bytes = None
    if "pdf_filename" not in st.session_state:
        st.session_state.pdf_filename = "OSIL_Executive_Report_latest.pdf"
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

    if not run_requested:
        st.info("Choose a mode, provide a file if needed, and click the run button.")
        return

    st.success(f"Loaded: {source_label}")

    # Diagnostics
    practice_type = detect_practice_type(df)
    df, anchor_used = normalize_service_anchor(df)
    df = ensure_minimum_columns(df)
    readiness_score = calculate_data_readiness(df)

    st.markdown("### Data Diagnostics")
    d1, d2, d3 = st.columns(3)
    d1.metric("Detected Dataset", practice_type.upper())
    d2.metric("Service Anchor Used", anchor_used)
    d3.metric("Data Readiness Score", f"{readiness_score:.1f}%")

    with st.expander("Preview (first 20 rows)"):
        st.dataframe(df.head(20), use_container_width=True)

    # Analysis
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

    # Save history only once per explicit run
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

    st.divider()

    # Executive snapshot
    st.subheader("Operational Stability Snapshot")
    st.caption(f"Organization: {tenant_name}")
    st.markdown(
        f"""
<div style="border:1px solid #D1D5DB;background:#F5F7FA;padding:14px 16px;border-radius:10px;line-height:1.55;">
  <div style="font-size:18px;">
    <b>BVSI™:</b> {bvsi:.1f}<br/>
    <b>Operating Posture:</b> {posture}
  </div>
  <div style="margin-top:10px;">
    <b>What This Means</b><br/>
    {exec_text}
  </div>
</div>
""",
        unsafe_allow_html=True,
    )

    st.divider()
    _render_bvsi_trend(tenant_name=tenant_name, admin_enabled=admin_enabled)

    st.divider()
    m1, m2, m3 = st.columns([1, 2, 1])
    m1.metric("BVSI™", f"{bvsi:.1f}")
    m2.metric("Operating Posture", posture)
    m3.metric("As-of Date", as_of)

    st.divider()

    st.markdown("### Operational Stability Profile (Radar)")
    rad_fig = radar_chart(domain_scores)
    rc1, rc2, rc3 = st.columns([1, 2, 1])
    with rc2:
        st.pyplot(rad_fig, use_container_width=False)

    st.markdown("### Stability Domain Scores (0–100)")
    st.dataframe(
        pd.DataFrame({"Domain": list(domain_scores.keys()), "Score": list(domain_scores.values())}),
        use_container_width=True
    )

    st.divider()

    service_risk_df = build_service_risk_df(roll)
    top10 = service_risk_df.head(10).copy()

    hm = top10.set_index(top10["Service"] + " (" + top10["Service_Tier"] + ")")[
        ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    ].rename(columns={
        "Recurrence_Risk": "Recurrence",
        "MTTR_Drag_Risk": "MTTR Drag",
        "Reopen_Churn_Risk": "Reopen Churn",
        "Change_Collision_Risk": "Change Collision"
    })
    hm = hm.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)

    st.markdown("### Service Stability Heatmap (Top 10 Services by Risk)")
    hm_fig = heatmap_chart(hm)
    hc1, hc2, hc3 = st.columns([1, 2, 1])
    with hc2:
        st.pyplot(hm_fig, use_container_width=False)

    st.markdown("**Top 10 Services — Risk Breakdown**")
    st.dataframe(top10, use_container_width=True)

    risk_csv = top10.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download Service Risk Table (CSV)",
        data=risk_csv,
        file_name=f"osil_service_risk_top10_{tenant_name}.csv".replace(" ", "_"),
        mime="text/csv",
    )

    st.divider()
    render_service_instability_leaders(service_risk_df)

    st.divider()

    st.markdown("### Top SIP Candidates (Next 30 Days)")
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
    st.dataframe(sip_view, use_container_width=True)

    sip_csv = sip_view.to_csv(index=False).encode("utf-8")
    st.download_button(
        label="Download SIP Candidates (CSV)",
        data=sip_csv,
        file_name=f"osil_sip_candidates_{tenant_name}.csv".replace(" ", "_"),
        mime="text/csv",
    )

    st.divider()

    st.markdown("### Executive PDF Report")

    payload = {
        "bvsi": bvsi,
        "posture": posture,
        "as_of": as_of,
        "executive_interpretation": exec_text,
        "domain_scores": domain_scores,
        "service_risk_top10": top10,
        "sip_candidates": sip_view,
        "data_readiness_score": readiness_score,
        "service_anchor_used": anchor_used,
        "detected_dataset": practice_type,
        "tenant_name": tenant_name,
    }

    colA, colB = st.columns([1, 2])

    with colA:
        if st.button("Generate / Refresh PDF", use_container_width=True):
            try:
                st.session_state.pdf_bytes = generate_pdf_bytes(payload)
                st.session_state.pdf_filename = f"OSIL_Executive_Report_{as_of}_{tenant_name}.pdf".replace(" ", "_")
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
