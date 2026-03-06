import io
import os
from datetime import date

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

from data_classifier import (
    detect_practice_type,
    normalize_service_anchor,
    calculate_data_readiness,
)

# Optional PDF generator
try:
    from report_generator import build_osil_pdf_report  # returns bytes or BytesIO (or a file path)
    REPORT_GEN_AVAILABLE = True
except Exception:
    REPORT_GEN_AVAILABLE = False


# =========================
# App Config
# =========================
st.set_page_config(page_title="OSIL™ by Xentrixus", page_icon="📊", layout="wide")
st.title("OSIL™ by Xentrixus")
st.caption("Operational Stability Intelligence (tool-agnostic) — Demo + Upload | Practice detection | Service Anchor | Data Readiness | Executive narrative")


# =========================
# Helpers
# =========================
def _safe_parse_datetime(series: pd.Series) -> pd.Series:
    return pd.to_datetime(series, errors="coerce")

def ensure_minimum_columns(df: pd.DataFrame) -> pd.DataFrame:
    defaults = {
        "Service_Tier": "Tier 3",
        "Priority": "P3",
        "Reopened_Flag": 0,
        "Change_Related_Flag": 0,
        "Category": "General",
    }
    for col, default_val in defaults.items():
        if col not in df.columns:
            df[col] = default_val

    # normalize tier/flags
    df["Service_Tier"] = df["Service_Tier"].astype(str).str.strip().replace({"": "Tier 3", "nan": "Tier 3"})
    for flag_col in ["Reopened_Flag", "Change_Related_Flag"]:
        df[flag_col] = pd.to_numeric(df[flag_col], errors="coerce").fillna(0).astype(int)

    return df

def infer_open_close_columns(df: pd.DataFrame) -> tuple[str | None, str | None]:
    open_candidates = ["Opened_Date", "Opened_At", "opened_at", "open_date", "created_at", "Created", "Created_At"]
    close_candidates = ["Closed_Date", "Closed_At", "resolved_at", "Resolved_At", "close_date", "closed_at", "Resolved", "Closed"]

    open_col = next((c for c in open_candidates if c in df.columns), None)
    close_col = next((c for c in close_candidates if c in df.columns), None)
    return open_col, close_col

def compute_mttr_hours(df: pd.DataFrame, open_col: str | None, close_col: str | None) -> pd.Series:
    if not open_col or not close_col:
        return pd.Series([np.nan] * len(df), index=df.index)
    opened = _safe_parse_datetime(df[open_col])
    closed = _safe_parse_datetime(df[close_col])
    return (closed - opened).dt.total_seconds() / 3600.0

def tier_weight(tier_val: str) -> float:
    t = str(tier_val).lower().strip()
    if "tier 0" in t or t == "0":
        return 3.5
    if "tier 1" in t or t == "1":
        return 3.0
    if "tier 2" in t or t == "2":
        return 2.0
    return 1.0

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

    # invert (lower risk => higher score)
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
    # executive-balanced, firm, "Your organization..."
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

def radar_chart(domain_scores: dict):
    labels = list(domain_scores.keys())
    values = list(domain_scores.values())

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]

    # Smaller + sharper radar (prevents "huge" look)
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
    # Ensure numeric + no NaNs to avoid int conversion crashes
    safe = heatmap_df.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)
    # We'll annotate with "—" when missing
    values = safe.to_numpy()

    fig = plt.figure(figsize=(7.6, 4.6), dpi=200)
    ax = plt.gca()
    im = ax.imshow(np.nan_to_num(values, nan=0.0), aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(range(len(safe.columns)))
    ax.set_xticklabels(list(safe.columns), fontsize=8)
    ax.set_yticks(range(len(safe.index)))
    ax.set_yticklabels(list(safe.index), fontsize=8)
    ax.set_title("Service × Stability Risk (0–100)", fontsize=11)

    # annotate safely
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

def build_service_risk_df(roll: pd.DataFrame) -> pd.DataFrame:
    # risk components (0-100)
    rec = normalize_0_100(roll["recurrence"])
    mttr = normalize_0_100(roll["mttr_hours"])
    reopen = normalize_0_100(roll["reopen_rate"])
    chg = normalize_0_100(roll["change_collision_rate"])

    df = pd.DataFrame({
        "Service": roll["Service_Anchor"].astype(str),
        "Service_Tier": roll["tier"].astype(str),
        "Recurrence_Risk": (rec * 1.0).round(1),
        "MTTR_Drag_Risk": (mttr * 1.0).round(1),
        "Reopen_Churn_Risk": (reopen * 1.0).round(1),
        "Change_Collision_Risk": (chg * 1.0).round(1),
    })

    df["Total_Service_Risk"] = (
        0.35 * df["Recurrence_Risk"] +
        0.25 * df["MTTR_Drag_Risk"] +
        0.25 * df["Reopen_Churn_Risk"] +
        0.15 * df["Change_Collision_Risk"]
    ).round(1)

    return df.sort_values("Total_Service_Risk", ascending=False).reset_index(drop=True)


# =========================
# Run Options
# =========================
st.subheader("Run Options")
c1, c2 = st.columns([1, 2])

with c1:
    run_demo = st.button("Run with Demo Data", use_container_width=True)

with c2:
    uploaded_file = st.file_uploader("Upload operational CSV export (any practice)", type=["csv"])

df = None
source_label = None

if run_demo:
    demo_path = os.path.join("data", "demo_incidents.csv")
    if not os.path.exists(demo_path):
        demo_path = "demo_incidents.csv"
    try:
        df = pd.read_csv(demo_path)
        source_label = f"Demo ({os.path.basename(demo_path)})"
    except Exception as e:
        st.error(f"Demo load failed: {e}")
        st.stop()

if df is None and uploaded_file is not None:
    try:
        df = pd.read_csv(uploaded_file)
        source_label = f"Upload ({uploaded_file.name})"
    except Exception as e:
        st.error(f"Upload load failed: {e}")
        st.stop()

if df is None:
    st.info("Upload a CSV or click **Run with Demo Data**.")
    st.stop()

st.success(f"Loaded: {source_label}")

# =========================
# Diagnostics (always run)
# =========================
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

# =========================
# OSIL Analysis (MVP)
# =========================
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

roll["tier_weight"] = roll["tier"].apply(tier_weight)

domain_scores = build_domain_scores(roll)
bvsi = compute_bvsi(domain_scores)
posture = operating_posture(bvsi)
gap = dominant_gap(domain_scores)
as_of = date.today().isoformat()
exec_text = executive_interpretation(bvsi, posture, gap)

st.divider()

m1, m2, m3 = st.columns([1, 2, 1])
m1.metric("BVSI™", f"{bvsi:.1f}")
m2.metric("Operating Posture", posture)
m3.metric("As-of Date", as_of)

st.markdown("### Executive Interpretation")
st.markdown(exec_text)

st.divider()

# =========================
# Visuals (centered)
# =========================
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

# =========================
# Heatmap + Service Risks
# =========================
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

# Force numeric & fill NaNs to avoid the crash
hm = hm.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan)

st.markdown("### Service Stability Heatmap (Top 10 Services by Risk)")
hm_fig = heatmap_chart(hm)
hc1, hc2, hc3 = st.columns([1, 2, 1])
with hc2:
    st.pyplot(hm_fig, use_container_width=False)

st.markdown("**Top 10 Services — Risk Breakdown**")
st.dataframe(top10, use_container_width=True)

st.divider()

# =========================
# Instability Leaders (Top 5)
# =========================
render_service_instability_leaders(service_risk_df)

st.divider()

# =========================
# SIP Candidates (Top 5)
# =========================
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

st.divider()

# =========================
# PDF Report
# =========================
st.markdown("### Executive PDF Report")

if not REPORT_GEN_AVAILABLE:
    st.info("PDF generator not available (report_generator.py import failed). Dashboard is still fully functional.")
else:
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
    }

    if st.button("Generate / Refresh PDF", use_container_width=True):
        try:
            pdf_obj = build_osil_pdf_report(payload)

            if isinstance(pdf_obj, (bytes, bytearray)):
                pdf_bytes = bytes(pdf_obj)
            elif isinstance(pdf_obj, io.BytesIO):
                pdf_bytes = pdf_obj.getvalue()
            elif isinstance(pdf_obj, str) and os.path.exists(pdf_obj):
                with open(pdf_obj, "rb") as f:
                    pdf_bytes = f.read()
            else:
                raise TypeError(f"Unsupported PDF return type: {type(pdf_obj)}")

            st.download_button(
                label="⬇️ Download OSIL™ Executive Report (PDF)",
                data=pdf_bytes,
                file_name=f"OSIL_Executive_Report_{as_of}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
            st.success("PDF generated.")
        except Exception as e:
            st.error(f"Report generation failed: {e}")
