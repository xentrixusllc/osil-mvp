import io
import os
from datetime import datetime, date

import numpy as np
import pandas as pd
import streamlit as st
import matplotlib.pyplot as plt

# --- Your helper module (you created this) ---
from data_classifier import (
    detect_practice_type,
    normalize_service_anchor,
    calculate_data_readiness,
)

# --- Optional: your existing PDF generator (if present) ---
try:
    from report_generator import build_osil_pdf_report  # should return bytes OR BytesIO
    REPORT_GEN_AVAILABLE = True
except Exception:
    REPORT_GEN_AVAILABLE = False


# =========================
# App Config
# =========================
st.set_page_config(
    page_title="OSIL by Xentrixus",
    page_icon="📈",
    layout="wide"
)

st.title("OSIL by Xentrixus — Stability Intelligence MVP")
st.caption("Upload operational exports (any practice) → OSIL detects type, normalizes Service Anchor, scores readiness, and produces executive stability intelligence.")


# =========================
# Utility Functions
# =========================

def _safe_parse_datetime(series: pd.Series) -> pd.Series:
    """Parse date/time safely across common formats."""
    return pd.to_datetime(series, errors="coerce", utc=False)

def ensure_minimum_columns(df: pd.DataFrame) -> pd.DataFrame:
    """
    Ensure OSIL minimum fields exist, even if the dataset is sparse.
    This prevents repeated 'missing required columns' errors.
    """
    # Canonical columns used by OSIL logic
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

    # Normalize tier if present but messy
    if "Service_Tier" in df.columns:
        df["Service_Tier"] = df["Service_Tier"].astype(str).str.strip().replace({"": "Tier 3", "nan": "Tier 3"})

    # Normalize flags
    for flag_col in ["Reopened_Flag", "Change_Related_Flag"]:
        if flag_col in df.columns:
            df[flag_col] = pd.to_numeric(df[flag_col], errors="coerce").fillna(0).astype(int)

    return df

def infer_open_close_columns(df: pd.DataFrame) -> tuple[str, str]:
    """
    Find best matching open/close columns from common exports.
    """
    open_candidates = ["Opened_Date", "Opened_At", "opened_at", "open_date", "created_at", "Created", "Created_At"]
    close_candidates = ["Closed_Date", "Closed_At", "resolved_at", "Resolved_At", "close_date", "closed_at", "Resolved", "Closed"]

    open_col = next((c for c in open_candidates if c in df.columns), None)
    close_col = next((c for c in close_candidates if c in df.columns), None)

    return open_col, close_col

def tier_weight(tier_val: str) -> float:
    """
    Convert tier label to weight. Tier 1 most critical.
    """
    t = str(tier_val).lower().strip()
    if "tier 1" in t or t == "1":
        return 3.0
    if "tier 2" in t or t == "2":
        return 2.0
    if "tier 0" in t or t == "0":
        return 3.5
    return 1.0  # Tier 3+ default

def normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(0)
    if s.max() == s.min():
        return pd.Series([0] * len(s), index=s.index)
    return ((s - s.min()) / (s.max() - s.min()) * 100).round(0)

def compute_mttr_hours(df: pd.DataFrame, open_col: str, close_col: str) -> pd.Series:
    if open_col is None or close_col is None:
        return pd.Series([np.nan] * len(df), index=df.index)

    opened = _safe_parse_datetime(df[open_col])
    closed = _safe_parse_datetime(df[close_col])
    delta = (closed - opened).dt.total_seconds() / 3600.0
    return delta

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

def executive_interpretation(bvsi: float, posture: str, dominant_gap: str) -> str:
    # Keep it executive-balanced, firm, and “Your organization…”
    if bvsi >= 85:
        return (
            f"Your organization demonstrates a **{posture}** operating posture (BVSI™ {bvsi:.1f}). "
            f"Stability signals are generally balanced, and targeted prevention work can further strengthen "
            f"executive confidence and preserve customer trust. Primary focus: **{dominant_gap}**."
        )
    if bvsi >= 70:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Governance and control exist, but improvement is uneven across stability domains. "
            f"Prioritize SIPs where Tier-1 exposure and recurrence patterns concentrate. Primary focus: **{dominant_gap}**."
        )
    if bvsi >= 55:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Operational control exists, but recurring instability patterns still create exposure—especially across "
            f"higher-impact services. Focused SIP execution over the next 30–60 days will increase executive confidence "
            f"and protect customer trust. Primary focus: **{dominant_gap}**."
        )
    if bvsi >= 40:
        return (
            f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
            f"Instability is material and likely visible to business stakeholders. "
            f"Immediate stabilization actions and tighter change controls are recommended. Primary focus: **{dominant_gap}**."
        )
    return (
        f"Your organization is operating in a **{posture}** posture (BVSI™ {bvsi:.1f}). "
        f"Operational risk is elevated and may threaten reliability commitments. "
        f"Stabilization should be treated as an executive priority. Primary focus: **{dominant_gap}**."
    )

def build_domain_scores(service_rollup: pd.DataFrame) -> dict:
    """
    Create 4 stability domain scores (0–100) from rollups.
    Heuristic by design for MVP (tool-agnostic).
    """
    # Inputs are expected columns: recurrence, reopen_rate, mttr_hours, change_collision_rate, tier_weight
    # Higher is worse for recurrence/mttr/reopen/change_collision, so we invert where appropriate.
    rec = normalize_0_100(service_rollup["recurrence"])
    mttr = normalize_0_100(service_rollup["mttr_hours"].fillna(service_rollup["mttr_hours"].median() if len(service_rollup) else 0))
    reopen = normalize_0_100(service_rollup["reopen_rate"])
    chg = normalize_0_100(service_rollup["change_collision_rate"])

    # Domain scoring (invert: lower risk -> higher score)
    service_resilience = float((100 - (0.55 * rec + 0.45 * mttr)).clip(0, 100).mean()) if len(service_rollup) else 0.0
    change_governance = float((100 - chg).clip(0, 100).mean()) if len(service_rollup) else 0.0
    structural_risk_debt = float((100 - (0.65 * rec + 0.35 * reopen)).clip(0, 100).mean()) if len(service_rollup) else 0.0
    reliability_momentum = float((100 - (0.50 * rec + 0.25 * reopen + 0.25 * mttr)).clip(0, 100).mean()) if len(service_rollup) else 0.0

    return {
        "Service Resilience": round(service_resilience, 1),
        "Change Governance": round(change_governance, 1),
        "Structural Risk Debt™": round(structural_risk_debt, 1),
        "Reliability Momentum": round(reliability_momentum, 1),
    }

def compute_bvsi(domain_scores: dict) -> float:
    # Balanced weights for MVP. You can tune later.
    w = {
        "Service Resilience": 0.30,
        "Change Governance": 0.25,
        "Structural Risk Debt™": 0.25,
        "Reliability Momentum": 0.20,
    }
    bvsi = 0.0
    for k, weight in w.items():
        bvsi += float(domain_scores.get(k, 0)) * weight
    return round(bvsi, 1)

def dominant_gap(domain_scores: dict) -> str:
    # Lowest score is the biggest gap
    if not domain_scores:
        return "Stability signal coverage"
    return sorted(domain_scores.items(), key=lambda x: x[1])[0][0]

def radar_chart(domain_scores: dict):
    labels = list(domain_scores.keys())
    values = list(domain_scores.values())

    if len(labels) == 0:
        return None

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + [values[0]]
    angles_loop = angles + [angles[0]]

    fig = plt.figure(figsize=(6, 6), dpi=180)
    ax = plt.subplot(111, polar=True)
    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.10)
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)
    ax.set_yticklabels([])
    ax.set_ylim(0, 100)
    ax.set_title("Operational Stability Profile (Radar)", pad=20)
    return fig

def heatmap_chart(heatmap_df: pd.DataFrame):
    fig = plt.figure(figsize=(10, 6), dpi=160)
    ax = plt.gca()
    im = ax.imshow(heatmap_df.values, aspect="auto")
    ax.set_xticks(range(len(heatmap_df.columns)))
    ax.set_xticklabels(list(heatmap_df.columns), rotation=0)
    ax.set_yticks(range(len(heatmap_df.index)))
    ax.set_yticklabels(list(heatmap_df.index))
    ax.set_title("Service × Stability Risk (0–100)")
    plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04, label="Risk Score (0–100)")

    # annotate
    for i in range(heatmap_df.shape[0]):
        for j in range(heatmap_df.shape[1]):
            ax.text(j, i, f"{int(heatmap_df.iat[i, j])}", ha="center", va="center", fontsize=9)

    return fig


# =========================
# Data Loading (Demo / Upload)
# =========================
st.subheader("Run Options")

cA, cB = st.columns([1, 2])

with cA:
    run_demo = st.button("Run with Demo Data")

with cB:
    uploaded_file = st.file_uploader("Upload your operational CSV export (any practice)", type=["csv"])

df = None
source_label = None

# Demo load
if run_demo:
    # Default demo file path
    demo_path = os.path.join("data", "demo_incidents.csv")
    if not os.path.exists(demo_path):
        # fallback: older location
        demo_path = "demo_incidents.csv"

    try:
        df = pd.read_csv(demo_path)
        source_label = f"Demo dataset ({os.path.basename(demo_path)})"
    except Exception as e:
        st.error(f"Demo load failed: {e}")

# Upload load
if (df is None) and (uploaded_file is not None):
    try:
        df = pd.read_csv(uploaded_file)
        source_label = f"Uploaded dataset ({uploaded_file.name})"
    except Exception as e:
        st.error(f"Upload load failed: {e}")

if df is None:
    st.info("Upload a CSV or click **Run with Demo Data**.")
    st.stop()

st.success(f"Loaded: {source_label}")


# =========================
# Practice Detection + Anchor + Readiness (ALWAYS runs for demo and upload)
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
# OSIL Analysis (MVP) — works even with imperfect exports
# =========================
open_col, close_col = infer_open_close_columns(df)

# Compute MTTR (hours) where possible
df["__mttr_hours"] = compute_mttr_hours(df, open_col, close_col)

# Build service rollup on Service_Anchor
roll = df.groupby("Service_Anchor", dropna=False).agg(
    recurrence=("Service_Anchor", "count"),
    reopen_rate=("Reopened_Flag", "mean"),
    change_collision_rate=("Change_Related_Flag", "mean"),
    mttr_hours=("__mttr_hours", "mean"),
    tier=("Service_Tier", lambda x: x.value_counts().index[0] if len(x) else "Tier 3"),
    category=("Category", lambda x: x.value_counts().index[0] if len(x) else "General"),
).reset_index()

roll["tier_weight"] = roll["tier"].apply(tier_weight)

# Domain scores + BVSI
domain_scores = build_domain_scores(roll)
bvsi = compute_bvsi(domain_scores)
posture = operating_posture(bvsi)
gap = dominant_gap(domain_scores)
as_of = date.today().isoformat()

# Executive interpretation
exec_text = executive_interpretation(bvsi, posture, gap)

st.divider()

# Header Metrics
m1, m2, m3 = st.columns([1, 2, 1])
m1.metric("BVSI™", f"{bvsi:.1f}")
m2.metric("Operating Posture", posture)
m3.metric("As-of Date", as_of)

st.markdown("### Executive Interpretation")
st.markdown(exec_text)

st.divider()

# =========================
# Visuals
# =========================
st.markdown("### Operational Stability Profile (Radar)")
rad = radar_chart(domain_scores)
if rad is not None:
    st.pyplot(rad, use_container_width=False)

st.markdown("### Stability Domain Scores (0–100)")
st.dataframe(
    pd.DataFrame({"Domain": list(domain_scores.keys()), "Score": list(domain_scores.values())}),
    use_container_width=True
)

# Heatmap (Top 10 services by risk)
st.markdown("### Service Stability Heatmap (Top 10 Services by Risk)")
heat_source = roll.copy()
heat_source["rec_norm"] = normalize_0_100(heat_source["recurrence"])
heat_source["mttr_norm"] = normalize_0_100(heat_source["mttr_hours"].fillna(heat_source["mttr_hours"].median() if len(heat_source) else 0))
heat_source["reopen_norm"] = normalize_0_100(heat_source["reopen_rate"])
heat_source["chg_norm"] = normalize_0_100(heat_source["change_collision_rate"])

heat_source["risk_total"] = (
    0.35 * heat_source["rec_norm"] +
    0.25 * heat_source["mttr_norm"] +
    0.25 * heat_source["reopen_norm"] +
    0.15 * heat_source["chg_norm"]
)

top10 = heat_source.sort_values("risk_total", ascending=False).head(10)

hm = pd.DataFrame(
    {
        "Recurrence": top10["rec_norm"].round(0),
        "MTTR Drag": top10["mttr_norm"].round(0),
        "Reopen Churn": top10["reopen_norm"].round(0),
        "Change Collision": top10["chg_norm"].round(0),
    },
    index=[f"{r.Service_Anchor} ({r.tier})" for r in top10.itertuples()]
)

hm_fig = heatmap_chart(hm)
st.pyplot(hm_fig, use_container_width=True)

st.divider()


# =========================
# SIP Candidates (Top 5 by default)
# =========================
st.markdown("### Top SIP Candidates (Next 30 Days)")

# SIP priority score heuristic
sip = roll.copy()

sip["SIP_Priority_Score"] = (
    sip["tier_weight"] *
    (
        0.45 * normalize_0_100(sip["recurrence"]) +
        0.25 * normalize_0_100(sip["mttr_hours"].fillna(sip["mttr_hours"].median() if len(sip) else 0)) +
        0.20 * normalize_0_100(sip["reopen_rate"]) +
        0.10 * normalize_0_100(sip["change_collision_rate"])
    )
).round(4)

def priority_label(score):
    if score >= 30:
        return "Next SIP"
    if score >= 12:
        return "Monitor"
    return "Backlog"

sip["Priority_Label"] = sip["SIP_Priority_Score"].apply(priority_label)
sip["Suggested_Theme"] = sip["category"].astype(str)

# Why flagged
why_parts = []
if "recurrence" in sip.columns:
    why_parts.append("Recurrence/Volume")
sip["Why_Flagged"] = (
    "Recurrence/Volume + MTTR drag + Reopen churn + Change collision"
)

sip_view = sip.sort_values("SIP_Priority_Score", ascending=False).head(5)[
    ["Service_Anchor", "tier", "Suggested_Theme", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"]
].rename(
    columns={"Service_Anchor": "Service", "tier": "Service_Tier"}
)

st.dataframe(sip_view, use_container_width=True)

st.divider()

# =========================
# Analyst Review (simple, practical)
# =========================
st.markdown("### Analyst Review (What to validate before committing)")
st.markdown(
    "- Confirm service naming consistency (Service vs Application vs CI).\n"
    "- Validate change linkage quality (Change_Related_Flag may be underreported).\n"
    "- Verify MTTR calculation (requires valid opened/closed timestamps).\n"
    "- Ensure Tier/Criticality aligns to business impact (Tier 1 should represent highest exposure).\n"
    "- Use SIP candidates as **starting points**; validate with SMEs before execution."
)

st.divider()

# =========================
# PDF Report Download
# =========================
st.markdown("### Executive PDF Report")

if not REPORT_GEN_AVAILABLE:
    st.info("PDF generator not available (report_generator.py import failed). The dashboard is still working.")
else:
    # Build a compact results dict that most report generators can use
    results_payload = {
        "bvsi": bvsi,
        "posture": posture,
        "as_of": as_of,
        "executive_interpretation": exec_text,
        "domain_scores": domain_scores,
        "sip_candidates": sip_view,
        "heatmap_df": hm,
        "data_readiness_score": readiness_score,
        "service_anchor_used": anchor_used,
        "detected_dataset": practice_type,
    }

    try:
        pdf_obj = build_osil_pdf_report(results_payload)

        if isinstance(pdf_obj, (bytes, bytearray)):
            pdf_bytes = bytes(pdf_obj)
        elif isinstance(pdf_obj, io.BytesIO):
            pdf_bytes = pdf_obj.getvalue()
        else:
            # Some generators return a file path
            if isinstance(pdf_obj, str) and os.path.exists(pdf_obj):
                with open(pdf_obj, "rb") as f:
                    pdf_bytes = f.read()
            else:
                raise TypeError(f"Unsupported PDF return type: {type(pdf_obj)}")

        filename = f"OSIL_Executive_Report_{as_of}.pdf"
        st.download_button(
            label="Download OSIL™ Executive Report (PDF)",
            data=pdf_bytes,
            file_name=filename,
            mime="application/pdf",
            use_container_width=True
        )

    except Exception as e:
        st.error(f"Report generation failed: {e}")
