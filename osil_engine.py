import pandas as pd
import numpy as np

# -----------------------------
# OSIL Engine (v0.6 - SIP by Service + Theme)
# -----------------------------
# Minimum required columns:
#   Service, Service_Tier, Opened_Date, Closed_Date, Priority, Reopened_Flag, Category
#
# Optional columns:
#   Resolved_Date            (preferred for MTTR if present)
#   Change_Related_Flag      (defaults to 0 if missing)
#
# Output:
#   - overall (tier-weighted domain scores + BVSI)
#   - posture
#   - service_table (service-level domain measures)
#   - sip_table (Service + Theme SIP candidates + why/priority label)

PRIORITY_WEIGHTS = {"P1": 5.0, "P2": 3.0, "P3": 1.0, "P4": 0.5, "P5": 0.2}
TIER_EMPHASIS = {"Tier 1": 0.50, "Tier 2": 0.30, "Tier 3": 0.20}

def _safe_dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce")

def _tier_weight(tier: str) -> float:
    return TIER_EMPHASIS.get(str(tier).strip(), 0.20)

def _norm(series: pd.Series) -> pd.Series:
    """Normalize to 0..100 where higher value = higher risk."""
    mn, mx = series.min(), series.max()
    if pd.isna(mn) or pd.isna(mx) or mx == mn:
        return pd.Series(np.zeros(len(series)), index=series.index)
    return (series - mn) / (mx - mn) * 100

def _priority_label(score_0_100: float) -> str:
    # Simple executive labeling
    if score_0_100 >= 70:
        return "Immediate SIP"
    if score_0_100 >= 40:
        return "Next SIP"
    if score_0_100 >= 15:
        return "Monitor"
    return "Low"

def _why_flagged(row) -> str:
    drivers = {
        "Recurrence/Volume": row.get("M_Recurrence", 0.0),
        "Change-Induced": row.get("M_Change", 0.0),
        "Reopen Churn": row.get("M_Reopen", 0.0),
        "MTTR Drag": row.get("M_MTTR", 0.0),
    }
    top = sorted(drivers.items(), key=lambda x: x[1], reverse=True)[:2]
    return " + ".join([t[0] for t in top])

def run_osil(df: pd.DataFrame) -> dict:
    df = df.copy()

    # -----------------------------
    # Validate minimum columns
    # -----------------------------
    required = ["Service", "Service_Tier", "Opened_Date", "Closed_Date", "Priority", "Reopened_Flag", "Category"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # -----------------------------
    # Optional defaults
    # -----------------------------
    if "Change_Related_Flag" not in df.columns:
        df["Change_Related_Flag"] = 0  # MVP default

    # -----------------------------
    # Parse dates
    # -----------------------------
    df["Opened_Date"] = _safe_dt(df["Opened_Date"])
    df["Closed_Date"] = _safe_dt(df["Closed_Date"])

    mttr_note = "MTTR uses Opened_Date → Closed_Date (Resolved_Date not provided)."
    if "Resolved_Date" in df.columns:
        df["Resolved_Date"] = _safe_dt(df["Resolved_Date"])
        df["Effective_End_Date"] = df["Resolved_Date"].where(df["Resolved_Date"].notna(), df["Closed_Date"])
        mttr_note = "MTTR uses Opened_Date → Resolved_Date when available; otherwise falls back to Closed_Date."
    else:
        df["Effective_End_Date"] = df["Closed_Date"]

    df = df.dropna(subset=["Opened_Date", "Effective_End_Date"])

    # -----------------------------
    # Compute MTTR in hours
    # -----------------------------
    df["MTTR_Hours"] = (df["Effective_End_Date"] - df["Opened_Date"]).dt.total_seconds() / 3600.0
    df["MTTR_Hours"] = df["MTTR_Hours"].clip(lower=0.0)

    # -----------------------------
    # Normalize flags and weights
    # -----------------------------
    df["Priority_Weight"] = df["Priority"].astype(str).str.strip().map(PRIORITY_WEIGHTS).fillna(1.0)
    df["Reopened_Flag"] = pd.to_numeric(df["Reopened_Flag"], errors="coerce").fillna(0).astype(int)
    df["Change_Related_Flag"] = pd.to_numeric(df["Change_Related_Flag"], errors="coerce").fillna(0).astype(int)
    df["Category"] = df["Category"].astype(str).fillna("Unclassified")

    # -----------------------------
    # Rolling 12 months
    # -----------------------------
    as_of = df["Opened_Date"].max()
    start_12m = as_of - pd.Timedelta(days=365)
    df12 = df[df["Opened_Date"] >= start_12m].copy()

    # -----------------------------
    # Service-level table (stable foundation)
    # -----------------------------
    svc = df12.groupby(["Service", "Service_Tier"]).agg(
        Incidents=("Service", "count"),
        Avg_MTTR=("MTTR_Hours", "mean"),
        Reopen_Rate=("Reopened_Flag", "mean"),
        Change_Induced_Ratio=("Change_Related_Flag", "mean"),
        PriorityWeighted=("Priority_Weight", "sum"),
    ).reset_index()

    # Risk-normalized measures (higher = worse)
    svc["M_Change"] = _norm(svc["Change_Induced_Ratio"])
    svc["M_Recurrence"] = _norm(svc["PriorityWeighted"])
    svc["M_Reopen"] = _norm(svc["Reopen_Rate"])
    svc["M_MTTR"] = _norm(svc["Avg_MTTR"])

    # Domain scores (0-100; higher is better)
    svc["Service_Resilience"] = (100 - (0.55 * svc["M_MTTR"] + 0.45 * svc["M_Reopen"])).clip(0, 100)
    svc["Change_Governance"]  = (100 - (0.70 * svc["M_Change"] + 0.30 * svc["M_MTTR"])).clip(0, 100)
    svc["Structural_Risk_Debt"] = (100 - (0.55 * svc["M_Recurrence"] + 0.30 * svc["M_Reopen"] + 0.15 * svc["M_Change"])).clip(0, 100)
    svc["Reliability_Momentum"] = (100 - (0.45 * svc["M_MTTR"] + 0.35 * svc["M_Recurrence"] + 0.20 * svc["M_Reopen"])).clip(0, 100)

    # -----------------------------
    # Tier-weighted overall
    # -----------------------------
    svc["Tier_Emphasis"] = svc["Service_Tier"].map(_tier_weight)

    def tier_weighted(col: str) -> float:
        denom = float(svc["Tier_Emphasis"].sum())
        if denom == 0:
            return 0.0
        return float((svc[col] * svc["Tier_Emphasis"]).sum() / denom)

    overall = {
        "Overall Service Resilience": tier_weighted("Service_Resilience"),
        "Overall Change Governance": tier_weighted("Change_Governance"),
        "Overall Structural Risk Debt": tier_weighted("Structural_Risk_Debt"),
        "Overall Reliability Momentum": tier_weighted("Reliability_Momentum"),
    }

    bvsi = float(np.mean(list(overall.values())))
    overall["BVSI"] = bvsi

    # Posture thresholds
    if bvsi < 40:
        posture = "Reactive Instability"
    elif bvsi < 60:
        posture = "Fragile Stability"
    elif bvsi < 80:
        posture = "Controlled but Exposed"
    else:
        posture = "High Confidence Operations"

    # -----------------------------
    # SIP by Service + Theme (Category)
    # -----------------------------
    sipg = df12.groupby(["Service", "Service_Tier", "Category"]).agg(
        Incidents=("Service", "count"),
        Avg_MTTR=("MTTR_Hours", "mean"),
        Reopen_Rate=("Reopened_Flag", "mean"),
        Change_Induced_Ratio=("Change_Related_Flag", "mean"),
        PriorityWeighted=("Priority_Weight", "sum"),
    ).reset_index()

    # Normalize risks across Service+Theme rows
    sipg["M_Change"] = _norm(sipg["Change_Induced_Ratio"])
    sipg["M_Recurrence"] = _norm(sipg["PriorityWeighted"])
    sipg["M_Reopen"] = _norm(sipg["Reopen_Rate"])
    sipg["M_MTTR"] = _norm(sipg["Avg_MTTR"])
    sipg["Tier_Emphasis"] = sipg["Service_Tier"].map(_tier_weight)

    # SIP Priority Score (higher = more urgent)
    sipg["SIP_Priority_Score"] = sipg["Tier_Emphasis"] * (
        0.40 * sipg["M_Recurrence"] +
        0.25 * sipg["M_Change"] +
        0.20 * sipg["M_Reopen"] +
        0.15 * sipg["M_MTTR"]
    )

    # Convert score to 0..100 for executive labeling
    sipg["SIP_Score_0_100"] = _norm(sipg["SIP_Priority_Score"])

    sipg["Priority_Label"] = sipg["SIP_Score_0_100"].apply(_priority_label)
    sipg["Why_Flagged"] = sipg.apply(_why_flagged, axis=1)

    sip_table = sipg.sort_values("SIP_Priority_Score", ascending=False).head(12)[
        ["Service", "Service_Tier", "Category", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"]
    ].rename(columns={"Category": "Suggested_Theme"}).copy()

    return {
        "as_of": str(as_of.date()),
        "posture": posture,
        "overall": overall,
        "service_table": svc,
        "sip_table": sip_table,
        "notes": {
            "mttr_units": "MTTR is computed in hours.",
            "mttr_logic": mttr_note,
            "change_note": "Change_Related_Flag is used when present; otherwise defaulted to 0 for MVP.",
            "sip_note": "SIP candidates are prioritized by Service + Theme (Category) with Tier weighting."
        }
    }
