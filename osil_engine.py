import pandas as pd
import numpy as np

# -----------------------------
# OSIL Engine (MVP, single CSV)
# -----------------------------
# Expected minimum columns:
#   Service, Service_Tier, Opened_Date, Closed_Date, Priority, Reopened_Flag, Category
#
# Optional columns supported:
#   Resolved_Date            (preferred for MTTR if present)
#   Change_Related_Flag      (defaults to 0 if missing)
#
# Output:
#   overall scores (tier-weighted), BVSI, posture label, SIP ranking, service table

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
    # Optional columns defaults
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

    # -----------------------------
    # Rolling 12 months
    # -----------------------------
    as_of = df["Opened_Date"].max()
    start_12m = as_of - pd.Timedelta(days=365)
    df12 = df[df["Opened_Date"] >= start_12m].copy()

    # -----------------------------
    # Service-level metrics
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

    # -----------------------------
    # Domain scores (0-100; higher is better)
    # -----------------------------
    # These are "ITIL-inspired" stability dimensions, tool-agnostic.
    svc["Service_Resilience"] = (100 - (0.55 * svc["M_MTTR"] + 0.45 * svc["M_Reopen"])).clip(0, 100)
    svc["Change_Governance"]  = (100 - (0.70 * svc["M_Change"] + 0.30 * svc["M_MTTR"])).clip(0, 100)

    # Structural Risk Debt™ (MVP proxy):
    # recurring weighted incidents + reopen churn + change-induced ratio
    svc["Structural_Risk_Debt"] = (100 - (0.55 * svc["M_Recurrence"] + 0.30 * svc["M_Reopen"] + 0.15 * svc["M_Change"])).clip(0, 100)

    svc["Reliability_Momentum"] = (100 - (0.45 * svc["M_MTTR"] + 0.35 * svc["M_Recurrence"] + 0.20 * svc["M_Reopen"])).clip(0, 100)

    # -----------------------------
    # Tier-weighted aggregation (50/30/20)
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

    # -----------------------------
    # BVSI posture thresholds (explainable)
    # -----------------------------
    if bvsi < 40:
        posture = "Reactive Instability"
    elif bvsi < 60:
        posture = "Fragile Stability"
    elif bvsi < 80:
        posture = "Controlled but Exposed"
    else:
        posture = "High Confidence Operations"

    # -----------------------------
    # SIP ranking (top candidates)
    # -----------------------------
    # Prioritize Tier 1 exposure and the combined risk drivers.
    svc["SIP_Priority_Score"] = svc["Tier_Emphasis"] * (
        0.40 * svc["M_Recurrence"] +
        0.25 * svc["M_Change"] +
        0.20 * svc["M_Reopen"] +
        0.15 * svc["M_MTTR"]
    )

    sip = svc.sort_values("SIP_Priority_Score", ascending=False).head(8)[
        ["Service", "Service_Tier", "SIP_Priority_Score"]
    ].copy()

    # Suggested theme: most frequent incident category by service
    theme = df12.groupby("Service")["Category"].agg(lambda x: x.value_counts().index[0]).reset_index()
    sip = sip.merge(theme, on="Service", how="left").rename(columns={"Category": "Suggested_Theme"})

    return {
        "as_of": str(as_of.date()),
        "posture": posture,
        "overall": overall,
        "service_table": svc,
        "sip_table": sip,
        "notes": {
            "mttr_units": "MTTR is computed in hours.",
            "mttr_logic": mttr_note,
            "change_note": "Change_Related_Flag is used when present; otherwise defaulted to 0 for MVP."
        }
    }
