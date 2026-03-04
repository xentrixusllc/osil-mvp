import pandas as pd
import numpy as np

# ---------- OSIL Engine (MVP) ----------
# Expects a single CSV with these columns (minimum):
# Service, Service_Tier, Opened_Date, Closed_Date, Priority, Reopened_Flag, Change_Related_Flag, Category
#
# MTTR is computed in hours from Opened_Date and Closed_Date.

PRIORITY_WEIGHTS = {"P1": 5, "P2": 3, "P3": 1, "P4": 0.5, "P5": 0.2}
TIER_EMPHASIS = {"Tier 1": 0.50, "Tier 2": 0.30, "Tier 3": 0.20}

def _safe_dt(s: pd.Series) -> pd.Series:
    return pd.to_datetime(s, errors="coerce", utc=True).dt.tz_convert(None)

def _tier_weight(tier: str) -> float:
    return TIER_EMPHASIS.get(str(tier).strip(), 0.20)

def run_osil(df: pd.DataFrame) -> dict:
    df = df.copy()

    # Validate minimum columns
    required = ["Service","Service_Tier","Opened_Date","Closed_Date","Priority",
            "Reopened_Flag","Category"]
# Optional column: Change_Related_Flag (default to 0 if missing)
if "Change_Related_Flag" not in df.columns:
    df["Change_Related_Flag"] = 0
    
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"Missing required columns: {missing}")

    # Parse dates
    df["Opened_Date"] = _safe_dt(df["Opened_Date"])
    df["Closed_Date"] = _safe_dt(df["Closed_Date"])
    df = df.dropna(subset=["Opened_Date","Closed_Date"])

    # Compute MTTR hours
    df["MTTR_Hours"] = (df["Closed_Date"] - df["Opened_Date"]).dt.total_seconds() / 3600.0
    df["MTTR_Hours"] = df["MTTR_Hours"].clip(lower=0.0)

    # Priority weights
    df["Priority_Weight"] = df["Priority"].astype(str).str.strip().map(PRIORITY_WEIGHTS).fillna(1.0)

    # Reopen + change flags
    df["Reopened_Flag"] = pd.to_numeric(df["Reopened_Flag"], errors="coerce").fillna(0).astype(int)
    df["Change_Related_Flag"] = pd.to_numeric(df["Change_Related_Flag"], errors="coerce").fillna(0).astype(int)

    # Rolling 12 months
    as_of = df["Opened_Date"].max()
    start_12m = as_of - pd.Timedelta(days=365)
    df12 = df[df["Opened_Date"] >= start_12m].copy()

    # Service metrics
    svc = df12.groupby(["Service","Service_Tier"]).agg(
        Incidents=("Service","count"),
        Avg_MTTR=("MTTR_Hours","mean"),
        Reopen_Rate=("Reopened_Flag","mean"),
        Change_Induced_Ratio=("Change_Related_Flag","mean"),
        PriorityWeighted=("Priority_Weight","sum"),
    ).reset_index()

    # Normalization helper (0-100 where higher metric value = more risk)
    def norm(series):
        mn, mx = series.min(), series.max()
        if mx == mn:
            return pd.Series(np.zeros(len(series)))
        return (series - mn) / (mx - mn) * 100

    svc["M_Change"] = norm(svc["Change_Induced_Ratio"])
    svc["M_Recurrence"] = norm(svc["PriorityWeighted"])
    svc["M_Reopen"] = norm(svc["Reopen_Rate"])
    svc["M_MTTR"] = norm(svc["Avg_MTTR"])

    # Domain scores (0-100; higher is better)
    svc["Service_Resilience"] = (100 - (0.55*svc["M_MTTR"] + 0.45*svc["M_Reopen"])).clip(0,100)
    svc["Change_Governance"]  = (100 - (0.70*svc["M_Change"] + 0.30*svc["M_MTTR"])).clip(0,100)
    svc["Structural_Risk_Debt"]= (100 - (0.55*svc["M_Recurrence"] + 0.30*svc["M_Reopen"] + 0.15*svc["M_Change"])).clip(0,100)
    svc["Reliability_Momentum"]= (100 - (0.45*svc["M_MTTR"] + 0.35*svc["M_Recurrence"] + 0.20*svc["M_Reopen"])).clip(0,100)

    # Tier emphasis aggregation (50/30/20)
    def tier_weighted(col):
        return (
            (svc[col] * svc["Service_Tier"].map(_tier_weight)).sum()
            / (svc["Service_Tier"].map(_tier_weight).sum())
        )

    overall = {
        "Overall Service Resilience": float(tier_weighted("Service_Resilience")),
        "Overall Change Governance": float(tier_weighted("Change_Governance")),
        "Overall Structural Risk Debt": float(tier_weighted("Structural_Risk_Debt")),
        "Overall Reliability Momentum": float(tier_weighted("Reliability_Momentum")),
    }
    bvsi = float(np.mean(list(overall.values())))
    overall["BVSI"] = bvsi

    # Classification thresholds
    if bvsi < 40:
        posture = "Reactive Instability"
    elif bvsi < 60:
        posture = "Fragile Stability"
    elif bvsi < 80:
        posture = "Controlled but Exposed"
    else:
        posture = "High Confidence Operations"

    # SIP ranking (simple: tier emphasis * (recurrence + change + reopen))
    svc["Tier_Emphasis"] = svc["Service_Tier"].map(_tier_weight)
    svc["SIP_Priority_Score"] = svc["Tier_Emphasis"] * (0.40*svc["M_Recurrence"] + 0.25*svc["M_Change"] + 0.20*svc["M_Reopen"] + 0.15*svc["M_MTTR"])
    sip = svc.sort_values("SIP_Priority_Score", ascending=False).head(5)[
        ["Service","Service_Tier","SIP_Priority_Score"]
    ]

    # Suggested theme (most common category per service)
    theme = df12.groupby("Service")["Category"].agg(lambda x: x.value_counts().index[0]).reset_index()
    sip = sip.merge(theme, on="Service", how="left").rename(columns={"Category":"Suggested_Theme"})

    return {
        "as_of": str(as_of.date()),
        "posture": posture,
        "overall": overall,
        "service_table": svc,
        "sip_table": sip,
    }
