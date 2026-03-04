from __future__ import annotations

from dataclasses import dataclass
from typing import Dict, Any, Tuple, List

import numpy as np
import pandas as pd


# -----------------------------
# OSIL™ Core Configuration
# -----------------------------
REQUIRED_COLUMNS = [
    "Service",
    "Service_Tier",
    "Opened_Date",
    "Closed_Date",
    "Priority",
    "Reopened_Flag",
    "Change_Related_Flag",
    "Category",
]


TIER_MULTIPLIER = {
    "Tier 1": 1.30,
    "Tier 2": 1.10,
    "Tier 3": 1.00,
    "Tier 4": 0.90,
    "Tier 0": 1.40,  # if someone has it
}


PRIORITY_WEIGHT = {
    "P1": 5,
    "P2": 4,
    "P3": 3,
    "P4": 2,
    "P5": 1,
}


@dataclass
class OSILResults:
    overall: Dict[str, float]
    posture: str
    as_of: str
    sip_table: pd.DataFrame
    analyst_review: str
    service_risk_df: pd.DataFrame


# -----------------------------
# Helpers
# -----------------------------
def _normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(0.0)
    if s.nunique() <= 1:
        return pd.Series([0.0] * len(s), index=s.index)
    mn, mx = float(s.min()), float(s.max())
    if mx - mn == 0:
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - mn) / (mx - mn) * 100.0).clip(0, 100)


def _safe_parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out["Opened_Date"] = pd.to_datetime(out["Opened_Date"], errors="coerce")
    out["Closed_Date"] = pd.to_datetime(out["Closed_Date"], errors="coerce")
    return out


def _duration_hours(opened: pd.Series, closed: pd.Series) -> pd.Series:
    delta = (closed - opened).dt.total_seconds() / 3600.0
    return pd.to_numeric(delta, errors="coerce").clip(lower=0)


def _tier_mult(series: pd.Series) -> pd.Series:
    return series.astype(str).map(TIER_MULTIPLIER).fillna(1.0)


def _priority_w(series: pd.Series) -> pd.Series:
    return series.astype(str).str.upper().map(PRIORITY_WEIGHT).fillna(2.0)


def validate_required_columns(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    missing = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    return (len(missing) == 0, missing)


def compute_service_risk(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """
    Returns Service × Stability Risk table:
      - Recurrence_Risk (0-100)
      - MTTR_Drag_Risk (0-100)
      - Reopen_Churn_Risk (0-100)
      - Change_Collision_Risk (0-100)
      - Total_Service_Risk (0-100)
    """
    if df.empty:
        return pd.DataFrame()

    d = df.copy()

    # Basic hygiene
    d["Service"] = d["Service"].astype(str).fillna("Unknown")
    d["Service_Tier"] = d["Service_Tier"].astype(str).fillna("Tier 3")
    d["Priority"] = d["Priority"].astype(str).str.upper().fillna("P3")
    d["Reopened_Flag"] = pd.to_numeric(d["Reopened_Flag"], errors="coerce").fillna(0).clip(0, 1)
    d["Change_Related_Flag"] = pd.to_numeric(d["Change_Related_Flag"], errors="coerce").fillna(0).clip(0, 1)

    # MTTR hours per record
    d["Duration_Hours"] = _duration_hours(d["Opened_Date"], d["Closed_Date"]).fillna(0)

    # weighted incident impact proxy (priority × tier multiplier)
    d["Tier_Mult"] = _tier_mult(d["Service_Tier"])
    d["Priority_Weight"] = _priority_w(d["Priority"])
    d["Impact_Weight"] = d["Tier_Mult"] * d["Priority_Weight"]

    # Aggregate by service
    g = d.groupby(["Service", "Service_Tier"], as_index=False).agg(
        Ticket_Count=("Service", "size"),
        Weighted_Count=("Impact_Weight", "sum"),
        Avg_MTTR_Hours=("Duration_Hours", "mean"),
        Reopen_Rate=("Reopened_Flag", "mean"),
        Change_Rate=("Change_Related_Flag", "mean"),
    )

    # Component risks (0–100)
    # Recurrence uses weighted count (more executive-relevant than raw volume)
    g["Recurrence_Risk"] = _normalize_0_100(g["Weighted_Count"])

    # MTTR drag normalized across services
    g["MTTR_Drag_Risk"] = _normalize_0_100(g["Avg_MTTR_Hours"])

    # Rates → 0–100 directly, but with gentle scaling (so 10% reopen is not 10 risk)
    # We use sqrt to emphasize early churn
    g["Reopen_Churn_Risk"] = (np.sqrt(g["Reopen_Rate"].clip(0, 1)) * 100.0).clip(0, 100)
    g["Change_Collision_Risk"] = (np.sqrt(g["Change_Rate"].clip(0, 1)) * 100.0).clip(0, 100)

    # Total risk (weighted)
    # You can tune these later — keep visible.
    w_rec = 0.35
    w_mttr = 0.25
    w_reopen = 0.20
    w_change = 0.20

    g["Total_Service_Risk_Raw"] = (
        g["Recurrence_Risk"] * w_rec
        + g["MTTR_Drag_Risk"] * w_mttr
        + g["Reopen_Churn_Risk"] * w_reopen
        + g["Change_Collision_Risk"] * w_change
    )

    # Apply tier multiplier to total (Tier 1 risk should surface higher)
    g["Total_Service_Risk"] = (g["Total_Service_Risk_Raw"] * _tier_mult(g["Service_Tier"])).clip(0, 100)

    # Keep top N by Total_Service_Risk
    g = g.sort_values("Total_Service_Risk", ascending=False).head(top_n).reset_index(drop=True)

    # Clean output
    out = g[
        [
            "Service",
            "Service_Tier",
            "Recurrence_Risk",
            "MTTR_Drag_Risk",
            "Reopen_Churn_Risk",
            "Change_Collision_Risk",
            "Total_Service_Risk",
            "Ticket_Count",
            "Avg_MTTR_Hours",
            "Reopen_Rate",
            "Change_Rate",
        ]
    ].copy()

    # Round for readability
    for c in ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk", "Total_Service_Risk"]:
        out[c] = out[c].round(1)

    out["Avg_MTTR_Hours"] = out["Avg_MTTR_Hours"].round(2)
    out["Reopen_Rate"] = (out["Reopen_Rate"] * 100).round(1)  # percentage
    out["Change_Rate"] = (out["Change_Rate"] * 100).round(1)  # percentage

    return out


def compute_overall_scores(df: pd.DataFrame) -> Dict[str, float]:
    """
    Produces OSIL domain scores and BVSI:
    - Overall Service Resilience
    - Overall Change Governance
    - Overall Structural Risk Debt
    - Overall Reliability Momentum
    - BVSI
    """
    if df.empty:
        return {
            "Overall Service Resilience": 0.0,
            "Overall Change Governance": 0.0,
            "Overall Structural Risk Debt": 0.0,
            "Overall Reliability Momentum": 0.0,
            "BVSI": 0.0,
        }

    # Using simple, defensible heuristics for MVP.
    # These can evolve later into ML/predictive models.

    d = df.copy()
    d["Duration_Hours"] = _duration_hours(d["Opened_Date"], d["Closed_Date"]).fillna(0)
    d["Reopened_Flag"] = pd.to_numeric(d["Reopened_Flag"], errors="coerce").fillna(0).clip(0, 1)
    d["Change_Related_Flag"] = pd.to_numeric(d["Change_Related_Flag"], errors="coerce").fillna(0).clip(0, 1)
    d["Tier_Mult"] = _tier_mult(d["Service_Tier"])
    d["Priority_Weight"] = _priority_w(d["Priority"])
    d["Impact_Weight"] = d["Tier_Mult"] * d["Priority_Weight"]

    # Global aggregates
    total = len(d)
    if total == 0:
        total = 1

    avg_mttr = float(d["Duration_Hours"].mean())
    reopen_rate = float(d["Reopened_Flag"].mean())
    change_rate = float(d["Change_Related_Flag"].mean())

    # Service Resilience: shorter MTTR and lower reopen improves resilience
    mttr_score = 100.0 - min(100.0, _normalize_0_100(pd.Series([avg_mttr])).iloc[0])  # single value -> 0; handle separately below
    # If only single value, normalize gives 0 → we need a sensible scale:
    # Use a practical MTTR scale: 0h->100, 72h->0 (clamped)
    mttr_score = float((1.0 - min(avg_mttr / 72.0, 1.0)) * 100.0)

    service_resilience = (mttr_score * 0.65 + (1.0 - reopen_rate) * 100.0 * 0.35)

    # Change Governance: lower change-related incident rate = stronger governance
    change_governance = float((1.0 - min(change_rate * 1.25, 1.0)) * 100.0)

    # Structural Risk Debt: recurrence + reopen + high MTTR implies debt
    structural_risk_debt = float(
        (min(avg_mttr / 72.0, 1.0) * 40.0) + (reopen_rate * 35.0) + (min(change_rate * 1.1, 1.0) * 25.0)
    )
    structural_risk_debt = float(np.clip(structural_risk_debt, 0, 100))

    # Reliability Momentum: proxy for operational improvement potential:
    # for MVP, we interpret it as "how stable is the control layer" = inverse of combined instability signals
    instability = (min(avg_mttr / 72.0, 1.0) * 0.5) + (reopen_rate * 0.3) + (change_rate * 0.2)
    reliability_momentum = float((1.0 - min(instability, 1.0)) * 100.0)

    # BVSI: weighted overall view — matches your “tier-weighted” concept at a high level.
    # Keep it simple and explainable for execs.
    # Note: Structural Risk Debt is "bad" so invert it.
    bvsi = float(
        service_resilience * 0.35
        + change_governance * 0.30
        + (100.0 - structural_risk_debt) * 0.20
        + reliability_momentum * 0.15
    )
    bvsi = float(np.clip(bvsi, 0, 100))

    return {
        "Overall Service Resilience": round(float(service_resilience), 1),
        "Overall Change Governance": round(float(change_governance), 1),
        "Overall Structural Risk Debt": round(float(structural_risk_debt), 1),
        "Overall Reliability Momentum": round(float(reliability_momentum), 1),
        "BVSI": round(float(bvsi), 1),
    }


def posture_from_bvsi(bvsi: float) -> str:
    if bvsi >= 80:
        return "High Confidence Operations"
    if bvsi >= 60:
        return "Controlled but Exposed"
    if bvsi >= 40:
        return "Fragile Stability"
    return "Reactive Instability"


def build_sip_candidates(df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    """
    SIP candidates table: Service + Category theme with a priority score.
    """
    if df.empty:
        return pd.DataFrame()

    d = df.copy()
    d["Service"] = d["Service"].astype(str).fillna("Unknown")
    d["Service_Tier"] = d["Service_Tier"].astype(str).fillna("Tier 3")
    d["Category"] = d["Category"].astype(str).fillna("General")
    d["Priority"] = d["Priority"].astype(str).str.upper().fillna("P3")
    d["Reopened_Flag"] = pd.to_numeric(d["Reopened_Flag"], errors="coerce").fillna(0).clip(0, 1)
    d["Change_Related_Flag"] = pd.to_numeric(d["Change_Related_Flag"], errors="coerce").fillna(0).clip(0, 1)

    d["Duration_Hours"] = _duration_hours(d["Opened_Date"], d["Closed_Date"]).fillna(0)
    d["Tier_Mult"] = _tier_mult(d["Service_Tier"])
    d["Priority_Weight"] = _priority_w(d["Priority"])

    g = d.groupby(["Service", "Service_Tier", "Category"], as_index=False).agg(
        Count=("Service", "size"),
        Avg_MTTR=("Duration_Hours", "mean"),
        Reopen_Rate=("Reopened_Flag", "mean"),
        Change_Rate=("Change_Related_Flag", "mean"),
        Avg_Priority=("Priority_Weight", "mean"),
    )

    # SIP Priority Score (explainable)
    # High volume + high reopen + high MTTR + tier weighting
    g["SIP_Priority_Score"] = (
        _normalize_0_100(g["Count"]) * 0.35
        + _normalize_0_100(g["Avg_MTTR"]) * 0.25
        + (np.sqrt(g["Reopen_Rate"].clip(0, 1)) * 100.0) * 0.20
        + (np.sqrt(g["Change_Rate"].clip(0, 1)) * 100.0) * 0.10
        + _normalize_0_100(g["Avg_Priority"]) * 0.10
    ) * _tier_mult(g["Service_Tier"])

    g["SIP_Priority_Score"] = g["SIP_Priority_Score"].clip(0, 100)

    def label(score: float) -> str:
        if score >= 75:
            return "Immediate SIP"
        if score >= 55:
            return "Next SIP"
        if score >= 35:
            return "Monitor"
        return "Observe"

    g["Priority_Label"] = g["SIP_Priority_Score"].apply(lambda x: label(float(x)))

    def why_flagged(row) -> str:
        flags = []
        if row["Count"] >= g["Count"].quantile(0.75):
            flags.append("Recurrence/Volume")
        if row["Reopen_Rate"] >= max(0.10, float(g["Reopen_Rate"].quantile(0.75))):
            flags.append("Reopen Churn")
        if row["Avg_MTTR"] >= float(g["Avg_MTTR"].quantile(0.75)):
            flags.append("MTTR Drag")
        if row["Change_Rate"] >= max(0.10, float(g["Change_Rate"].quantile(0.75))):
            flags.append("Change Collision")
        return " + ".join(flags) if flags else "Emerging Pattern"

    g["Why_Flagged"] = g.apply(why_flagged, axis=1)

    out = g.rename(columns={"Category": "Suggested_Theme"}).sort_values("SIP_Priority_Score", ascending=False)

    out = out[
        [
            "Service",
            "Service_Tier",
            "Suggested_Theme",
            "SIP_Priority_Score",
            "Priority_Label",
            "Why_Flagged",
        ]
    ].head(top_n)

    out["SIP_Priority_Score"] = out["SIP_Priority_Score"].round(2)

    return out.reset_index(drop=True)


def build_analyst_review(overall: Dict[str, float], posture: str) -> str:
    bvsi = float(overall.get("BVSI", 0))
    sr = float(overall.get("Overall Service Resilience", 0))
    cg = float(overall.get("Overall Change Governance", 0))
    srd = float(overall.get("Overall Structural Risk Debt", 0))
    rm = float(overall.get("Overall Reliability Momentum", 0))

    # Executive-balanced, firm tone; “Your organization …”; light references to trust
    parts = []
    parts.append(
        f"Your organization demonstrates a <b>{posture}</b> operating posture (BVSI™ {bvsi:.1f})."
    )

    # Focus areas
    focus = []
    if srd >= 60:
        focus.append("structural risk debt is accumulating")
    if rm < 55:
        focus.append("reliability momentum is constrained")
    if sr < 55:
        focus.append("service resilience remains exposed")
    if cg >= 75:
        parts.append("Change governance appears comparatively strong, providing a stable control layer for improvement.")

    if focus:
        parts.append(
            "However, " + ", ".join(focus) + ", which can limit executive confidence and erode customer trust during disruption events."
        )
    else:
        parts.append("Overall stability signals are balanced, and targeted prevention work can further strengthen executive confidence.")

    parts.append("Recommended next step: prioritize SIPs against Tier-1 services where recurrence, MTTR drag, or reopen churn is concentrated.")
    return " ".join(parts)


def run_osil(df: pd.DataFrame) -> Dict[str, Any]:
    """
    Main entry point for the app.
    Returns a results dict used by Streamlit + report generation.
    """
    if df is None or df.empty:
        return {
            "overall": {},
            "posture": "Reactive Instability",
            "as_of": "",
            "sip_table": pd.DataFrame(),
            "analyst_review": "No data was provided.",
            "service_risk_df": pd.DataFrame(),
        }

    ok, missing = validate_required_columns(df)
    if not ok:
        raise ValueError(f"Missing required columns: {missing}")

    d = _safe_parse_dates(df)

    # Drop rows missing key dates (or they break MTTR)
    d = d.dropna(subset=["Opened_Date", "Closed_Date", "Service", "Service_Tier"]).copy()

    overall = compute_overall_scores(d)
    posture = posture_from_bvsi(float(overall.get("BVSI", 0)))

    # As-of date: max closed date
    as_of_dt = d["Closed_Date"].max()
    as_of = "" if pd.isna(as_of_dt) else str(as_of_dt.date())

    sip_table = build_sip_candidates(d, top_n=10)
    service_risk_df = compute_service_risk(d, top_n=10)
    analyst_review = build_analyst_review(overall, posture)

    return {
        "overall": overall,
        "posture": posture,
        "as_of": as_of,
        "sip_table": sip_table,
        "analyst_review": analyst_review,
        "service_risk_df": service_risk_df,
    }
