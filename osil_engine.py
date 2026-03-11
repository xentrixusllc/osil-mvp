from __future__ import annotations
from datetime import date
from typing import Any, Dict, Optional, Tuple

import numpy as np
import pandas as pd

INCIDENT_REQUIRED_COLUMNS = ["Service", "Opened_Date", "Priority"]

def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if pd.isna(val):
            return default
        return float(val)
    except Exception:
        return default

def _to_bool_series(series: pd.Series) -> pd.Series:
    if series is None:
        return pd.Series(dtype=float)
    lowered = series.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "yes", "y", "t"}
    return lowered.isin(truthy).astype(int)

def _first_non_null_mode(series: pd.Series, default: str) -> str:
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return default
    vc = s.value_counts()
    return str(vc.index[0]) if not vc.empty else default

def _normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    if s.empty:
        return s
    mn = float(s.min())
    mx = float(s.max())
    if mx <= mn:
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - mn) / (mx - mn) * 100.0).round(1)

def _priority_weight(priority: Any) -> float:
    p = str(priority).strip().upper()
    mapping = {
        "P1": 1.50,
        "P2": 1.25,
        "P3": 1.00,
        "P4": 0.80,
        "P5": 0.60,
        "1": 1.50,
        "2": 1.25,
        "3": 1.00,
        "4": 0.80,
        "5": 0.60,
        "CRITICAL": 1.50,
        "HIGH": 1.25,
        "MEDIUM": 1.00,
        "LOW": 0.80,
    }
    return mapping.get(p, 1.0)

def _is_high_urgency(priority: Any) -> int:
    p = str(priority).strip().upper()
    if p in ["P1", "P2", "1", "2", "CRITICAL", "HIGH"]:
        return 1
    return 0

def _is_low_urgency(priority: Any) -> int:
    p = str(priority).strip().upper()
    if p in ["P3", "P4", "P5", "3", "4", "5", "MEDIUM", "LOW"]:
        return 1
    return 0

def _operating_posture(bvsi: float) -> str:
    if bvsi >= 85:
        return "High Confidence Operations"
    if bvsi >= 70:
        return "Controlled and Improving"
    if bvsi >= 55:
        return "Controlled but Exposed"
    if bvsi >= 40:
        return "Reactive and Exposed"
    return "Fragile Operations"

def _executive_interpretation(
    bvsi: float,
    posture: str,
    weakest_domain: str,
    practice_text: str,
) -> str:
    return (
        f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
        f"Current stability signals suggest the greatest exposure sits in {weakest_domain}. "
        f"{practice_text}"
    )

def _prepare_incidents(df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    if df is None or df.empty:
        raise ValueError("Incident dataset is empty.")
    
    out = df.copy()
    
    if "Service_Anchor" not in out.columns:
        if "Service" in out.columns:
            out["Service_Anchor"] = out["Service"]
            anchor_used = "Service"
        else:
            raise ValueError("Incident dataset requires an operational anchor mapped to Service.")
    else:
        anchor_used = "Service_Anchor"
        
    out["Service_Anchor"] = (
        out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")
    )
    
    if "Service" not in out.columns:
        out["Service"] = out["Service_Anchor"]
        
    out["Opened_Date"] = pd.to_datetime(out["Opened_Date"], errors="coerce")
    
    if "Resolved_Date" in out.columns:
        out["Resolved_Date"] = pd.to_datetime(out["Resolved_Date"], errors="coerce")
    else:
        out["Resolved_Date"] = pd.NaT
        
    if "Closed_Date" in out.columns:
        out["Closed_Date"] = pd.to_datetime(out["Closed_Date"], errors="coerce")
    else:
        out["Closed_Date"] = pd.NaT
        
    close_col = (
        "Resolved_Date"
        if isinstance(out["Resolved_Date"], pd.Series) and out["Resolved_Date"].notna().any()
        else "Closed_Date"
    )
    
    if close_col in out.columns:
        out["MTTR_Hours"] = (out[close_col] - out["Opened_Date"]).dt.total_seconds() / 3600.0
    else:
        out["MTTR_Hours"] = 0.0
        
    out["MTTR_Hours"] = (
        pd.to_numeric(out["MTTR_Hours"], errors="coerce")
        .replace([np.inf, -np.inf], np.nan)
        .fillna(0.0)
    )
    
    if "Reopened_Flag" not in out.columns:
        out["Reopened_Flag"] = 0
    out["Reopened_Flag"] = _to_bool_series(out["Reopened_Flag"])
    
    if "Service_Tier" not in out.columns:
        out["Service_Tier"] = "Unspecified"
    out["Service_Tier"] = out["Service_Tier"].fillna("Unspecified").astype(str)
    
    if "Category" not in out.columns:
        out["Category"] = "Stability Improvement"
    out["Category"] = out["Category"].fillna("Stability Improvement").astype(str)
    
    if "Change_Related_Flag" not in out.columns:
        out["Change_Related_Flag"] = 0
    out["Change_Related_Flag"] = _to_bool_series(out["Change_Related_Flag"])
    
    if "Problem_ID" not in out.columns:
        out["Problem_ID"] = np.nan
        
    out["Priority_Weight"] = out["Priority"].apply(_priority_weight)
    out["Is_High_Urgency"] = out["Priority"].apply(_is_high_urgency)
    out["Is_Low_Urgency"] = out["Priority"].apply(_is_low_urgency)
    
    return out, anchor_used

def _prepare_changes(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
        
    out = df.copy()
    
    if "Service_Anchor" not in out.columns:
        out["Service_Anchor"] = out["Service"] if "Service" in out.columns else "Unknown"
        
    out["Service_Anchor"] = (
        out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")
    )
    
    out["Change_Start"] = (
        pd.to_datetime(out["Change_Start"], errors="coerce")
        if "Change_Start" in out.columns
        else pd.NaT
    )
    
    out["Change_End"] = (
        pd.to_datetime(out["Change_End"], errors="coerce")
        if "Change_End" in out.columns
        else pd.NaT
    )
    
    if "Change_ID" not in out.columns:
        out["Change_ID"] = [f"CHG {i+1}" for i in range(len(out))]
        
    if "Failed_Flag" not in out.columns:
        out["Failed_Flag"] = 0
    out["Failed_Flag"] = _to_bool_series(out["Failed_Flag"])
    
    if "Rollback_Flag" not in out.columns:
        out["Rollback_Flag"] = 0
    out["Rollback_Flag"] = _to_bool_series(out["Rollback_Flag"])
    
    if "Category" not in out.columns:
        out["Category"] = "Change"
        
    return out

def _prepare_problems(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
        
    out = df.copy()
    
    if "Service_Anchor" not in out.columns:
        out["Service_Anchor"] = out["Service"] if "Service" in out.columns else "Unknown"
        
    out["Service_Anchor"] = (
        out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")
    )
    
    if "Problem_ID" not in out.columns:
        out["Problem_ID"] = [f"PRB {i+1}" for i in range(len(out))]
        
    out["Opened_Date"] = (
        pd.to_datetime(out["Opened_Date"], errors="coerce")
        if "Opened_Date" in out.columns
        else pd.NaT
    )
    
    out["Resolved_Date"] = (
        pd.to_datetime(out["Resolved_Date"], errors="coerce")
        if "Resolved_Date" in out.columns
        else pd.NaT
    )
    
    out["Closed_Date"] = (
        pd.to_datetime(out["Closed_Date"], errors="coerce")
        if "Closed_Date" in out.columns
        else pd.NaT
    )
    
    if "State" not in out.columns:
        out["State"] = "Unknown"
        
    if "RCA_Completed_Flag" not in out.columns:
        out["RCA_Completed_Flag"] = 0
    out["RCA_Completed_Flag"] = _to_bool_series(out["RCA_Completed_Flag"])
    
    if "Known_Error_Flag" not in out.columns:
        out["Known_Error_Flag"] = 0
    out["Known_Error_Flag"] = _to_bool_series(out["Known_Error_Flag"])
    
    if "Workaround_Available" not in out.columns:
        out["Workaround_Available"] = 0
    out["Workaround_Available"] = _to_bool_series(out["Workaround_Available"])
    
    if "Permanent_Fix_Flag" not in out.columns:
        out["Permanent_Fix_Flag"] = 0
    out["Permanent_Fix_Flag"] = _to_bool_series(out["Permanent_Fix_Flag"])
    
    if "Root_Cause_Text" not in out.columns:
        out["Root_Cause_Text"] = ""
        
    if "Contributing_Cause_Text" not in out.columns:
        out["Contributing_Cause_Text"] = ""
        
    return out

def _detect_change_collision(inc: pd.DataFrame, changes: pd.DataFrame) -> pd.DataFrame:
    out = inc.copy()
    out["Change_Collision_Flag"] = out["Change_Related_Flag"].copy()
    
    if changes is None or changes.empty:
        return out
        
    for _, change in changes.iterrows():
        svc = change.get("Service_Anchor", "Unknown")
        start = change.get("Change_Start", pd.NaT)
        end = change.get("Change_End", pd.NaT)
        
        if pd.isna(start):
            continue
            
        if pd.isna(end):
            end = start + pd.Timedelta(hours=6)
            
        mask = (
            (out["Service_Anchor"] == svc)
            & (out["Opened_Date"] >= start)
            & (out["Opened_Date"] <= end + pd.Timedelta(hours=12))
        )
        out.loc[mask, "Change_Collision_Flag"] = 1
        
    out["Change_Collision_Flag"] = (
        pd.to_numeric(out["Change_Collision_Flag"], errors="coerce").fillna(0).clip(0, 1)
    )
    return out

def _problem_signals_by_service(inc: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    services = pd.Series(inc["Service_Anchor"].dropna().unique(), name="Service_Anchor")
    rows = []
    
    for svc in services:
        inc_svc = inc[inc["Service_Anchor"] == svc]
        incident_count = len(inc_svc)
        linked_ratio = inc_svc["Problem_ID"].notna().mean() if "Problem_ID" in inc_svc.columns else 0.0
        
        prob_svc = (
            probs[probs["Service_Anchor"] == svc]
            if probs is not None and not probs.empty
            else pd.DataFrame()
        )
        
        if prob_svc.empty:
            rca_rate = known_error_rate = workaround_rate = permanent_fix_rate = open_problem_penalty = empty_rca_penalty = 0.0
        else:
            rca_rate = (
                prob_svc["RCA_Completed_Flag"].mean()
                if "RCA_Completed_Flag" in prob_svc.columns
                else 0.0
            )
            
            rca_text_fill_rate = 0.0
            if "Root_Cause_Text" in prob_svc.columns:
                 rca_text_filled = prob_svc["Root_Cause_Text"].replace("", np.nan).notna().sum()
                 rca_text_fill_rate = rca_text_filled / len(prob_svc) if len(prob_svc) > 0 else 0.0
                 
            empty_rca_penalty = max(0.0, rca_rate - rca_text_fill_rate)
            
            known_error_rate = (
                prob_svc["Known_Error_Flag"].mean()
                if "Known_Error_Flag" in prob_svc.columns
                else 0.0
            )
            
            workaround_rate = (
                prob_svc["Workaround_Available"].mean()
                if "Workaround_Available" in prob_svc.columns
                else 0.0
            )
            
            permanent_fix_rate = (
                prob_svc["Permanent_Fix_Flag"].mean()
                if "Permanent_Fix_Flag" in prob_svc.columns
                else 0.0
            )
            
            state = (
                prob_svc["State"].astype(str).str.lower()
                if "State" in prob_svc.columns
                else pd.Series(dtype=str)
            )
            
            open_problem_penalty = (
                (~state.isin(["closed", "resolved", "complete", "completed"])).mean()
                if not state.empty
                else 0.0
            )
            
        missing_problem_penalty = (1.0 - linked_ratio) * min(incident_count * 0.08, 1.0)
        missing_rca_penalty = 1.0 - rca_rate
        weak_fix_penalty = 1.0 - max(permanent_fix_rate, workaround_rate, known_error_rate)
        
        problem_gap_risk = (
            0.30 * missing_problem_penalty * 100.0
            + 0.20 * missing_rca_penalty * 100.0
            + 0.15 * empty_rca_penalty * 100.0
            + 0.20 * weak_fix_penalty * 100.0
            + 0.15 * open_problem_penalty * 100.0
        )
        
        rows.append(
            {
                "Service_Anchor": svc,
                "Problem_Gap_Risk": round(float(min(problem_gap_risk, 100.0)), 1),
                "RCA_Completion_Rate": round(float(rca_rate * 100.0), 1),
            }
        )
        
    return pd.DataFrame(rows)

def _extract_rca_themes(probs: pd.DataFrame) -> pd.DataFrame:
    """Extracts actual RCA text to map Structural Risk Debt themes."""
    if probs is None or probs.empty or "Root_Cause_Text" not in probs.columns:
        return pd.DataFrame(columns=["Service", "Problem_Count", "Documented_Themes"])

    valid = probs[probs["Root_Cause_Text"].astype(str).str.strip().replace("nan", "") != ""]
    if valid.empty:
        return pd.DataFrame(columns=["Service", "Problem_Count", "Documented_Themes"])

    grouped = valid.groupby("Service_Anchor").agg(
        Problem_Count=("Problem_ID", "nunique"),
        Documented_Themes=("Root_Cause_Text", lambda x: " | ".join(x.astype(str).str.strip().unique()[:3]))
    ).reset_index()
    
    return grouped.rename(columns={"Service_Anchor": "Service"}).sort_values("Problem_Count", ascending=False)

def _analyze_trust_gap(roll: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    """Analyzes the holistic P1 to P5 distribution to calculate the business trust gap."""
    if roll.empty or "high_urgency_count" not in roll.columns:
        return pd.DataFrame(), "Insufficient priority data to calculate business trust gap."

    df = roll[["Service_Anchor", "high_urgency_count", "low_urgency_count"]].copy()
    df.rename(columns={
        "Service_Anchor": "Service",
        "high_urgency_count": "Active_Disruption_P1_P2",
        "low_urgency_count": "Silent_Friction_P3_P5"
    }, inplace=True)

    total_crit = int(df["Active_Disruption_P1_P2"].sum())
    total_fric = int(df["Silent_Friction_P3_P5"].sum())

    narrative = (
        f"The Xentrixus OSIL™ framework identifies a perception gap when silent friction outweighs acknowledged disruption. "
        f"Currently, the business is absorbing {total_fric} low priority (P3 to P5) friction points compared to {total_crit} high priority (P1 and P2) disruptions. "
    )

    if total_crit == 0 and total_fric > 0:
         narrative += "This absolute imbalance creates severe risk; technology dashboards show zero disruption while the business bleeds productivity."
    elif total_fric > (total_crit * 4):
        narrative += "This massive ratio erodes IT credibility. Operational dashboards report green SLAs while the business experiences ongoing productivity loss."
    elif total_fric > 0:
        narrative += "This ratio indicates standard operational load, but services with concentrated low urgency friction must be monitored to protect business trust."
    else:
        narrative += "Incident volumes are currently contained with no material perception gap detected."

    df = df[(df["Active_Disruption_P1_P2"] > 0) | (df["Silent_Friction_P3_P5"] > 0)]
    df["Friction_Ratio"] = (df["Silent_Friction_P3_P5"] / (df["Active_Disruption_P1_P2"] + 1)).round(1)

    return df.sort_values("Friction_Ratio", ascending=False).head(5), narrative

def _build_rollup(inc: pd.DataFrame, changes: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    base = inc.groupby("Service_Anchor", dropna=False).agg(
        recurrence=("Service_Anchor", "count"),
        high_urgency_count=("Is_High_Urgency", "sum"),
        low_urgency_count=("Is_Low_Urgency", "sum"),
        reopen_rate=("Reopened_Flag", "mean"),
        change_collision_rate=("Change_Collision_Flag", "mean"),
        mttr_hours=("MTTR_Hours", "mean"),
        avg_priority_weight=("Priority_Weight", "mean"),
        tier=("Service_Tier", lambda x: _first_non_null_mode(x, "Unspecified")),
        category=("Category", lambda x: _first_non_null_mode(x, "Stability Improvement")),
    ).reset_index()
    
    if changes is not None and not changes.empty:
        chg_roll = changes.groupby("Service_Anchor", dropna=False).agg(
            change_count=("Change_ID", "nunique"),
            failed_change_rate=("Failed_Flag", "mean"),
            rollback_rate=("Rollback_Flag", "mean"),
        ).reset_index()
        base = base.merge(chg_roll, on="Service_Anchor", how="left")
    else:
        base["change_count"] = 0.0
        base["failed_change_rate"] = 0.0
        base["rollback_rate"] = 0.0
        
    problem_roll = _problem_signals_by_service(inc, probs)
    base = base.merge(problem_roll, on="Service_Anchor", how="left")
    
    for c in ["change_count", "failed_change_rate", "rollback_rate", "Problem_Gap_Risk", "RCA_Completion_Rate"]:
        base[c] = pd.to_numeric(base[c], errors="coerce").fillna(0.0)
        
    base["mttr_hours"] = pd.to_numeric(base["mttr_hours"], errors="coerce").fillna(0.0)
    return base

def _build_service_risk_df(roll: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if roll.empty:
        return pd.DataFrame(
            columns=[
                "Service",
                "Service_Tier",
                "Recurrence_Risk",
                "MTTR_Drag_Risk",
                "Reopen_Churn_Risk",
                "Change_Collision_Risk",
                "Problem_Gap_Risk",
                "Total_Service_Risk",
            ]
        )
        
    rec = _normalize_0_100(roll["recurrence"] * roll["avg_priority_weight"])
    mttr = _normalize_0_100(roll["mttr_hours"])
    reopen = _normalize_0_100(roll["reopen_rate"] * 100)
    change = _normalize_0_100(
        (roll["change_collision_rate"] * 70)
        + (roll["failed_change_rate"] * 20)
        + (roll["rollback_rate"] * 10)
    )
    problem = _normalize_0_100(roll["Problem_Gap_Risk"])
    
    out = pd.DataFrame(
        {
            "Service": roll["Service_Anchor"].astype(str),
            "Service_Tier": roll["tier"].fillna("Unspecified").astype(str),
            "Recurrence_Risk": rec.round(1),
            "MTTR_Drag_Risk": mttr.round(1),
            "Reopen_Churn_Risk": reopen.round(1),
            "Change_Collision_Risk": change.round(1),
            "Problem_Gap_Risk": problem.round(1),
        }
    )
    
    out["Total_Service_Risk"] = (
        0.28 * out["Recurrence_Risk"]
        + 0.24 * out["MTTR_Drag_Risk"]
        + 0.14 * out["Reopen_Churn_Risk"]
        + 0.17 * out["Change_Collision_Risk"]
        + 0.17 * out["Problem_Gap_Risk"]
    ).round(1)
    
    return out.sort_values("Total_Service_Risk", ascending=False).head(top_n).reset_index(drop=True)

def _build_domain_scores(service_risk_df: pd.DataFrame) -> Dict[str, float]:
    if service_risk_df.empty:
        return {
            "Service Resilience": 0.0,
            "Change Governance": 0.0,
            "Structural Risk Debt™": 0.0,
            "Reliability Momentum": 0.0,
        }
        
    service_resilience = np.clip(
        100 - (0.60 * service_risk_df["MTTR_Drag_Risk"].mean() + 0.40 * service_risk_df["Reopen_Churn_Risk"].mean()),
        0,
        100,
    )
    change_governance = np.clip(100 - service_risk_df["Change_Collision_Risk"].mean(), 0, 100)
    structural_risk_debt = np.clip(
        100 - (0.55 * service_risk_df["Recurrence_Risk"].mean() + 0.45 * service_risk_df["Problem_Gap_Risk"].mean()),
        0,
        100,
    )
    reliability_momentum = np.clip(
        100
        - (
            0.35 * service_risk_df["Recurrence_Risk"].mean()
            + 0.30 * service_risk_df["MTTR_Drag_Risk"].mean()
            + 0.15 * service_risk_df["Reopen_Churn_Risk"].mean()
            + 0.20 * service_risk_df["Change_Collision_Risk"].mean()
        ),
        0,
        100,
    )
    
    return {
        "Service Resilience": round(float(service_resilience), 1),
        "Change Governance": round(float(change_governance), 1),
        "Structural Risk Debt™": round(float(structural_risk_debt), 1),
        "Reliability Momentum": round(float(reliability_momentum), 1),
    }

def _build_sip_candidates(service_risk_df: pd.DataFrame, roll: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if service_risk_df.empty:
        return pd.DataFrame(
            columns=[
                "Service",
                "Service_Tier",
                "Suggested_Theme",
                "SIP_Priority_Score",
                "Priority_Label",
                "Why_Flagged",
            ]
        )
        
    merged = service_risk_df.merge(
        roll[["Service_Anchor", "category", "RCA_Completion_Rate"]],
        left_on="Service",
        right_on="Service_Anchor",
        how="left",
    )
    
    merged["SIP_Priority_Score"] = (
        0.70 * merged["Total_Service_Risk"]
        + 0.20 * merged["Problem_Gap_Risk"]
        + 0.10 * (100 - merged["RCA_Completion_Rate"].fillna(0))
    ).round(1)
    
    def _label(score: float) -> str:
        if score >= 70:
            return "Next SIP"
        if score >= 45:
            return "Monitor"
        return "Backlog"
        
    def _why(row: pd.Series) -> str:
        parts = []
        if _safe_float(row.get("Recurrence_Risk", 0)) >= 60:
            parts.append("high recurrence")
        if _safe_float(row.get("MTTR_Drag_Risk", 0)) >= 60:
            parts.append("MTTR drag")
        if _safe_float(row.get("Change_Collision_Risk", 0)) >= 60:
            parts.append("change instability")
        if _safe_float(row.get("Problem_Gap_Risk", 0)) >= 60:
            parts.append("structural learning gap")
        return " + ".join(parts) if parts else "multi factor stability exposure"
        
    merged["Priority_Label"] = merged["SIP_Priority_Score"].apply(_label)
    merged["Suggested_Theme"] = merged["category"].fillna("Stability Improvement").astype(str)
    merged["Why_Flagged"] = merged.apply(_why, axis=1)
    
    return (
        merged[
            [
                "Service",
                "Service_Tier",
                "Suggested_Theme",
                "SIP_Priority_Score",
                "Priority_Label",
                "Why_Flagged",
            ]
        ]
        .sort_values("SIP_Priority_Score", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )

def run_osil(
    incidents_df: pd.DataFrame,
    changes_df: Optional[pd.DataFrame] = None,
    problems_df: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    inc, anchor_used = _prepare_incidents(incidents_df)
    chg = _prepare_changes(changes_df)
    prb = _prepare_problems(problems_df)
    
    inc = _detect_change_collision(inc, chg)
    roll = _build_rollup(inc, chg, prb)
    service_risk_df = _build_service_risk_df(roll, top_n=10)
    domain_scores = _build_domain_scores(service_risk_df)
    
    rca_themes_df = _extract_rca_themes(prb)
    trust_gap_df, trust_gap_narrative = _analyze_trust_gap(roll)
    
    bvsi = round(float(np.mean(list(domain_scores.values()))), 1) if domain_scores else 0.0
    posture = _operating_posture(bvsi)
    weakest_domain = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Service Resilience"
    
    incident_signal = (
        "Incident restoration efficiency appears controlled."
        if domain_scores["Service Resilience"] >= 70
        else "Incident restoration signals show visible drag."
    )
    problem_signal = (
        "Structural learning signals are improving."
        if domain_scores["Structural Risk Debt™"] >= 70
        else "Structural learning signals remain inconsistent."
    )
    change_signal = (
        "Change governance appears steady."
        if domain_scores["Change Governance"] >= 70
        else "Change driven instability is contributing to exposure."
    )
    
    exec_text = _executive_interpretation(
        bvsi,
        posture,
        weakest_domain,
        f"{incident_signal} {problem_signal} {change_signal}",
    )
    
    sip_view = _build_sip_candidates(service_risk_df, roll, top_n=10)
    
    close_candidates = []
    if "Resolved_Date" in inc.columns and isinstance(inc["Resolved_Date"], pd.Series) and inc["Resolved_Date"].notna().any():
        close_candidates.append(inc["Resolved_Date"].max())
    if "Closed_Date" in inc.columns and isinstance(inc["Closed_Date"], pd.Series) and inc["Closed_Date"].notna().any():
        close_candidates.append(inc["Closed_Date"].max())
        
    as_of = str(max(close_candidates).date()) if close_candidates else date.today().isoformat()
    
    readiness_checks = 0
    readiness_total = 9
    readiness_checks += 1 if "Service" in inc.columns or "Service_Anchor" in inc.columns else 0
    readiness_checks += 1 if "Opened_Date" in inc.columns else 0
    readiness_checks += 1 if "Priority" in inc.columns else 0
    readiness_checks += 1 if ("Resolved_Date" in inc.columns or "Closed_Date" in inc.columns) else 0
    readiness_checks += 1 if "Reopened_Flag" in inc.columns else 0
    readiness_checks += 1 if changes_df is not None and not changes_df.empty else 0
    readiness_checks += 1 if problems_df is not None and not problems_df.empty else 0
    readiness_checks += 1 if problems_df is not None and not problems_df.empty and "Root_Cause_Text" in problems_df.columns else 0
    
    readiness_score = round((readiness_checks / readiness_total) * 100, 1)
    
    practice_parts = ["INCIDENT"]
    if changes_df is not None and not changes_df.empty:
        practice_parts.append("CHANGE")
    if problems_df is not None and not problems_df.empty:
        practice_parts.append("PROBLEM")
        
    return {
        "source_label": "",
        "practice_type": " + ".join(practice_parts),
        "anchor_used": anchor_used,
        "readiness_score": readiness_score,
        "bvsi": bvsi,
        "posture": posture,
        "gap": weakest_domain,
        "as_of": as_of,
        "exec_text": exec_text,
        "trust_gap_narrative": trust_gap_narrative,
        "trust_gap_df": trust_gap_df.copy(),
        "rca_themes_df": rca_themes_df.copy(),
        "domain_scores": domain_scores,
        "service_risk_df": service_risk_df.copy(),
        "top10": service_risk_df.copy(),
        "sip_view": sip_view.copy(),
        "tenant_name": "Default",
        "preview_df": inc.head(20).copy(),
        "incidents_enriched": inc.copy(),
        "changes_prepared": chg.copy(),
        "problems_prepared": prb.copy(),
    }
