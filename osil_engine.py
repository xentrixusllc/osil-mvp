from __future__ import annotations
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

INCIDENT_REQUIRED_COLUMNS = [
    "Service", 
    "Opened_Date", 
    "Priority", 
    "State", 
    "Assignment_Group"
]

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
    if not vc.empty:
        return str(vc.index[0])
    return default

def _normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    if s.empty:
        return s
    mn = float(s.min())
    mx = float(s.max())
    if mx <= mn:
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - mn) / (mx - mn) * 100.0).round(1)

def _normalize_priority(priority: Any) -> str:
    p = str(priority).strip().upper()
    
    if "P1" in p or "CRIT" in p or "URGENT" in p or "EMERGENCY" in p:
        return "P1"
    if "P2" in p or "HIGH" in p:
        return "P2"
    if "P3" in p or "MED" in p or "MOD" in p:
        return "P3"
    if "P4" in p or "LOW" in p:
        return "P4"
    if "P5" in p or "MINOR" in p or "PLAN" in p:
        return "P5"
        
    if "1" in p:
        return "P1"
    if "2" in p:
        return "P2"
    if "3" in p:
        return "P3"
    if "4" in p:
        return "P4"
    if "5" in p:
        return "P5"
        
    return "P3"

def _priority_weight(priority: Any) -> float:
    p = _normalize_priority(priority)
    mapping = {
        "P1": 1.50,
        "P2": 1.25,
        "P3": 1.00,
        "P4": 0.80,
        "P5": 0.60,
    }
    return mapping.get(p, 1.0)

def _is_high_urgency(priority: Any) -> int:
    p = _normalize_priority(priority)
    if p in ["P1", "P2"]:
        return 1
    return 0

def _is_low_urgency(priority: Any) -> int:
    p = _normalize_priority(priority)
    if p in ["P3", "P4", "P5"]:
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

    if "State" in out.columns:
        invalid_states = ["canceled", "cancelled", "duplicate", "withdrawn", "rejected"]
        out = out[~out["State"].astype(str).str.lower().isin(invalid_states)].copy()

    if "Assignment_Group" not in out.columns:
        out["Assignment_Group"] = "Unassigned"
    
    out["Assignment_Group"] = out["Assignment_Group"].fillna("Unassigned").astype(str)
    out["Is_Assigned"] = (~out["Assignment_Group"].str.lower().isin(["unassigned", ""])).astype(int)

    if "Reassignment_Count" not in out.columns:
        out["Reassignment_Count"] = 0
    
    out["Reassignment_Count"] = pd.to_numeric(out["Reassignment_Count"], errors="coerce").fillna(0)
        
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

    if "Channel" not in out.columns:
        out["Channel"] = "Unknown"
    out["Channel"] = out["Channel"].fillna("Unknown").astype(str)
    
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
    
    if "Actual_Start" in out.columns:
        out["Change_Start"] = pd.to_datetime(out["Actual_Start"], errors="coerce")
    elif "Change_Start" in out.columns:
        out["Change_Start"] = pd.to_datetime(out["Change_Start"], errors="coerce")
    else:
        out["Change_Start"] = pd.NaT

    if "Actual_End" in out.columns:
        out["Change_End"] = pd.to_datetime(out["Actual_End"], errors="coerce")
    elif "Change_End" in out.columns:
        out["Change_End"] = pd.to_datetime(out["Change_End"], errors="coerce")
    else:
        out["Change_End"] = pd.NaT
    
    if "Change_ID" not in out.columns:
        out["Change_ID"] = [f"CHG {i+1}" for i in range(len(out))]

    if "Change_Type" not in out.columns:
        out["Change_Type"] = "Normal"
    out["Change_Type"] = out["Change_Type"].fillna("Normal").astype(str)
        
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
        
    if "Opened_Date" in out.columns:
        out["Opened_Date"] = pd.to_datetime(out["Opened_Date"], errors="coerce")
    else:
        out["Opened_Date"] = pd.NaT
    
    if "Resolved_Date" in out.columns:
        out["Resolved_Date"] = pd.to_datetime(out["Resolved_Date"], errors="coerce")
    else:
        out["Resolved_Date"] = pd.NaT
    
    if "Closed_Date" in out.columns:
        out["Closed_Date"] = pd.to_datetime(out["Closed_Date"], errors="coerce")
    else:
        out["Closed_Date"] = pd.NaT
    
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
    
    if "Root_Cause_Category" not in out.columns:
        out["Root_Cause_Category"] = ""
        
    if "Root_Cause_Text" not in out.columns:
        out["Root_Cause_Text"] = ""
        
    if "Contributing_Cause_Text" not in out.columns:
        out["Contributing_Cause_Text"] = ""

    if "Assignment_Group" not in out.columns:
        out["Assignment_Group"] = "Unassigned"
        
    return out

def _prepare_requests(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
        
    out = df.copy()
    if "Item" not in out.columns:
        out["Item"] = "General Request"
    out["Item"] = out["Item"].fillna("General Request").astype(str)
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

def _problem_signals_by_service(
    inc: pd.DataFrame, 
    probs: pd.DataFrame, 
    dynamic_cols: List[str] = None
) -> pd.DataFrame:
    
    services = pd.Series(inc["Service_Anchor"].dropna().unique(), name="Service_Anchor")
    rows = []
    
    for svc in services:
        inc_svc = inc[inc["Service_Anchor"] == svc]
        incident_count = len(inc_svc)
        
        if "Problem_ID" in inc_svc.columns:
            linked_ratio = inc_svc["Problem_ID"].notna().mean() 
        else:
            linked_ratio = 0.0
        
        if probs is not None and not probs.empty:
            prob_svc = probs[probs["Service_Anchor"] == svc]
        else:
            prob_svc = pd.DataFrame()
        
        rca_fidelity = 0.5 
        
        if prob_svc.empty:
            rca_rate = 0.0
            known_error_rate = 0.0
            workaround_rate = 0.0
            permanent_fix_rate = 0.0
            open_problem_penalty = 0.0
            empty_rca_penalty = 0.0
        else:
            if "RCA_Completed_Flag" in prob_svc.columns:
                rca_rate = prob_svc["RCA_Completed_Flag"].mean()
            else:
                rca_rate = 0.0
            
            rca_text_fill_rate = 0.0
            if "Root_Cause_Text" in prob_svc.columns:
                 rca_text_filled = prob_svc["Root_Cause_Text"].replace("", np.nan).notna().sum()
                 if len(prob_svc) > 0:
                     rca_text_fill_rate = rca_text_filled / len(prob_svc)
                 
            empty_rca_penalty = max(0.0, rca_rate - rca_text_fill_rate)
            
            if "Known_Error_Flag" in prob_svc.columns:
                known_error_rate = prob_svc["Known_Error_Flag"].mean()
            else:
                known_error_rate = 0.0
            
            if "Workaround_Available" in prob_svc.columns:
                workaround_rate = prob_svc["Workaround_Available"].mean()
            else:
                workaround_rate = 0.0
            
            if "Permanent_Fix_Flag" in prob_svc.columns:
                permanent_fix_rate = prob_svc["Permanent_Fix_Flag"].mean()
            else:
                permanent_fix_rate = 0.0
            
            if "State" in prob_svc.columns:
                state = prob_svc["State"].astype(str).str.lower()
            else:
                state = pd.Series(dtype=str)
            
            if not state.empty:
                open_problem_penalty = (~state.isin(["closed", "resolved", "complete", "completed"])).mean()
            else:
                open_problem_penalty = 0.0
            
            closed_mask = state.isin(["closed", "resolved", "complete", "completed"])
            closed_prbs = prob_svc[closed_mask]
            
            if not closed_prbs.empty:
                if "Root_Cause_Text" in closed_prbs.columns:
                    populated = closed_prbs["Root_Cause_Text"].astype(str).str.strip() != ""
                    rca_fidelity = populated.mean()
                else:
                    rca_fidelity = 0.0
            elif not prob_svc.empty:
                rca_fidelity = 0.5
                
            # TELEMETRY RICHNESS MODIFIER
            if dynamic_cols:
                valid_dynamic_cols = [c for c in dynamic_cols if c in prob_svc.columns]
                if valid_dynamic_cols:
                    fill_rate = prob_svc[valid_dynamic_cols].replace("", np.nan).notna().mean().mean()
                    # Boost fidelity by up to 25% for high operational discipline
                    rca_fidelity = min(1.0, rca_fidelity + (fill_rate * 0.25))
            
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
                "RCA_Fidelity": rca_fidelity, 
            }
        )
        
    return pd.DataFrame(rows)

def _extract_rca_themes(probs: pd.DataFrame, dynamic_cols: List[str] = None) -> pd.DataFrame:
    cols = ["Service", "Problem_Count", "Documented_Themes"]
    if dynamic_cols: 
        cols.extend(dynamic_cols)
        
    if probs is None or probs.empty or "Root_Cause_Text" not in probs.columns:
        return pd.DataFrame(columns=cols)

    valid = probs[probs["Root_Cause_Text"].astype(str).str.strip().replace("nan", "") != ""]
    if valid.empty:
        return pd.DataFrame(columns=cols)

    agg_dict = {
        "Problem_ID": "nunique",
        "Root_Cause_Text": lambda x: " | ".join(x.astype(str).str.strip().unique()[:3])
    }
    
    if dynamic_cols:
        for dc in dynamic_cols:
            if dc in valid.columns:
                agg_dict[dc] = lambda x: " | ".join(x.astype(str).str.strip().replace("nan", "").unique()[:3])

    grouped = valid.groupby("Service_Anchor").agg(agg_dict).reset_index()
    
    grouped.rename(
        columns={
            "Service_Anchor": "Service", 
            "Problem_ID": "Problem_Count", 
            "Root_Cause_Text": "Documented_Themes"
        }, 
        inplace=True
    )
    
    return grouped.sort_values("Problem_Count", ascending=False)

def _build_rca_pareto(probs: pd.DataFrame) -> pd.DataFrame:
    if probs is None or probs.empty:
        return pd.DataFrame(columns=["Theme", "Frequency", "Cumulative_Pct"])
        
    use_col = "Root_Cause_Text"
    if "Root_Cause_Category" in probs.columns:
        valid_cats = probs["Root_Cause_Category"].astype(str).str.strip().replace("nan", "")
        if valid_cats[valid_cats != ""].any():
            use_col = "Root_Cause_Category"

    if use_col not in probs.columns:
         return pd.DataFrame(columns=["Theme", "Frequency", "Cumulative_Pct"])
         
    valid = probs[probs[use_col].astype(str).str.strip().replace("nan", "") != ""]
    if valid.empty:
        return pd.DataFrame(columns=["Theme", "Frequency", "Cumulative_Pct"])
        
    valid.loc[:, "Theme"] = valid[use_col].astype(str).str.strip().str.title()
    
    pareto = valid.groupby("Theme").size().reset_index(name="Frequency")
    pareto = pareto.sort_values("Frequency", ascending=False).head(10).reset_index(drop=True)
    pareto["Cumulative_Pct"] = (pareto["Frequency"].cumsum() / pareto["Frequency"].sum() * 100).round(1)
    
    return pareto

def _analyze_trust_gap(roll: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    if roll.empty or "high_urgency_count" not in roll.columns:
        return pd.DataFrame(), "Insufficient priority data to calculate business trust gap."

    df = roll[["Service_Anchor", "high_urgency_count", "low_urgency_count"]].copy()
    df.rename(
        columns={
            "Service_Anchor": "Service",
            "high_urgency_count": "Active_Disruption_P1_P2",
            "low_urgency_count": "Silent_Friction_P3_P5"
        }, 
        inplace=True
    )

    total_crit = int(df["Active_Disruption_P1_P2"].sum())
    total_fric = int(df["Silent_Friction_P3_P5"].sum())

    narrative = (
        f"The Xentrixus OSIL™ framework identifies a perception gap when silent friction outweighs acknowledged disruption. "
        f"Currently, the business is absorbing {total_fric} low priority friction points compared to {total_crit} high priority disruptions. "
    )

    if total_crit == 0 and total_fric > 0:
         narrative += "This absolute imbalance creates severe risk; technology dashboards show zero disruption while the business bleeds productivity."
    elif total_fric > (total_crit * 4):
        narrative += "This massive ratio erodes IT credibility. Operational dashboards report green SLAs while the business experiences ongoing productivity loss."
    elif total_fric > 0:
        narrative += "This ratio indicates standard operational load, but services with concentrated low urgency friction must be monitored to protect business trust."
    else:
        narrative += "Incident volumes are currently contained with no material perception gap detected."

    df = df[(df["Active_Disruption_P1_P2"] > 0) | (df["Silent_Friction_P3_P5"] > 0)].copy()
    df["Friction_Ratio"] = (df["Silent_Friction_P3_P5"] / (df["Active_Disruption_P1_P2"] + 1)).round(1)

    return df.sort_values("Friction_Ratio", ascending=False).head(5), narrative

def _build_automation_strike_zone(inc: pd.DataFrame, req: pd.DataFrame) -> pd.DataFrame:
    rows = []
    
    if "Channel" in inc.columns:
        channel_series = inc["Channel"].astype(str).str.lower().str.strip()
        
        # 1. System Alerts
        alert_mask = channel_series.isin(["alert", "event", "system", "integration", "api"])
        alerts = inc[alert_mask]
        
        if not alerts.empty:
            alert_grouped = alerts.groupby("Service_Anchor").size().reset_index(name="Volume")
            alert_grouped = alert_grouped.sort_values("Volume", ascending=False).head(3)
            for _, r in alert_grouped.iterrows():
                rows.append({
                    "Target_Service": r["Service_Anchor"],
                    "Automation_Type": "Automated Remediation",
                    "Signal_Source": "System Alerts",
                    "Volume": r["Volume"],
                    "Wasted_Hours": round(r["Volume"] * 0.25, 1), 
                    "Mandate": "Implement AIOps capacity scaling or script automated service restarts."
                })
                
        # 2. Manual Toil
        manual_mask = ~alert_mask & (channel_series != "unknown") & (channel_series != "nan") & (channel_series != "")
        manuals = inc[manual_mask]
        
        if not manuals.empty:
            manual_grouped = manuals.groupby("Service_Anchor").size().reset_index(name="Volume")
            manual_grouped = manual_grouped.sort_values("Volume", ascending=False).head(3)
            for _, r in manual_grouped.iterrows():
                rows.append({
                    "Target_Service": r["Service_Anchor"],
                    "Automation_Type": "Tier 1 Deflection",
                    "Signal_Source": "Manual Incident Intake",
                    "Volume": r["Volume"],
                    "Wasted_Hours": round(r["Volume"] * 0.5, 1), 
                    "Mandate": "Deploy self-service workflows and conversational AI to deflect manual ticket creation."
                })
                
    if req is not None and not req.empty and "Item" in req.columns:
        req_grouped = req.groupby("Item").size().reset_index(name="Volume")
        req_grouped = req_grouped.sort_values("Volume", ascending=False).head(3)
        for _, r in req_grouped.iterrows():
            rows.append({
                "Target_Service": r["Item"],
                "Automation_Type": "Fulfillment Orchestration",
                "Signal_Source": "Service Requests",
                "Volume": r["Volume"],
                "Wasted_Hours": round(r["Volume"] * 0.5, 1), 
                "Mandate": "Build zero touch orchestration scripts to fulfill automatically."
            })
            
    df = pd.DataFrame(rows)
    if df.empty:
        return pd.DataFrame(
            columns=[
                "Target_Service", 
                "Automation_Type", 
                "Signal_Source", 
                "Volume", 
                "Wasted_Hours", 
                "Mandate"
            ]
        )
        
    return df.sort_values("Wasted_Hours", ascending=False).head(6)

def _build_rollup(
    inc: pd.DataFrame, 
    changes: pd.DataFrame, 
    probs: pd.DataFrame, 
    dynamic_cols: List[str] = None
) -> pd.DataFrame:
    
    base = inc.groupby("Service_Anchor", dropna=False).agg(
        recurrence=("Service_Anchor", "count"),
        high_urgency_count=("Is_High_Urgency", "sum"),
        low_urgency_count=("Is_Low_Urgency", "sum"),
        reopen_rate=("Reopened_Flag", "mean"),
        change_collision_rate=("Change_Collision_Flag", "mean"),
        mttr_hours=("MTTR_Hours", "mean"),
        reassignment_mean=("Reassignment_Count", "mean"),
        avg_priority_weight=("Priority_Weight", "mean"),
        tier=("Service_Tier", lambda x: _first_non_null_mode(x, "Unspecified")),
        category=("Category", lambda x: _first_non_null_mode(x, "Stability Improvement")),
        assignment_hygiene=("Is_Assigned", "mean"), 
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
        
    problem_roll = _problem_signals_by_service(inc, probs, dynamic_cols)
    base = base.merge(problem_roll, on="Service_Anchor", how="left")
    
    cols_to_fill = [
        "change_count", 
        "failed_change_rate", 
        "rollback_rate", 
        "Problem_Gap_Risk", 
        "RCA_Completion_Rate"
    ]
    
    for c in cols_to_fill:
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
                "Execution_Churn_Risk",
                "Reopen_Churn_Risk",
                "Change_Collision_Risk",
                "Problem_Gap_Risk",
                "Active_Disruption_P1_P2",
                "Data_Confidence",
                "Total_Service_Risk",
            ]
        )
        
    rec = _normalize_0_100(roll["recurrence"] * roll["avg_priority_weight"])
    mttr = _normalize_0_100(roll["mttr_hours"])
    reopen = _normalize_0_100(roll["reopen_rate"] * 100)
    exec_churn = np.clip(roll["reassignment_mean"] * 33.3, 0, 100) 
    
    change = _normalize_0_100(
        (roll["change_collision_rate"] * 70)
        + (roll["failed_change_rate"] * 20)
        + (roll["rollback_rate"] * 10)
    )
    problem = _normalize_0_100(roll["Problem_Gap_Risk"])
    
    if "RCA_Fidelity" not in roll.columns:
        roll["RCA_Fidelity"] = 0.5
    if "assignment_hygiene" not in roll.columns:
        roll["assignment_hygiene"] = 1.0

    conf_score = (roll["assignment_hygiene"].fillna(1.0) * 50.0) + (roll["RCA_Fidelity"].fillna(0.5) * 50.0)
    
    def _get_conf_label(s: float) -> str:
        if s >= 85.0:
            return "High"
        elif s >= 60.0:
            return "Medium"
        return "Low"
        
    data_confidence = conf_score.apply(_get_conf_label)
    
    out = pd.DataFrame(
        {
            "Service": roll["Service_Anchor"].astype(str),
            "Service_Tier": roll["tier"].fillna("Unspecified").astype(str),
            "Recurrence_Risk": rec.round(1),
            "MTTR_Drag_Risk": mttr.round(1),
            "Execution_Churn_Risk": exec_churn.round(1),
            "Reopen_Churn_Risk": reopen.round(1),
            "Change_Collision_Risk": change.round(1),
            "Problem_Gap_Risk": problem.round(1),
            "Active_Disruption_P1_P2": roll["high_urgency_count"].fillna(0).astype(int) if "high_urgency_count" in roll.columns else 0,
            "Data_Confidence": data_confidence 
        }
    )
    
    out["Total_Service_Risk"] = (
        0.25 * out["Recurrence_Risk"]
        + 0.15 * out["MTTR_Drag_Risk"]
        + 0.15 * out["Execution_Churn_Risk"]
        + 0.10 * out["Reopen_Churn_Risk"]
        + 0.15 * out["Change_Collision_Risk"]
        + 0.20 * out["Problem_Gap_Risk"]
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
        100 - (
            0.40 * service_risk_df["MTTR_Drag_Risk"].mean() 
            + 0.30 * service_risk_df["Reopen_Churn_Risk"].mean() 
            + 0.30 * service_risk_df["Execution_Churn_Risk"].mean()
        ),
        0,
        100,
    )
    
    change_governance = np.clip(
        100 - service_risk_df["Change_Collision_Risk"].mean(), 
        0, 
        100
    )
    
    structural_risk_debt = np.clip(
        100 - (
            0.55 * service_risk_df["Recurrence_Risk"].mean() 
            + 0.45 * service_risk_df["Problem_Gap_Risk"].mean()
        ),
        0,
        100,
    )
    
    reliability_momentum = np.clip(
        100
        - (
            0.35 * service_risk_df["Recurrence_Risk"].mean()
            + 0.25 * service_risk_df["MTTR_Drag_Risk"].mean()
            + 0.10 * service_risk_df["Execution_Churn_Risk"].mean()
            + 0.10 * service_risk_df["Reopen_Churn_Risk"].mean()
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
                "Data_Hygiene_Check",
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
        if _safe_float(row.get("Execution_Churn_Risk", 0)) >= 60:
            parts.append("high assignment churn")
        elif _safe_float(row.get("MTTR_Drag_Risk", 0)) >= 60:
            parts.append("MTTR drag")
        if _safe_float(row.get("Change_Collision_Risk", 0)) >= 60:
            parts.append("change instability")
        if _safe_float(row.get("Problem_Gap_Risk", 0)) >= 60:
            parts.append("structural learning gap")
            
        if parts:
            return " + ".join(parts)
        return "multi factor stability exposure"
        
    merged["Priority_Label"] = merged["SIP_Priority_Score"].apply(_label)
    merged["Suggested_Theme"] = merged["category"].fillna("Stability Improvement").astype(str)
    merged["Why_Flagged"] = merged.apply(_why, axis=1)
    merged["Data_Hygiene_Check"] = merged["Data_Confidence"] 
    
    return (
        merged[
            [
                "Service",
                "Service_Tier",
                "Suggested_Theme",
                "SIP_Priority_Score",
                "Priority_Label",
                "Why_Flagged",
                "Data_Hygiene_Check",
            ]
        ]
        .sort_values("SIP_Priority_Score", ascending=False)
        .head(top_n)
        .reset_index(drop=True)
    )

def _build_svs_scores(domain_scores: Dict[str, float], readiness_score: float) -> Dict[str, float]:
    """Translates technical OSIL domain scores into ITIL 4 Service Value System Health Scores"""
    
    # 1. Governance (Change & Release Control)
    gov = domain_scores.get("Change Governance", 0.0)
    
    # 2. Continual Improvement (Proactive Elimination of Debt)
    ci = domain_scores.get("Structural Risk Debt™", 0.0)
    
    # 3. Practices (Execution, Process Hygiene, Resilience)
    prac = round((domain_scores.get("Service Resilience", 0.0) + readiness_score) / 2.0, 1)
    
    # 4. Guiding Principles (Focus on Value / Trust Gap / Momentum)
    gp = domain_scores.get("Reliability Momentum", 0.0)
    
    return {
        "Governance": gov,
        "Continual Improvement": ci,
        "Practices": prac,
        "Guiding Principles": gp
    }

def _build_svc_scores(domain_scores: Dict[str, float], readiness_score: float) -> Dict[str, float]:
    """Translates technical OSIL domain scores into ITIL 4 Service Value Chain Scores"""
    
    ds = domain_scores.get("Service Resilience", 0.0)
    dt = domain_scores.get("Change Governance", 0.0)
    imp = domain_scores.get("Structural Risk Debt™", 0.0)
    eng = domain_scores.get("Reliability Momentum", 0.0)
    plan = readiness_score
    ob = round((dt + ds) / 2.0, 1)
    
    return {
        "Plan": plan,
        "Improve": imp,
        "Engage": eng,
        "Design & Transition": dt,
        "Obtain/Build": ob,
        "Deliver & Support": ds
    }

def run_osil(
    incidents_df: pd.DataFrame,
    changes_df: Optional[pd.DataFrame] = None,
    problems_df: Optional[pd.DataFrame] = None,
    requests_df: Optional[pd.DataFrame] = None,
    dynamic_prb_cols: List[str] = None
) -> Dict[str, Any]:
    
    inc, anchor_used = _prepare_incidents(incidents_df)
    chg = _prepare_changes(changes_df)
    prb = _prepare_problems(problems_df)
    req = _prepare_requests(requests_df)
    
    inc = _detect_change_collision(inc, chg)
    roll = _build_rollup(inc, chg, prb, dynamic_prb_cols)
    service_risk_df = _build_service_risk_df(roll, top_n=10)
    domain_scores = _build_domain_scores(service_risk_df)
    
    rca_themes_df = _extract_rca_themes(prb, dynamic_prb_cols)
    rca_pareto_df = _build_rca_pareto(prb)
    trust_gap_df, trust_gap_narrative = _analyze_trust_gap(roll)
    automation_df = _build_automation_strike_zone(inc, req)
    
    if domain_scores:
        bvsi = round(float(np.mean(list(domain_scores.values()))), 1)
    else:
        bvsi = 0.0
        
    posture = _operating_posture(bvsi)
    
    if domain_scores:
        weakest_domain = min(domain_scores.items(), key=lambda x: x[1])[0]
    else:
        weakest_domain = "Service Resilience"
    
    if domain_scores.get("Service Resilience", 0) >= 70:
        incident_signal = "Incident restoration efficiency appears controlled."
    else:
        incident_signal = "Incident restoration signals show visible drag and execution churn."
        
    if domain_scores.get("Structural Risk Debt™", 0) >= 70:
        problem_signal = "Structural learning signals are improving."
    else:
        problem_signal = "Structural learning signals remain inconsistent."
        
    if domain_scores.get("Change Governance", 0) >= 70:
        change_signal = "Change governance appears steady."
    else:
        change_signal = "Change driven instability is contributing to exposure."
    
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
        
    if close_candidates:
        as_of = str(max(close_candidates).date())
    else:
        as_of = date.today().isoformat()
    
    readiness_checks = 0
    readiness_total = 11
    
    if "Service" in inc.columns or "Service_Anchor" in inc.columns:
        readiness_checks += 1
    if "Opened_Date" in inc.columns:
        readiness_checks += 1
    if "Priority" in inc.columns:
        readiness_checks += 1
    if "State" in inc.columns:
        readiness_checks += 1
    if "Assignment_Group" in inc.columns:
        readiness_checks += 1
        
    if changes_df is not None and not changes_df.empty:
        readiness_checks += 1
        if "Change_Type" in changes_df.columns:
            readiness_checks += 1
            
    if problems_df is not None and not problems_df.empty:
        readiness_checks += 1
        if "Root_Cause_Text" in problems_df.columns:
            readiness_checks += 1
        if "Assignment_Group" in problems_df.columns:
            readiness_checks += 1
            
    if requests_df is not None and not requests_df.empty:
        readiness_checks += 1
    
    readiness_score = round((readiness_checks / readiness_total) * 100, 1)
    
    svs_scores = _build_svs_scores(domain_scores, readiness_score)
    svc_scores = _build_svc_scores(domain_scores, readiness_score)
    
    practice_parts = ["INCIDENT"]
    if changes_df is not None and not changes_df.empty:
        practice_parts.append("CHANGE")
    if problems_df is not None and not problems_df.empty:
        practice_parts.append("PROBLEM")
    if requests_df is not None and not requests_df.empty:
        practice_parts.append("REQUEST")
        
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
        "rca_pareto_df": rca_pareto_df.copy(),
        "domain_scores": domain_scores,
        "svs_scores": svs_scores,
        "svc_scores": svc_scores,
        "service_risk_df": service_risk_df.copy(),
        "top10": service_risk_df.copy(),
        "sip_view": sip_view.copy(),
        "automation_df": automation_df.copy(),
        "tenant_name": "Default",
        "preview_df": inc.head(20).copy(),
    }
