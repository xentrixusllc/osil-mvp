from __future__ import annotations

from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd


# ============================================================
# Canonical schemas
# ============================================================
INCIDENT_REQUIRED_COLUMNS = [
    "Service",
    "Service_Tier",
    "Opened_Date",
    "Priority",
]

INCIDENT_OPTIONAL_COLUMNS = [
    "Resolved_Date",
    "Closed_Date",
    "Reopened_Flag",
    "Category",
    "Change_Related_Flag",
    "Problem_ID",
]

CHANGE_REQUIRED_COLUMNS = [
    "Service",
    "Change_ID",
    "Change_Start",
]

CHANGE_OPTIONAL_COLUMNS = [
    "Change_End",
    "Change_Status",
    "Failed_Flag",
    "Risk",
    "Category",
]

PROBLEM_REQUIRED_COLUMNS = [
    "Service",
    "Problem_ID",
]

PROBLEM_OPTIONAL_COLUMNS = [
    "Opened_Date",
    "Resolved_Date",
    "Closed_Date",
    "State",
    "RCA_Completed_Flag",
    "Known_Error_Flag",
    "Priority",
    "Category",
]

REQUIRED_COLUMNS = INCIDENT_REQUIRED_COLUMNS.copy()


# ============================================================
# Column aliases
# ============================================================
INCIDENT_ALIASES = {
    "Service": ["Service", "service", "Business Service", "business_service", "Application", "application", "CI", "ci"],
    "Service_Tier": ["Service_Tier", "service_tier", "Tier", "tier", "Criticality", "criticality"],
    "Opened_Date": ["Opened_Date", "opened_date", "Opened", "opened", "Created", "created", "Created_At", "created_at"],
    "Resolved_Date": ["Resolved_Date", "resolved_date", "Resolved", "resolved", "Resolved_At", "resolved_at"],
    "Closed_Date": ["Closed_Date", "closed_date", "Closed", "closed", "Closed_At", "closed_at"],
    "Priority": ["Priority", "priority", "Severity", "severity"],
    "Reopened_Flag": ["Reopened_Flag", "reopened_flag", "Reopened", "reopened", "Reopen_Flag", "reopen_flag"],
    "Category": ["Category", "category", "Subcategory", "subcategory", "Type", "type"],
    "Change_Related_Flag": ["Change_Related_Flag", "change_related_flag", "RFC_Flag", "rfc_flag"],
    "Problem_ID": ["Problem_ID", "problem_id", "Problem", "problem", "Problem Number", "problem number"],
}

CHANGE_ALIASES = {
    "Service": ["Service", "service", "Business Service", "business_service", "Application", "application", "CI", "ci"],
    "Change_ID": ["Change_ID", "change_id", "Change", "change", "RFC", "rfc"],
    "Change_Start": [
        "Change_Start",
        "change_start",
        "Change_Start_Date",
        "change_start_date",
        "Start_Date",
        "start_date",
        "Planned_Start",
        "planned_start",
        "Implemented_Date",
        "implemented_date",
    ],
    "Change_End": [
        "Change_End",
        "change_end",
        "Change_End_Date",
        "change_end_date",
        "End_Date",
        "end_date",
        "Planned_End",
        "planned_end",
        "Completed_Date",
        "completed_date",
    ],
    "Change_Status": ["Change_Status", "change_status", "State", "state", "Status", "status"],
    "Failed_Flag": [
        "Failed_Flag",
        "failed_flag",
        "Failure_Flag",
        "failure_flag",
        "Change_Failed_Flag",
        "change_failed_flag",
        "Implementation_Success_Flag",
        "implementation_success_flag",
    ],
    "Risk": ["Risk", "risk", "Risk_Level", "risk_level"],
    "Category": ["Category", "category", "Type", "type", "Change_Type", "change_type"],
}

PROBLEM_ALIASES = {
    "Service": ["Service", "service", "Business Service", "business_service", "Application", "application", "CI", "ci"],
    "Problem_ID": ["Problem_ID", "problem_id", "Problem", "problem", "Problem Number", "problem number"],
    "Opened_Date": ["Opened_Date", "opened_date", "Opened", "opened", "Created", "created"],
    "Resolved_Date": ["Resolved_Date", "resolved_date", "Resolved", "resolved"],
    "Closed_Date": ["Closed_Date", "closed_date", "Closed", "closed"],
    "State": ["State", "state", "Status", "status"],
    "RCA_Completed_Flag": ["RCA_Completed_Flag", "rca_completed_flag", "RCA_Completed", "rca_completed"],
    "Known_Error_Flag": ["Known_Error_Flag", "known_error_flag", "Known_Error", "known_error"],
    "Priority": ["Priority", "priority", "Severity", "severity"],
    "Category": ["Category", "category", "Type", "type"],
}


# ============================================================
# Small helpers
# ============================================================
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
    s = series.copy()
    if s.dtype == bool:
        return s.astype(int)
    lowered = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "yes", "y", "t"}
    return lowered.isin(truthy).astype(int)


def normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").fillna(0.0)
    if len(s) == 0:
        return s
    mn = float(s.min())
    mx = float(s.max())
    if mx <= mn:
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - mn) / (mx - mn) * 100.0).round(1)


def _first_existing_col(df: pd.DataFrame, aliases: List[str]) -> Optional[str]:
    for c in aliases:
        if c in df.columns:
            return c
    return None


def _canonicalize_columns(df: Optional[pd.DataFrame], alias_map: Dict[str, List[str]]) -> pd.DataFrame:
    if df is None:
        return pd.DataFrame()

    out = df.copy()
    rename_map = {}

    for canonical, aliases in alias_map.items():
        found = _first_existing_col(out, aliases)
        if found and found != canonical:
            rename_map[found] = canonical

    out = out.rename(columns=rename_map)
    return out


def _parse_dates(df: pd.DataFrame, cols: List[str]) -> pd.DataFrame:
    out = df.copy()
    for c in cols:
        if c in out.columns:
            out[c] = pd.to_datetime(out[c], errors="coerce")
    return out


def validate_required_columns(df: pd.DataFrame, required: List[str]) -> Tuple[bool, List[str]]:
    missing = [c for c in required if c not in df.columns]
    return len(missing) == 0, missing


def calculate_data_readiness(
    incidents_df: pd.DataFrame,
    changes_df: Optional[pd.DataFrame] = None,
    problems_df: Optional[pd.DataFrame] = None,
) -> float:
    checks = []

    inc_cols = {str(c).strip().lower() for c in incidents_df.columns} if incidents_df is not None else set()
    chg_cols = {str(c).strip().lower() for c in changes_df.columns} if changes_df is not None and not changes_df.empty else set()
    prb_cols = {str(c).strip().lower() for c in problems_df.columns} if problems_df is not None and not problems_df.empty else set()

    checks.extend([
        any(c in inc_cols for c in {"opened_date", "opened", "created"}),
        any(c in inc_cols for c in {"resolved_date", "closed_date", "resolved", "closed"}),
        "priority" in inc_cols,
        any(c in inc_cols for c in {"service", "application", "ci", "business service", "business_service"}),
        any(c in inc_cols for c in {"reopened_flag", "reopened", "reopen_flag"}),
        any(c in inc_cols for c in {"service_tier", "tier", "criticality"}),
    ])

    if chg_cols:
        checks.extend([
            any(c in chg_cols for c in {"change_id", "change", "rfc"}),
            any(c in chg_cols for c in {"change_start", "change_start_date", "start_date", "planned_start", "implemented_date"}),
        ])

    if prb_cols:
        checks.extend([
            any(c in prb_cols for c in {"problem_id", "problem", "problem number"}),
            any(c in prb_cols for c in {"state", "status"}),
        ])

    if not checks:
        return 0.0

    return round((sum(bool(x) for x in checks) / len(checks)) * 100, 1)


def _normalize_service_anchor(df: pd.DataFrame, service_col: str = "Service") -> Tuple[pd.DataFrame, str]:
    out = df.copy()
    if service_col in out.columns:
        out["Service_Anchor"] = out[service_col].astype(str).str.strip().replace("", "Unknown")
        return out, service_col

    out["Service_Anchor"] = "Unknown"
    return out, "None"


def _choose_mttr_close_col(df: pd.DataFrame) -> Optional[str]:
    if "Resolved_Date" in df.columns and df["Resolved_Date"].notna().any():
        return "Resolved_Date"
    if "Closed_Date" in df.columns and df["Closed_Date"].notna().any():
        return "Closed_Date"
    return None


def _compute_mttr_hours(df: pd.DataFrame) -> pd.Series:
    open_col = "Opened_Date" if "Opened_Date" in df.columns else None
    close_col = _choose_mttr_close_col(df)
    if not open_col or not close_col:
        return pd.Series([np.nan] * len(df), index=df.index)

    hours = (df[close_col] - df[open_col]).dt.total_seconds() / 3600.0
    return hours.replace([np.inf, -np.inf], np.nan)


def _tier_weight(series: pd.Series) -> pd.Series:
    mapping = {
        "tier 1": 1.35,
        "tier 2": 1.15,
        "tier 3": 1.00,
        "tier 4": 0.90,
    }
    return series.astype(str).str.strip().str.lower().map(mapping).fillna(1.0)


def _priority_weight(priority: str) -> float:
    p = str(priority).strip().upper()
    mapping = {
        "P1": 1.50,
        "P2": 1.25,
        "P3": 1.00,
        "P4": 0.80,
        "P5": 0.60,
    }
    return mapping.get(p, 1.0)


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


def executive_interpretation(bvsi: float, posture: str, gap: str) -> str:
    if bvsi >= 70:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Core stability controls are visible, but additional attention to {gap} would improve resilience and executive confidence."
        )
    if bvsi >= 55:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Operational control exists, but recurring instability patterns still create exposure—especially across higher-impact services. "
            f"Focused SIP execution over the next 30–60 days will increase executive confidence and protect customer trust. Primary focus: {gap}."
        )
    return (
        f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
        f"Instability is materially visible and may be affecting reliability, customer confidence, and operating efficiency. "
        f"Immediate stabilization should focus first on {gap}."
    )


# ============================================================
# Change / Problem enrichment
# ============================================================
def _prepare_changes(changes_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    raw_df = changes_df.copy() if changes_df is not None else pd.DataFrame()
    df = _canonicalize_columns(changes_df, CHANGE_ALIASES)
    if df.empty:
        return df

    df = _parse_dates(df, ["Change_Start", "Change_End"])
    df, _ = _normalize_service_anchor(df, "Service")

    if "Change_Start" not in df.columns:
        raise ValueError("Missing required change column after alias mapping: Change_Start")

    if "Change_ID" not in df.columns:
        df["Change_ID"] = [f"CHG-{i+1}" for i in range(len(df))]

    if "Change_End" not in df.columns or not df["Change_End"].notna().any():
        df["Change_End"] = df["Change_Start"] + pd.Timedelta(hours=4)

    if "Failed_Flag" in df.columns:
        original_cols = {str(c) for c in raw_df.columns} if not raw_df.empty else set()
        if "Implementation_Success_Flag" in original_cols or "implementation_success_flag" in original_cols:
            success = _to_bool_series(df["Failed_Flag"])
            df["Failed_Flag"] = 1 - success
        else:
            df["Failed_Flag"] = _to_bool_series(df["Failed_Flag"])
    else:
        status = df["Change_Status"].astype(str).str.lower() if "Change_Status" in df.columns else pd.Series("", index=df.index)
        df["Failed_Flag"] = status.isin(["failed", "unsuccessful", "backed out"]).astype(int)

    return df


def _prepare_problems(problems_df: Optional[pd.DataFrame]) -> pd.DataFrame:
    df = _canonicalize_columns(problems_df, PROBLEM_ALIASES)
    if df.empty:
        return df

    df = _parse_dates(df, ["Opened_Date", "Resolved_Date", "Closed_Date"])
    df, _ = _normalize_service_anchor(df, "Service")

    if "RCA_Completed_Flag" in df.columns:
        df["RCA_Completed_Flag"] = _to_bool_series(df["RCA_Completed_Flag"])
    else:
        df["RCA_Completed_Flag"] = 0

    state = df["State"].astype(str).str.strip().str.lower() if "State" in df.columns else pd.Series("", index=df.index)
    df["Is_Open_Problem"] = ~state.isin(["closed", "resolved", "complete", "completed"])
    return df


def _enrich_incidents_with_change_collisions(inc: pd.DataFrame, changes: pd.DataFrame) -> pd.DataFrame:
    out = inc.copy()
    out["Change_Collision_Flag"] = out.get("Change_Related_Flag", 0)

    if changes is None or changes.empty:
        out["Change_Collision_Flag"] = pd.to_numeric(out["Change_Collision_Flag"], errors="coerce").fillna(0).clip(0, 1)
        return out

    if "Opened_Date" not in out.columns:
        out["Change_Collision_Flag"] = pd.to_numeric(out["Change_Collision_Flag"], errors="coerce").fillna(0).clip(0, 1)
        return out

    collision = pd.Series(0, index=out.index, dtype=int)

    change_groups = {svc: grp for svc, grp in changes.groupby("Service_Anchor", dropna=False)}
    for svc, inc_idx in out.groupby("Service_Anchor", dropna=False).groups.items():
        if svc not in change_groups:
            continue

        cg = change_groups[svc]
        starts = cg["Change_Start"].dropna().tolist()
        ends = cg["Change_End"].dropna().tolist()
        if not starts:
            continue

        for idx in inc_idx:
            opened = out.at[idx, "Opened_Date"]
            if pd.isna(opened):
                continue

            hit = False
            for st, en in zip(starts, ends):
                if pd.isna(st):
                    continue
                if pd.isna(en):
                    en = st + pd.Timedelta(hours=24)
                if st <= opened <= (en + pd.Timedelta(hours=24)):
                    hit = True
                    break
            collision.at[idx] = 1 if hit else 0

    base = pd.to_numeric(out.get("Change_Related_Flag", 0), errors="coerce").fillna(0).clip(0, 1).astype(int)
    out["Change_Collision_Flag"] = np.maximum(base, collision)
    return out


def _problem_gap_by_service(inc: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    services = sorted(set(inc["Service_Anchor"].dropna().astype(str).tolist()))
    if not services:
        return pd.DataFrame(columns=["Service_Anchor", "Problem_Gap_Risk"])

    if probs is None or probs.empty:
        return pd.DataFrame({"Service_Anchor": services, "Problem_Gap_Risk": [100.0] * len(services)})

    out_rows = []
    probs_by_service = probs.groupby("Service_Anchor", dropna=False)

    for svc in services:
        inc_svc = inc[inc["Service_Anchor"] == svc].copy()
        prob_svc = probs_by_service.get_group(svc).copy() if svc in probs_by_service.groups else pd.DataFrame()

        recurring_count = len(inc_svc)
        linked_problem_ratio = 0.0
        if "Problem_ID" in inc_svc.columns:
            linked_problem_ratio = inc_svc["Problem_ID"].notna().mean()

        open_problem_penalty = 0.0
        aged_problem_penalty = 0.0
        rca_penalty = 0.0

        if not prob_svc.empty:
            open_problem_penalty = min(float(prob_svc["Is_Open_Problem"].mean()) * 100.0, 100.0)
            if "Opened_Date" in prob_svc.columns and prob_svc["Opened_Date"].notna().any():
                age_days = (pd.Timestamp.utcnow().tz_localize(None) - prob_svc["Opened_Date"]).dt.days
                aged_problem_penalty = min(age_days.fillna(0).clip(lower=0).mean(), 100.0)
            if "RCA_Completed_Flag" in prob_svc.columns:
                rca_penalty = (1.0 - float(prob_svc["RCA_Completed_Flag"].mean())) * 100.0
        else:
            open_problem_penalty = 100.0
            aged_problem_penalty = 100.0
            rca_penalty = 100.0

        recurrence_factor = min(recurring_count * 8.0, 100.0)
        no_problem_penalty = (1.0 - linked_problem_ratio) * recurrence_factor

        risk = round(
            0.40 * no_problem_penalty +
            0.25 * open_problem_penalty +
            0.20 * aged_problem_penalty +
            0.15 * rca_penalty,
            1,
        )

        out_rows.append({"Service_Anchor": svc, "Problem_Gap_Risk": risk})

    return pd.DataFrame(out_rows)


# ============================================================
# Core aggregation
# ============================================================
def _prepare_incidents(incidents_df: pd.DataFrame) -> Tuple[pd.DataFrame, str]:
    df = _canonicalize_columns(incidents_df, INCIDENT_ALIASES)

    ok, missing = validate_required_columns(df, INCIDENT_REQUIRED_COLUMNS)
    if not ok:
        raise ValueError(f"Missing required incident columns: {missing}")

    for c in INCIDENT_OPTIONAL_COLUMNS:
        if c not in df.columns:
            df[c] = np.nan

    df = _parse_dates(df, ["Opened_Date", "Resolved_Date", "Closed_Date"])
    df, anchor_used = _normalize_service_anchor(df, "Service")

    df["Reopened_Flag"] = _to_bool_series(df["Reopened_Flag"]).fillna(0)
    df["Change_Related_Flag"] = _to_bool_series(df["Change_Related_Flag"]).fillna(0)
    df["Priority_Weight"] = df["Priority"].apply(_priority_weight)
    df["MTTR_Hours"] = _compute_mttr_hours(df)

    close_col = "Resolved_Date" if "Resolved_Date" in df.columns and df["Resolved_Date"].notna().any() else "Closed_Date"
    as_of_dt = df[close_col].max() if close_col in df.columns else pd.NaT
    df["__as_of_date"] = as_of_dt

    if "Category" not in df.columns:
        df["Category"] = "General"

    if "Service_Tier" not in df.columns:
        df["Service_Tier"] = "Tier 3"

    return df, anchor_used


def _build_rollup(inc: pd.DataFrame, changes: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    base = inc.groupby("Service_Anchor", dropna=False).agg(
        recurrence=("Service_Anchor", "count"),
        reopen_rate=("Reopened_Flag", "mean"),
        change_collision_rate=("Change_Collision_Flag", "mean"),
        mttr_hours=("MTTR_Hours", "mean"),
        tier=("Service_Tier", lambda x: x.value_counts().index[0] if len(x) else "Tier 3"),
        category=("Category", lambda x: x.value_counts().index[0] if len(x) else "General"),
    ).reset_index()

    if changes is not None and not changes.empty:
        chg_roll = changes.groupby("Service_Anchor", dropna=False).agg(
            change_count=("Change_ID", "nunique"),
            failed_change_rate=("Failed_Flag", "mean"),
        ).reset_index()
        base = base.merge(chg_roll, on="Service_Anchor", how="left")
    else:
        base["change_count"] = 0.0
        base["failed_change_rate"] = 0.0

    problem_gap = _problem_gap_by_service(inc, probs)
    base = base.merge(problem_gap, on="Service_Anchor", how="left")

    base["change_count"] = pd.to_numeric(base["change_count"], errors="coerce").fillna(0.0)
    base["failed_change_rate"] = pd.to_numeric(base["failed_change_rate"], errors="coerce").fillna(0.0)
    base["Problem_Gap_Risk"] = pd.to_numeric(base["Problem_Gap_Risk"], errors="coerce").fillna(0.0)

    base["tier_weight"] = _tier_weight(base["tier"])
    return base


def build_domain_scores(roll: pd.DataFrame) -> Dict[str, float]:
    if roll.empty:
        return {
            "Service Resilience": 0.0,
            "Change Governance": 0.0,
            "Structural Risk Debt™": 0.0,
            "Reliability Momentum": 0.0,
        }

    rec_norm = normalize_0_100(roll["recurrence"])
    mttr_norm = normalize_0_100(roll["mttr_hours"].fillna(roll["mttr_hours"].median() if roll["mttr_hours"].notna().any() else 0))
    reopen_norm = normalize_0_100(roll["reopen_rate"])
    chg_norm = normalize_0_100((roll["change_collision_rate"] * 100).fillna(0))
    problem_norm = normalize_0_100(roll["Problem_Gap_Risk"])

    service_resilience = float(np.clip(100 - (0.60 * mttr_norm.mean() + 0.40 * reopen_norm.mean()), 0, 100))
    change_governance = float(np.clip(100 - (0.65 * chg_norm.mean() + 0.35 * normalize_0_100(roll["failed_change_rate"] * 100).mean()), 0, 100))
    structural_risk_debt = float(np.clip(100 - (0.55 * rec_norm.mean() + 0.45 * problem_norm.mean()), 0, 100))
    reliability_momentum = float(np.clip(100 - (0.45 * rec_norm.mean() + 0.35 * mttr_norm.mean() + 0.20 * chg_norm.mean()), 0, 100))

    return {
        "Service Resilience": round(service_resilience, 1),
        "Change Governance": round(change_governance, 1),
        "Structural Risk Debt™": round(structural_risk_debt, 1),
        "Reliability Momentum": round(reliability_momentum, 1),
    }


def compute_bvsi(domain_scores: Dict[str, float]) -> float:
    vals = [float(v) for v in domain_scores.values()]
    if not vals:
        return 0.0
    return round(float(np.mean(vals)), 1)


def dominant_gap(domain_scores: Dict[str, float]) -> str:
    if not domain_scores:
        return "Service Resilience"
    return min(domain_scores.items(), key=lambda x: x[1])[0]


def build_service_risk_df(roll: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if roll.empty:
        return pd.DataFrame(columns=[
            "Service",
            "Service_Tier",
            "Recurrence_Risk",
            "MTTR_Drag_Risk",
            "Reopen_Churn_Risk",
            "Change_Collision_Risk",
            "Problem_Gap_Risk",
            "Total_Service_Risk",
        ])

    rec_norm = normalize_0_100(roll["recurrence"])
    mttr_norm = normalize_0_100(roll["mttr_hours"].fillna(roll["mttr_hours"].median() if roll["mttr_hours"].notna().any() else 0))
    reopen_norm = normalize_0_100(roll["reopen_rate"])
    chg_norm = normalize_0_100((roll["change_collision_rate"] * 100).fillna(0))
    problem_norm = normalize_0_100(roll["Problem_Gap_Risk"])

    df = pd.DataFrame(
        {
            "Service": roll["Service_Anchor"],
            "Service_Tier": roll["tier"],
            "Recurrence_Risk": rec_norm.round(1),
            "MTTR_Drag_Risk": mttr_norm.round(1),
            "Reopen_Churn_Risk": reopen_norm.round(1),
            "Change_Collision_Risk": chg_norm.round(1),
            "Problem_Gap_Risk": problem_norm.round(1),
        }
    )

    df["Total_Service_Risk"] = (
        0.35 * df["Recurrence_Risk"] +
        0.25 * df["MTTR_Drag_Risk"] +
        0.20 * df["Reopen_Churn_Risk"] +
        0.20 * df["Change_Collision_Risk"]
    ).round(1)

    return df.sort_values("Total_Service_Risk", ascending=False).head(top_n).reset_index(drop=True)


def build_sip_candidates(roll: pd.DataFrame, service_risk_df: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
    if roll.empty or service_risk_df.empty:
        return pd.DataFrame(columns=[
            "Service", "Service_Tier", "Suggested_Theme", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"
        ])

    merged = service_risk_df.merge(
        roll[["Service_Anchor", "category", "failed_change_rate"]],
        left_on="Service",
        right_on="Service_Anchor",
        how="left",
    )

    merged["SIP_Priority_Score"] = (
        merged["Total_Service_Risk"] * 0.70 +
        normalize_0_100(merged["Problem_Gap_Risk"].fillna(0)) * 0.20 +
        normalize_0_100((merged["failed_change_rate"].fillna(0) * 100)) * 0.10
    ).round(1)

    def label_sip(score: float) -> str:
        if score >= 70:
            return "Next SIP"
        if score >= 45:
            return "Monitor"
        return "Backlog"

    merged["Priority_Label"] = merged["SIP_Priority_Score"].apply(label_sip)
    merged["Suggested_Theme"] = merged["category"].fillna("Stability Improvement").astype(str)

    def why_flagged(row: pd.Series) -> str:
        parts = []
        if _safe_float(row.get("Recurrence_Risk", 0)) >= 60:
            parts.append("high recurrence")
        if _safe_float(row.get("MTTR_Drag_Risk", 0)) >= 60:
            parts.append("MTTR drag")
        if _safe_float(row.get("Reopen_Churn_Risk", 0)) >= 60:
            parts.append("reopen churn")
        if _safe_float(row.get("Change_Collision_Risk", 0)) >= 60:
            parts.append("change collision")
        if _safe_float(row.get("Problem_Gap_Risk", 0)) >= 60:
            parts.append("problem coverage gap")
        return " + ".join(parts) if parts else "tier exposure + multi-factor stability risk"

    merged["Why_Flagged"] = merged.apply(why_flagged, axis=1)

    out = merged[[
        "Service",
        "Service_Tier",
        "Suggested_Theme",
        "SIP_Priority_Score",
        "Priority_Label",
        "Why_Flagged",
    ]].head(top_n).reset_index(drop=True)

    return out


# ============================================================
# Public entrypoint
# ============================================================
def run_osil(
    incidents_df: pd.DataFrame,
    changes_df: Optional[pd.DataFrame] = None,
    problems_df: Optional[pd.DataFrame] = None,
) -> Dict[str, Any]:
    inc, anchor_used = _prepare_incidents(incidents_df)
    chg = _prepare_changes(changes_df)
    prb = _prepare_problems(problems_df)

    inc = _enrich_incidents_with_change_collisions(inc, chg)
    roll = _build_rollup(inc, chg, prb)

    domain_scores = build_domain_scores(roll)
    bvsi = compute_bvsi(domain_scores)
    posture = operating_posture(bvsi)
    gap = dominant_gap(domain_scores)

    as_of = date.today().isoformat()
    if "__as_of_date" in inc.columns and inc["__as_of_date"].notna().any():
        as_of = str(pd.to_datetime(inc["__as_of_date"].max()).date())

    exec_text = executive_interpretation(bvsi, posture, gap)
    service_risk_df = build_service_risk_df(roll, top_n=10)
    sip_view = build_sip_candidates(roll, service_risk_df, top_n=10)

    readiness_score = calculate_data_readiness(inc, chg, prb)

    practice_parts = ["INCIDENT"]
    if chg is not None and not chg.empty:
        practice_parts.append("CHANGE")
    if prb is not None and not prb.empty:
        practice_parts.append("PROBLEM")
    practice_type = " + ".join(practice_parts)

    return {
        "source_label": "",
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
        "top10": service_risk_df.head(10).copy(),
        "sip_view": sip_view.head(10).copy(),
        "tenant_name": "Default",
        "preview_df": inc.head(20).copy(),
        "incidents_enriched": inc,
        "changes_prepared": chg,
        "problems_prepared": prb,
    }
