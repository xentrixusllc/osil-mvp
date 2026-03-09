# Full corrected osil_engine.py
# This version restores PDF-compatible output fields so the on-screen results and PDF align.

from __future__ import annotations

from datetime import date
from typing import Any, Dict, Optional, Tuple

import numpy as np
import pandas as pd


INCIDENT_REQUIRED_COLUMNS = [
    "Service",
    "Opened_Date",
    "Priority",
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
    s = series.copy()
    lowered = s.astype(str).str.strip().str.lower()
    truthy = {"1", "true", "yes", "y", "t"}
    return lowered.isin(truthy).astype(int)


def _first_non_null_mode(series: pd.Series, default: str) -> str:
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return default
    vc = s.value_counts()
    if vc.empty:
        return default
    return str(vc.index[0])


def _normalize_0_100(series: pd.Series) -> pd.Series:
    s = pd.to_numeric(series, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)
    if s.empty:
        return s
    mn = float(s.min())
    mx = float(s.max())
    if mx <= mn:
        return pd.Series([0.0] * len(s), index=s.index)
    return ((s - mn) / (mx - mn) * 100.0).round(1)


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


def _executive_interpretation(bvsi: float, posture: str, weakest_domain: str) -> str:
    if bvsi >= 85:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Operational control is broadly strong, though focused attention to {weakest_domain} would further strengthen resilience."
        )
    if bvsi >= 70:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Core controls are visible, but targeted improvement in {weakest_domain} would increase confidence and stability at scale."
        )
    if bvsi >= 55:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Operational control exists, but recurring instability patterns still create exposure. "
            f"Primary focus should be {weakest_domain} over the next 30–60 days."
        )
    if bvsi >= 40:
        return (
            f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
            f"Instability is visible across operational signals and may be affecting business confidence. "
            f"Immediate stabilization should focus first on {weakest_domain}."
        )
    return (
        f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). "
        f"Operational fragility is high and executive attention is warranted. "
        f"Priority should begin with {weakest_domain}."
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

    out["Service_Anchor"] = out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")
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

    close_col = "Resolved_Date" if out["Resolved_Date"].notna().any() else "Closed_Date"
    out["MTTR_Hours"] = (
        (out[close_col] - out["Opened_Date"]).dt.total_seconds() / 3600.0
        if close_col in out.columns else 0.0
    )
    out["MTTR_Hours"] = pd.to_numeric(out["MTTR_Hours"], errors="coerce").replace([np.inf, -np.inf], np.nan)

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

    return out, anchor_used


def _prepare_changes(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "Service_Anchor" not in out.columns:
        if "Service" in out.columns:
            out["Service_Anchor"] = out["Service"]
        else:
            out["Service_Anchor"] = "Unknown"

    out["Service_Anchor"] = out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")

    if "Change_Start" in out.columns:
        out["Change_Start"] = pd.to_datetime(out["Change_Start"], errors="coerce")
    else:
        out["Change_Start"] = pd.NaT

    if "Change_End" in out.columns:
        out["Change_End"] = pd.to_datetime(out["Change_End"], errors="coerce")
    else:
        out["Change_End"] = pd.NaT

    if "Change_ID" not in out.columns:
        out["Change_ID"] = [f"CHG-{i+1}" for i in range(len(out))]

    if "Failed_Flag" not in out.columns:
        out["Failed_Flag"] = 0
    out["Failed_Flag"] = _to_bool_series(out["Failed_Flag"])

    if "Category" not in out.columns:
        out["Category"] = "Change"

    return out


def _prepare_problems(df: Optional[pd.DataFrame]) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()

    out = df.copy()

    if "Service_Anchor" not in out.columns:
        if "Service" in out.columns:
            out["Service_Anchor"] = out["Service"]
        else:
            out["Service_Anchor"] = "Unknown"

    out["Service_Anchor"] = out["Service_Anchor"].astype(str).str.strip().replace("", "Unknown").fillna("Unknown")

    if "Problem_ID" not in out.columns:
        out["Problem_ID"] = [f"PRB-{i+1}" for i in range(len(out))]

    if "Opened_Date" in out.columns:
        out["Opened_Date"] = pd.to_datetime(out["Opened_Date"], errors="coerce")
    else:
        out["Opened_Date"] = pd.NaT

    if "State" not in out.columns:
        out["State"] = "Unknown"

    if "RCA_Completed_Flag" not in out.columns:
        out["RCA_Completed_Flag"] = 0
    out["RCA_Completed_Flag"] = _to_bool_series(out["RCA_Completed_Flag"])

    return out


def _detect_change_collision(inc: pd.DataFrame, changes: pd.DataFrame) -> pd.DataFrame:
    out = inc.copy()
    out["Change_Collision_Flag"] = out["Change_Related_Flag"].copy()

    if changes is None or changes.empty or "Change_Start" not in changes.columns:
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

    out["Change_Collision_Flag"] = pd.to_numeric(out["Change_Collision_Flag"], errors="coerce").fillna(0).clip(0, 1)
    return out


def _problem_gap_by_service(inc: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    services = pd.Series(inc["Service_Anchor"].dropna().unique(), name="Service_Anchor")
    if services.empty:
        return pd.DataFrame(columns=["Service_Anchor", "Problem_Gap_Risk"])

    rows = []
    for svc in services:
        inc_svc = inc[inc["Service_Anchor"] == svc]
        incident_count = len(inc_svc)
        linked_ratio = inc_svc["Problem_ID"].notna().mean() if "Problem_ID" in inc_svc.columns else 0.0

        if probs is not None and not probs.empty:
            prob_svc = probs[probs["Service_Anchor"] == svc]
        else:
            prob_svc = pd.DataFrame()

        if prob_svc.empty:
            if incident_count >= 5:
                risk = 85.0
            elif incident_count >= 3:
                risk = 70.0
            elif incident_count >= 1:
                risk = 45.0
            else:
                risk = 20.0
        else:
            open_penalty = 0.0
            if "State" in prob_svc.columns:
                state = prob_svc["State"].astype(str).str.lower()
                open_penalty = (~state.isin(["closed", "resolved", "complete", "completed"])).mean() * 40.0

            rca_penalty = 0.0
            if "RCA_Completed_Flag" in prob_svc.columns:
                rca_penalty = (1.0 - prob_svc["RCA_Completed_Flag"].mean()) * 30.0

            no_link_penalty = (1.0 - linked_ratio) * min(incident_count * 8.0, 30.0)
            risk = min(100.0, open_penalty + rca_penalty + no_link_penalty)

        rows.append({"Service_Anchor": svc, "Problem_Gap_Risk": round(float(risk), 1)})

    return pd.DataFrame(rows)


def _build_rollup(inc: pd.DataFrame, changes: pd.DataFrame, probs: pd.DataFrame) -> pd.DataFrame:
    base = inc.groupby("Service_Anchor", dropna=False).agg(
        recurrence=("Service_Anchor", "count"),
        reopen_rate=("Reopened_Flag", "mean"),
        change_collision_rate=("Change_Collision_Flag", "mean"),
        mttr_hours=("MTTR_Hours", "mean"),
        tier=("Service_Tier", lambda x: _first_non_null_mode(x, "Unspecified")),
        category=("Category", lambda x: _first_non_null_mode(x, "Stability Improvement")),
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
    base["mttr_hours"] = pd.to_numeric(base["mttr_hours"], errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)

    return base


def _build_service_risk_df(roll: pd.DataFrame, top_n: int = 10) -> pd.DataFrame:
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

    rec = _normalize_0_100(roll["recurrence"])
    mttr = _normalize_0_100(roll["mttr_hours"])
    reopen = _normalize_0_100(roll["reopen_rate"] * 100)
    change = _normalize_0_100(roll["change_collision_rate"] * 100)
    problem = _normalize_0_100(roll["Problem_Gap_Risk"])

    out = pd.DataFrame({
        "Service": roll["Service_Anchor"].astype(str),
        "Service_Tier": roll["tier"].fillna("Unspecified").astype(str),
        "Recurrence_Risk": rec.round(1),
        "MTTR_Drag_Risk": mttr.round(1),
        "Reopen_Churn_Risk": reopen.round(1),
        "Change_Collision_Risk": change.round(1),
        "Problem_Gap_Risk": problem.round(1),
    })

    out["Total_Service_Risk"] = (
        0.30 * out["Recurrence_Risk"]
        + 0.25 * out["MTTR_Drag_Risk"]
        + 0.15 * out["Reopen_Churn_Risk"]
        + 0.15 * out["Change_Collision_Risk"]
        + 0.15 * out["Problem_Gap_Risk"]
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
        100 - (0.65 * service_risk_df["MTTR_Drag_Risk"].mean() + 0.35 * service_risk_df["Reopen_Churn_Risk"].mean()),
        0,
        100,
    )
    change_governance = np.clip(
        100 - service_risk_df["Change_Collision_Risk"].mean(),
        0,
        100,
    )
    structural_risk_debt = np.clip(
        100 - (0.55 * service_risk_df["Recurrence_Risk"].mean() + 0.45 * service_risk_df["Problem_Gap_Risk"].mean()),
        0,
        100,
    )
    reliability_momentum = np.clip(
        100 - (
            0.40 * service_risk_df["Recurrence_Risk"].mean()
            + 0.30 * service_risk_df["MTTR_Drag_Risk"].mean()
            + 0.15 * service_risk_df["Reopen_Churn_Risk"].mean()
            + 0.15 * service_risk_df["Change_Collision_Risk"].mean()
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
        return pd.DataFrame(columns=[
            "Service", "Service_Tier", "Suggested_Theme", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"
        ])

    merged = service_risk_df.merge(
        roll[["Service_Anchor", "category"]],
        left_on="Service",
        right_on="Service_Anchor",
        how="left",
    )

    merged["SIP_Priority_Score"] = (
        0.70 * merged["Total_Service_Risk"]
        + 0.20 * merged["Problem_Gap_Risk"]
        + 0.10 * merged["Change_Collision_Risk"]
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
        if _safe_float(row.get("Reopen_Churn_Risk", 0)) >= 60:
            parts.append("reopen churn")
        if _safe_float(row.get("Change_Collision_Risk", 0)) >= 60:
            parts.append("change collision")
        if _safe_float(row.get("Problem_Gap_Risk", 0)) >= 60:
            parts.append("problem coverage gap")
        return " + ".join(parts) if parts else "multi-factor stability exposure"

    merged["Priority_Label"] = merged["SIP_Priority_Score"].apply(_label)
    merged["Suggested_Theme"] = merged["category"].fillna("Stability Improvement").astype(str)
    merged["Why_Flagged"] = merged.apply(_why, axis=1)

    return merged[[
        "Service",
        "Service_Tier",
        "Suggested_Theme",
        "SIP_Priority_Score",
        "Priority_Label",
        "Why_Flagged",
    ]].sort_values("SIP_Priority_Score", ascending=False).head(top_n).reset_index(drop=True)


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

    bvsi = round(float(np.mean(list(domain_scores.values()))), 1) if domain_scores else 0.0
    posture = _operating_posture(bvsi)
    weakest_domain = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Service Resilience"
    exec_text = _executive_interpretation(bvsi, posture, weakest_domain)

    sip_view = _build_sip_candidates(service_risk_df, roll, top_n=10)

    close_candidates = []
    if "Resolved_Date" in inc.columns and inc["Resolved_Date"].notna().any():
        close_candidates.append(inc["Resolved_Date"].max())
    if "Closed_Date" in inc.columns and inc["Closed_Date"].notna().any():
        close_candidates.append(inc["Closed_Date"].max())
    as_of = str(max(close_candidates).date()) if close_candidates else date.today().isoformat()

    readiness_checks = 0
    readiness_total = 6
    readiness_checks += 1 if "Service" in inc.columns or "Service_Anchor" in inc.columns else 0
    readiness_checks += 1 if "Opened_Date" in inc.columns else 0
    readiness_checks += 1 if "Priority" in inc.columns else 0
    readiness_checks += 1 if ("Resolved_Date" in inc.columns or "Closed_Date" in inc.columns) else 0
    readiness_checks += 1 if changes_df is not None and not changes_df.empty else 0
    readiness_checks += 1 if problems_df is not None and not problems_df.empty else 0
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
