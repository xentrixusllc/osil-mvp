
# ============================================================
# OSIL ENGINE — Multi‑Dataset Stability Intelligence Engine
# Supports Incident, Change, Problem datasets
# Safe handling for incomplete customer datasets
# ============================================================

import pandas as pd
import numpy as np
from datetime import datetime

# ============================================================
# Public constants used by app
# ============================================================

INCIDENT_REQUIRED_COLUMNS = [
    "Service",
    "Opened_Date",
    "Priority",
]

# ============================================================
# Utility helpers
# ============================================================

def _first_non_null_mode(series: pd.Series, default: str):
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return default
    vc = s.value_counts()
    if vc.empty:
        return default
    return str(vc.index[0])


def _safe_numeric(series):
    return pd.to_numeric(series, errors="coerce").fillna(0)


def _tier_weight(series):
    mapping = {
        "Tier 0": 1.5,
        "Tier 1": 1.3,
        "Tier 2": 1.1,
        "Tier 3": 1.0,
        "Tier 4": 0.9
    }
    return series.map(mapping).fillna(1.0)


# ============================================================
# Incident preparation
# ============================================================

def _prepare_incidents(df):

    df = df.copy()

    if "Service_Anchor" not in df.columns:
        if "Service" in df.columns:
            df["Service_Anchor"] = df["Service"]
        else:
            raise ValueError("Incident dataset requires a Service / Anchor column")

    df["Opened_Date"] = pd.to_datetime(df["Opened_Date"], errors="coerce")

    if "Resolved_Date" in df.columns:
        df["Resolved_Date"] = pd.to_datetime(df["Resolved_Date"], errors="coerce")

    if "Closed_Date" in df.columns:
        df["Closed_Date"] = pd.to_datetime(df["Closed_Date"], errors="coerce")

    if "Resolved_Date" in df.columns:
        df["MTTR_Hours"] = (df["Resolved_Date"] - df["Opened_Date"]).dt.total_seconds() / 3600
    else:
        df["MTTR_Hours"] = 0

    if "Reopened_Flag" not in df.columns:
        df["Reopened_Flag"] = 0

    if "Service_Tier" not in df.columns:
        df["Service_Tier"] = "Tier 3"

    if "Category" not in df.columns:
        df["Category"] = "General"

    if "Change_Related_Flag" not in df.columns:
        df["Change_Related_Flag"] = 0

    return df


# ============================================================
# Change preparation
# ============================================================

def _prepare_changes(df):

    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    if "Service_Anchor" not in df.columns:
        if "Service" in df.columns:
            df["Service_Anchor"] = df["Service"]
        else:
            df["Service_Anchor"] = "Unknown"

    if "Change_Start" in df.columns:
        df["Change_Start"] = pd.to_datetime(df["Change_Start"], errors="coerce")

    if "Change_End" in df.columns:
        df["Change_End"] = pd.to_datetime(df["Change_End"], errors="coerce")

    if "Change_ID" not in df.columns:
        df["Change_ID"] = ["CHG-"+str(i) for i in range(len(df))]

    if "Failed_Flag" not in df.columns:
        df["Failed_Flag"] = 0

    return df


# ============================================================
# Problem preparation
# ============================================================

def _prepare_problems(df):

    if df is None or df.empty:
        return pd.DataFrame()

    df = df.copy()

    if "Service_Anchor" not in df.columns:
        if "Service" in df.columns:
            df["Service_Anchor"] = df["Service"]
        else:
            df["Service_Anchor"] = "Unknown"

    if "Problem_ID" not in df.columns:
        df["Problem_ID"] = ["PRB-"+str(i) for i in range(len(df))]

    return df


# ============================================================
# Change collision detection
# ============================================================

def _detect_change_collision(inc, changes):

    if changes.empty:
        inc["Change_Collision_Flag"] = 0
        return inc

    inc["Change_Collision_Flag"] = 0

    for idx, change in changes.iterrows():

        start = change.get("Change_Start")
        end = change.get("Change_End")

        if pd.isna(start):
            continue

        if pd.isna(end):
            end = start + pd.Timedelta(hours=6)

        mask = (
            (inc["Service_Anchor"] == change["Service_Anchor"])
            & (inc["Opened_Date"] >= start)
            & (inc["Opened_Date"] <= end + pd.Timedelta(hours=12))
        )

        inc.loc[mask, "Change_Collision_Flag"] = 1

    return inc


# ============================================================
# Problem gap detection
# ============================================================

def _problem_gap_by_service(inc, probs):

    if probs.empty:
        return pd.DataFrame({
            "Service_Anchor": inc["Service_Anchor"].unique(),
            "Problem_Gap_Risk": 50
        })

    problem_services = probs["Service_Anchor"].unique()

    services = inc["Service_Anchor"].unique()

    rows = []

    for svc in services:

        incidents = inc[inc["Service_Anchor"] == svc]

        if svc not in problem_services and len(incidents) >= 3:
            risk = 80
        else:
            risk = 20

        rows.append({
            "Service_Anchor": svc,
            "Problem_Gap_Risk": risk
        })

    return pd.DataFrame(rows)


# ============================================================
# Build service rollup
# ============================================================

def _build_rollup(inc, changes, probs):

    base = inc.groupby("Service_Anchor", dropna=False).agg(

        recurrence=("Service_Anchor", "count"),
        reopen_rate=("Reopened_Flag", "mean"),
        change_collision_rate=("Change_Collision_Flag", "mean"),
        mttr_hours=("MTTR_Hours", "mean"),
        tier=("Service_Tier", lambda x: _first_non_null_mode(x, "Tier 3")),
        category=("Category", lambda x: _first_non_null_mode(x, "General"))

    ).reset_index()

    if not changes.empty:

        chg_roll = changes.groupby("Service_Anchor").agg(
            change_count=("Change_ID", "nunique"),
            failed_change_rate=("Failed_Flag", "mean")
        ).reset_index()

        base = base.merge(chg_roll, on="Service_Anchor", how="left")

    else:

        base["change_count"] = 0
        base["failed_change_rate"] = 0

    problem_gap = _problem_gap_by_service(inc, probs)

    base = base.merge(problem_gap, on="Service_Anchor", how="left")

    base["tier_weight"] = _tier_weight(base["tier"])

    return base


# ============================================================
# Risk scoring
# ============================================================

def _score_services(df):

    df["Recurrence_Risk"] = np.clip(df["recurrence"] * 5, 0, 100)
    df["MTTR_Drag_Risk"] = np.clip(df["mttr_hours"] / 4 * 10, 0, 100)
    df["Reopen_Churn_Risk"] = np.clip(df["reopen_rate"] * 100, 0, 100)
    df["Change_Collision_Risk"] = np.clip(df["change_collision_rate"] * 100, 0, 100)

    df["Total_Service_Risk"] = (
        df["Recurrence_Risk"] +
        df["MTTR_Drag_Risk"] +
        df["Reopen_Churn_Risk"] +
        df["Change_Collision_Risk"]
    ) / 4

    return df


# ============================================================
# Main OSIL run
# ============================================================

def run_osil(incidents_df, changes_df=None, problems_df=None):

    inc = _prepare_incidents(incidents_df)
    chg = _prepare_changes(changes_df)
    prb = _prepare_problems(problems_df)

    inc = _detect_change_collision(inc, chg)

    roll = _build_rollup(inc, chg, prb)

    roll = _score_services(roll)

    top10 = roll.sort_values("Total_Service_Risk", ascending=False).head(10)

    bvsi = round(100 - roll["Total_Service_Risk"].mean(), 1)

    if bvsi >= 80:
        posture = "Stable"
    elif bvsi >= 60:
        posture = "Controlled"
    elif bvsi >= 40:
        posture = "Reactive"
    else:
        posture = "Unstable"

    domain_scores = {
        "Recurrence Control": 100 - roll["Recurrence_Risk"].mean(),
        "Recovery Efficiency": 100 - roll["MTTR_Drag_Risk"].mean(),
        "Resolution Quality": 100 - roll["Reopen_Churn_Risk"].mean(),
        "Change Governance": 100 - roll["Change_Collision_Risk"].mean(),
    }

    sip_view = top10[["Service_Anchor","Total_Service_Risk"]].rename(
        columns={"Service_Anchor":"Service"}
    )

    return {

        "bvsi": bvsi,
        "posture": posture,
        "as_of": datetime.now().strftime("%Y-%m-%d"),
        "exec_text": "Operational stability analysis generated from incident, change, and problem signals.",
        "domain_scores": domain_scores,
        "service_risk_df": roll.rename(columns={"Service_Anchor":"Service"}),
        "top10": top10.rename(columns={"Service_Anchor":"Service"}),
        "sip_view": sip_view,
        "readiness_score": 80,
        "anchor_used": "Service / Application / CI",
        "practice_type": "Multi Dataset OSIL"
    }
