import os
from difflib import get_close_matches
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

from osil_engine import INCIDENT_REQUIRED_COLUMNS, run_osil
from report_generator import build_osil_pdf_report

st.set_page_config(page_title="OSIL™ by Xentrixus", layout="wide")

APP_TITLE = "Xentrixus OSIL™ — Stability Intelligence MVP"
APP_SUB = "Upload Incident / Change / Problem exports → get BVSI™, Structural Risk Debt™, SIP priorities, and executive interpretation."

DEMO_INCIDENTS = "data/demo_incidents.csv"
DEMO_CHANGES = "data/demo_changes.csv"
DEMO_PROBLEMS = "data/demo_problems.csv"

INCIDENT_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor (Service / Application / CI / etc.)","required": True,"aliases": ["Service","Business Service","business_service","Application","Application Name","CI","Configuration Item","CI Name","Affected Service","Service Offering","Product","System","App","Application Service","Application CI","Configuration Item Name","Business Application","System Name","Product Name"]},
    "Service_Tier": {"label": "Service Tier (optional)","required": False,"aliases": ["Service_Tier","Tier","Criticality","Support Tier","Business Criticality"]},
    "Opened_Date": {"label": "Opened / Created Date","required": True,"aliases": ["Opened_Date","Opened","Created","Created_At","Open Date","Created Date","Opened At"]},
    "Resolved_Date": {"label": "Resolved Date (optional)","required": False,"aliases": ["Resolved_Date","Resolved","Resolved_At","Resolution Date","Resolved At"]},
    "Closed_Date": {"label": "Closed Date (optional)","required": False,"aliases": ["Closed_Date","Closed","Closed_At","Close Date","Closed At"]},
    "Priority": {"label": "Priority / Severity","required": True,"aliases": ["Priority","Severity","Impact/Priority","Incident Priority","Urgency"]},
    "Reopened_Flag": {"label": "Reopened Flag (optional)","required": False,"aliases": ["Reopened_Flag","Reopened","Reopen_Flag","Reopened?","Was Reopened"]},
    "Category": {"label": "Category / Type (optional)","required": False,"aliases": ["Category","Subcategory","Type","Assignment Category","Incident Category"]},
    "Change_Related_Flag": {"label": "Change-Related Flag (optional)","required": False,"aliases": ["Change_Related_Flag","RFC_Flag","Change Related","Was Change Related","Related To Change"]},
    "Problem_ID": {"label": "Problem ID / Link (optional)","required": False,"aliases": ["Problem_ID","Problem","Problem Number","Linked Problem","Parent Problem"]},
    "Resolution_Code": {"label": "Resolution Code (optional)","required": False,"aliases": ["Resolution_Code","Resolution Code","Close Code","Closure Code"]},
    "Closure_Code": {"label": "Closure Code (optional)","required": False,"aliases": ["Closure_Code","Closure Code","Close Code","Closing Code"]},
}

CHANGE_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor for Changes","required": True,"aliases": ["Service","Business Service","Application","Application Name","CI","Configuration Item","CI Name","Service Offering","Product","System","Application CI","Configuration Item Name","Business Application","System Name"]},
    "Change_ID": {"label": "Change ID / RFC (optional)","required": False,"aliases": ["Change_ID","Change","RFC","Change Number","RFC Number"]},
    "Change_Start": {"label": "Change Start Date","required": True,"aliases": ["Change_Start","Change_Start_Date","Start_Date","Planned_Start","Implemented_Date","Implementation Start","Start","Planned Start"]},
    "Change_End": {"label": "Change End Date (optional)","required": False,"aliases": ["Change_End","Change_End_Date","End_Date","Planned_End","Completed_Date","Implementation End","End","Planned End"]},
    "Change_Status": {"label": "Change Status (optional)","required": False,"aliases": ["Change_Status","Status","State","Implementation Status"]},
    "Failed_Flag": {"label": "Failure Flag or Success Flag (optional)","required": False,"aliases": ["Failed_Flag","Failure_Flag","Implementation_Success_Flag","Success_Flag","Successful Flag"]},
    "Risk": {"label": "Risk (optional)","required": False,"aliases": ["Risk","Risk_Level","Risk Level"]},
    "Category": {"label": "Change Category / Type (optional)","required": False,"aliases": ["Category","Type","Change_Type","Change Category"]},
    "Rollback_Flag": {"label": "Rollback Flag (optional)","required": False,"aliases": ["Rollback_Flag","Rolled_Back","Rollback","Backout Flag","Backed Out"]},
    "Change_Result": {"label": "Change Result / Outcome (optional)","required": False,"aliases": ["Change_Result","Outcome","Implementation Result","Result"]},
}

PROBLEM_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor for Problems","required": True,"aliases": ["Service","Business Service","Application","Application Name","CI","Configuration Item","CI Name","Service Offering","Product","System","Application CI","Configuration Item Name","Business Application","System Name"]},
    "Problem_ID": {"label": "Problem ID","required": True,"aliases": ["Problem_ID","Problem","Problem Number","PRB Number"]},
    "Opened_Date": {"label": "Problem Opened Date (optional)","required": False,"aliases": ["Opened_Date","Opened","Created","Opened At"]},
    "Resolved_Date": {"label": "Problem Resolved Date (optional)","required": False,"aliases": ["Resolved_Date","Resolved","Resolved At"]},
    "Closed_Date": {"label": "Problem Closed Date (optional)","required": False,"aliases": ["Closed_Date","Closed","Closed At"]},
    "State": {"label": "Problem State / Status (optional)","required": False,"aliases": ["State","Status","Problem State"]},
    "RCA_Completed_Flag": {"label": "RCA Completed Flag (optional)","required": False,"aliases": ["RCA_Completed_Flag","RCA_Completed","RCA Complete","RCA Done"]},
    "Known_Error_Flag": {"label": "Known Error Flag (optional)","required": False,"aliases": ["Known_Error_Flag","Known_Error","Known Error"]},
    "Root_Cause": {"label": "Root Cause / RCA Text (optional)","required": False,"aliases": ["Root_Cause","Root Cause","RCA","Cause","Cause Description","Problem Cause","Analysis Result"]},
    "Root_Cause_Category": {"label": "Root Cause Category (optional)","required": False,"aliases": ["Root_Cause_Category","Root Cause Category","Cause Category","RCA Category"]},
    "Workaround_Available": {"label": "Workaround Available (optional)","required": False,"aliases": ["Workaround_Available","Workaround","Workaround Available","Has Workaround"]},
    "Permanent_Fix_Flag": {"label": "Permanent Fix Flag (optional)","required": False,"aliases": ["Permanent_Fix_Flag","Permanent Fix","Fix Implemented","Permanent Resolution"]},
    "Priority": {"label": "Problem Priority (optional)","required": False,"aliases": ["Priority","Severity","Urgency"]},
    "Category": {"label": "Problem Category (optional)","required": False,"aliases": ["Category","Type","Problem Category"]},
}

# Keep the rest of your current app.py structure from the latest working version.
# Add these helper functions if not already present:
def _normalize_col_name(x: str) -> str:
    return str(x).strip().lower().replace("_", " ").replace("-", " ")

def _fuzzy_suggest(columns: List[str], aliases: List[str]) -> str:
    if not columns:
        return "-- None --"
    norm_to_original = {_normalize_col_name(c): c for c in columns}
    norm_columns = list(norm_to_original.keys())
    for alias in aliases:
        a = _normalize_col_name(alias)
        if a in norm_to_original:
            return norm_to_original[a]
    for alias in aliases:
        a = _normalize_col_name(alias)
        for nc in norm_columns:
            if a in nc or nc in a:
                return norm_to_original[nc]
    for alias in aliases:
        a = _normalize_col_name(alias)
        match = get_close_matches(a, norm_columns, n=1, cutoff=0.72)
        if match:
            return norm_to_original[match[0]]
    return "-- None --"

def _name_similarity_score(col_name: str, aliases: List[str]) -> float:
    norm = _normalize_col_name(col_name)
    alias_norms = [_normalize_col_name(a) for a in aliases]
    if norm in alias_norms:
        return 100.0
    if any((a in norm or norm in a) for a in alias_norms):
        return 85.0
    if get_close_matches(norm, alias_norms, n=1, cutoff=0.60):
        return 70.0
    return 0.0

def _cardinality_quality(series: pd.Series) -> float:
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return 0.0
    unique_ratio = s.nunique() / max(len(s), 1)
    if 0.02 <= unique_ratio <= 0.40:
        return 100.0
    if 0.40 < unique_ratio <= 0.70:
        return 75.0
    if 0.005 <= unique_ratio < 0.02:
        return 60.0
    if 0.70 < unique_ratio <= 0.90:
        return 45.0
    return 20.0

def _business_relevance_score(col_name: str) -> float:
    norm = _normalize_col_name(col_name)
    strong = ["service","application","system","product","platform","ci","configuration item"]
    weak = ["assignment group","group","state","status","priority","severity","category","location"]
    if any(token in norm for token in strong):
        return 100.0
    if any(token in norm for token in weak):
        return 20.0
    return 50.0

def _cross_file_overlap_score(base_series: pd.Series, other_series_list: List[pd.Series]) -> float:
    base = set(base_series.dropna().astype(str).str.strip())
    base = {x for x in base if x}
    if not base:
        return 0.0
    overlap_scores = []
    for s in other_series_list:
        vals = set(s.dropna().astype(str).str.strip())
        vals = {x for x in vals if x}
        if not vals:
            continue
        overlap = len(base.intersection(vals))
        denom = max(min(len(base), len(vals)), 1)
        overlap_scores.append((overlap / denom) * 100.0)
    return max(overlap_scores) if overlap_scores else 0.0

def _infer_anchor_candidates(df: pd.DataFrame, aliases: List[str], other_candidate_series=None, max_candidates: int = 3):
    if df is None or df.empty:
        return []
    scored = []
    for col in df.columns:
        total = (
            0.40 * _name_similarity_score(col, aliases)
            + 0.20 * _cardinality_quality(df[col])
            + 0.30 * _cross_file_overlap_score(df[col], other_candidate_series or [])
            + 0.10 * _business_relevance_score(col)
        )
        scored.append((col, round(total, 1)))
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[:max_candidates]

# In your existing mapping UI, keep manual override and use Service candidates from _infer_anchor_candidates.
