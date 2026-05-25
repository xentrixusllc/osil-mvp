"""OSIL Streamlit Application - Executive Command Center"""
import os
import textwrap
from difflib import get_close_matches
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

from osil_engine import INCIDENT_REQUIRED_COLUMNS, run_osil
from report_generator import build_osil_pdf_report

try:
    import tenant_db
    DB_AVAILABLE = True
except ImportError:
    DB_AVAILABLE = False

plt.switch_backend('Agg')

st.set_page_config(page_title="OSIL™ by Xentrixus", layout="wide")

def check_password() -> bool:
    def password_entered():
        if st.session_state["password"] == "Xentrixus2026!":
            st.session_state["password_correct"] = True
            del st.session_state["password"]  
        else:
            st.session_state["password_correct"] = False
    if "password_correct" not in st.session_state:
        st.markdown("### Xentrixus OSIL™")
        st.text_input("Enter Executive Master Key", type="password", on_change=password_entered, key="password")
        return False
    elif not st.session_state["password_correct"]:
        st.error("Access Denied.")
        return False
    return True

APP_TITLE = "Xentrixus OSIL™ • Stability Intelligence Engine"
APP_SUB = "Proprietary framework for Business Value Stability Index (BVSI™) and Structural Risk Debt™ generation."

DEMO_INCIDENTS = "data/demo_incidents.csv"
DEMO_CHANGES = "data/demo_changes.csv"
DEMO_PROBLEMS = "data/demo_problems.csv"
DEMO_REQUESTS = "data/demo_requests.csv"

# --- MAPPING SPECS ---
INCIDENT_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor", "required": True, "aliases": ["Service", "Application", "CI", "Business Service"]},
    "Service_Tier": {"label": "Service Tier", "required": False, "aliases": ["Service_Tier", "Tier", "Criticality"]},
    "Opened_Date": {"label": "Opened Date", "required": True, "aliases": ["Opened_Date", "Opened", "Created"]},
    "Resolved_Date": {"label": "Resolved Date", "required": False, "aliases": ["Resolved_Date", "Resolved"]},
    "Closed_Date": {"label": "Closed Date", "required": False, "aliases": ["Closed_Date", "Closed"]},
    "State": {"label": "Incident State", "required": True, "aliases": ["State", "Status"]},
    "Priority": {"label": "Priority", "required": True, "aliases": ["Priority", "Severity", "Urgency"]},
    "Assignment_Group": {"label": "Assignment Group", "required": True, "aliases": ["Assignment Group", "Assignment_Group"]},
    "Channel": {"label": "Channel", "required": False, "aliases": ["Channel", "Contact Type"]},
    "Reassignment_Count": {"label": "Reassignment Count", "required": False, "aliases": ["Reassignment Count", "Reassignments"]},
    "Reopened_Flag": {"label": "Reopened Flag", "required": False, "aliases": ["Reopened_Flag", "Reopened"]},
    "Category": {"label": "Category", "required": False, "aliases": ["Category", "Type"]},
    "Change_Related_Flag": {"label": "Change Related Flag", "required": False, "aliases": ["Change_Related_Flag", "RFC_Flag"]},
    "Problem_ID": {"label": "Problem ID", "required": False, "aliases": ["Problem_ID", "Problem"]},
}

CHANGE_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor", "required": True, "aliases": ["Service", "Application", "CI"]},
    "Change_ID": {"label": "Change ID", "required": False, "aliases": ["Change_ID", "Change", "RFC"]},
    "Change_Type": {"label": "Change Type", "required": True, "aliases": ["Type", "Change Type"]},
    "Actual_Start": {"label": "Actual Start Date", "required": True, "aliases": ["Actual Start", "Actual_Start"]},
    "Actual_End": {"label": "Actual End Date", "required": False, "aliases": ["Actual End", "Actual_End"]},
    "Failed_Flag": {"label": "Failed Flag", "required": False, "aliases": ["Failed_Flag", "Failure_Flag"]},
}

PROBLEM_MAPPING_SPEC = {
    "Service": {"label": "Operational Anchor", "required": True, "aliases": ["Service", "Application", "CI"]},
    "Problem_ID": {"label": "Problem ID", "required": True, "aliases": ["Problem_ID", "Problem"]},
    "RCA_Completed_Flag": {"label": "RCA Completed Flag", "required": False, "aliases": ["RCA_Completed_Flag"]},
    "Known_Error_Flag": {"label": "Known Error Flag", "required": False, "aliases": ["Known_Error_Flag"]},
    "Root_Cause_Text": {"label": "Root Cause Description", "required": False, "aliases": ["Cause", "Root_Cause_Text"]},
}

REQUEST_MAPPING_SPEC = {
    "Item": {"label": "Catalog Item", "required": True, "aliases": ["Item", "Catalog Item"]},
}

# --- UTILS ---
def safe_read_csv(file_or_path):
    for enc in ["utf-8", "utf-8-sig", "cp1252", "latin1"]:
        try:
            if isinstance(file_or_path, str): return pd.read_csv(file_or_path, encoding=enc, engine="python")
            file_or_path.seek(0)
            return pd.read_csv(file_or_path, encoding=enc, engine="python")
        except: continue
    raise ValueError("Unable to read CSV.")

def _safe_read_csv(path: str) -> pd.DataFrame:
    return safe_read_csv(path) if os.path.exists(path) else pd.DataFrame()

def _load_demo_data():
    return _safe_read_csv(DEMO_INCIDENTS), _safe_read_csv(DEMO_CHANGES), _safe_read_csv(DEMO_PROBLEMS), _safe_read_csv(DEMO_REQUESTS)

def _required_template_text():
    return ",".join(INCIDENT_REQUIRED_COLUMNS) + "\nCustomer Portal,2026-01-05 08:00,P2,Closed,Network Services"

def _normalize_col_name(x: str) -> str:
    return str(x).strip().lower().replace("_", " ").replace("-", " ")

def _fuzzy_suggest(columns, aliases):
    norm_to_original = {_normalize_col_name(c): c for c in columns}
    for alias in aliases:
        a = _normalize_col_name(alias)
        if a in norm_to_original: return norm_to_original[a]
    return "None"

def _render_mapping_ui(df, spec, title, key_prefix):
    st.markdown(f"#### {title}")
    columns = list(df.columns)
    options = ["None"] + columns
    mapping = {}
    for canonical, cfg in spec.items():
        suggested = _fuzzy_suggest(columns, cfg["aliases"]) if columns else "None"
        idx = options.index(suggested) if suggested in options else 0
        selected = st.selectbox(f"{cfg['label']}{' *' if cfg['required'] else ''}", options=options, index=idx, key=f"{key_prefix}_{canonical}")
        mapping[canonical] = None if selected == "None" else selected
    return mapping

def _apply_mapping(df, mapping):
    return df.rename(columns={v: k for k, v in mapping.items() if v and v in df.columns})

def _validate_mapping(mapping, spec, dataset_name):
    return [f"{dataset_name}: {canonical}" for canonical, cfg in spec.items() if cfg["required"] and not mapping.get(canonical)]

# --- MAIN APP ---
def main():
    if not check_password(): return
    
    with st.sidebar:
        st.markdown("### Executive Controls")
        if st.button("Log Out", use_container_width=True):
            st.session_state["password_correct"] = False
            st.rerun()

    st.title(APP_TITLE)
    
    # ... [Insert your existing Sidebar / File Uploader / Mapping Logic Here] ...
    # (Ensure run_osil is called with all mappings applied)
    
    # If run_requested:
    results = run_osil(...)
    
    # --- DASHBOARD RENDERING ---
    # Metric Row
    m1, m2, m3 = st.columns(3)
    m1.metric("BVSI™", f"{results['bvsi']:.1f}")
    m2.metric("Posture", results["posture"])
    m3.metric("Data Quality", f"{results['readiness_score']:.1f}%")

    # SVC DIAGRAM
    st.divider()
    st.subheader("ITIL 4 Service Value Chain (SVC)")
    svc = results.get("svc_scores", {})
    if svc:
        def _get_svc_color(s): return "#059669" if s >= 70 else "#D97706" if s >= 55 else "#DC2626"
        st.button(f"Plan ({svc.get('Plan', 0.0):.1f})", use_container_width=True, type="primary")
        c1, c2 = st.columns([1, 2])
        with c1: st.button(f"Engage\n({svc.get('Engage', 0.0):.1f})", use_container_width=True)
        with c2:
            st.button(f"Design & Transition ({svc.get('Design & Transition', 0.0):.1f})", use_container_width=True)
            st.button(f"Obtain / Build ({svc.get('Obtain/Build', 0.0):.1f})", use_container_width=True)
            st.button(f"Deliver & Support ({svc.get('Deliver & Support', 0.0):.1f})", use_container_width=True)
        st.button(f"Improve ({svc.get('Improve', 0.0):.1f})", use_container_width=True)
    
    # REMEDIATION MANDATE
    st.markdown("---")
    st.subheader("Executive Remediation Mandate")
    weakest_svs = min(results.get("svs_scores", {}), key=results.get("svs_scores", {}).get)
    weakest_svc = min(results.get("svc_scores", {}), key=results.get("svc_scores", {}).get)
    
    svs_adv = {"Governance": "Invest in policy enforcement.", "Continual Improvement": "Reallocate capacity to debt paydown.", "Practices": "Standardize assignment routing.", "Guiding Principles": "Realign IT with business outcomes."}
    svc_adv = {"Plan": "Immediate data audit.", "Improve": "Implement formal Problem Management.", "Engage": "Bridge trust gaps.", "Design & Transition": "Enforce Tier 1 validation.", "Obtain/Build": "Formalize quality gates.", "Deliver & Support": "Deploy automated playbooks."}

    c1, c2 = st.columns(2)
    with c1: st.info(f"**SVS Focus ({weakest_svs}):** {svs_adv.get(weakest_svs, 'Review health.')}")
    with c2: st.warning(f"**SVC Focus ({weakest_svc}):** {svc_adv.get(weakest_svc, 'Review throughput.')}")

    # ... [Keep your Charts, Interpretation, and PDF download button here] ...

if __name__ == "__main__":
    main()
