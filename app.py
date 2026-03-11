"""OSIL Streamlit Application"""
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

plt.switch_backend('Agg')

st.set_page_config(
    page_title="OSIL™ by Xentrixus",
    layout="wide",
)

APP_TITLE = "Xentrixus OSIL™ • Stability Intelligence MVP"
APP_SUB = (
    "Upload Incident / Change / Problem exports to get BVSI™, Structural Risk Debt™, "
    "SIP priorities, and executive interpretation."
)

DEMO_INCIDENTS = "data/demo_incidents.csv"
DEMO_CHANGES = "data/demo_changes.csv"
DEMO_PROBLEMS = "data/demo_problems.csv"

INCIDENT_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor (Service / Application / CI / etc.)",
        "required": True,
        "aliases": [
            "Service", "Business Service", "business_service", "Application", "Application Name",
            "CI", "Configuration Item", "CI Name", "Affected Service", "Service Offering",
            "Product", "System", "App", "Application Service", "Application CI", "Configuration Item Name"
        ],
    },
    "Service_Tier": {
        "label": "Service Tier (optional)",
        "required": False,
        "aliases": ["Service_Tier", "Tier", "Criticality", "Support Tier", "Business Criticality"],
    },
    "Opened_Date": {
        "label": "Opened / Created Date",
        "required": True,
        "aliases": ["Opened_Date", "Opened", "Created", "Created_At", "Open Date", "Created Date", "Opened At"],
    },
    "Resolved_Date": {
        "label": "Resolved Date (optional)",
        "required": False,
        "aliases": ["Resolved_Date", "Resolved", "Resolved_At", "Resolution Date", "Resolved At"],
    },
    "Closed_Date": {
        "label": "Closed Date (optional)",
        "required": False,
        "aliases": ["Closed_Date", "Closed", "Closed_At", "Close Date", "Closed At"],
    },
    "Priority": {
        "label": "Priority / Severity",
        "required": True,
        "aliases": ["Priority", "Severity", "Impact/Priority", "Incident Priority", "Urgency"],
    },
    "Reopened_Flag": {
        "label": "Reopened Flag (optional)",
        "required": False,
        "aliases": ["Reopened_Flag", "Reopened", "Reopen_Flag", "Reopened?", "Was Reopened"],
    },
    "Category": {
        "label": "Category / Type (optional)",
        "required": False,
        "aliases": ["Category", "Subcategory", "Type", "Assignment Category", "Incident Category"],
    },
    "Change_Related_Flag": {
        "label": "Change Related Flag (optional)",
        "required": False,
        "aliases": ["Change_Related_Flag", "RFC_Flag", "Change Related", "Was Change Related", "Related To Change"],
    },
    "Problem_ID": {
        "label": "Problem ID / Link (optional)",
        "required": False,
        "aliases": ["Problem_ID", "Problem", "Problem Number", "Linked Problem", "Parent Problem"],
    },
}

CHANGE_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor for Changes",
        "required": True,
        "aliases": [
            "Service", "Business Service", "Application", "Application Name", "CI",
            "Configuration Item", "CI Name", "Service Offering", "Product", "System",
            "Application CI", "Configuration Item Name"
        ],
    },
    "Change_ID": {
        "label": "Change ID / RFC (optional)",
        "required": False,
        "aliases": ["Change_ID", "Change", "RFC", "Change Number", "RFC Number"],
    },
    "Change_Start": {
        "label": "Change Start Date",
        "required": True,
        "aliases": [
            "Change_Start", "Change_Start_Date", "Start_Date", "Planned_Start",
            "Implemented_Date", "Implementation Start", "Start", "Planned Start"
        ],
    },
    "Change_End": {
        "label": "Change End Date (optional)",
        "required": False,
        "aliases": [
            "Change_End", "Change_End_Date", "End_Date", "Planned_End",
            "Completed_Date", "Implementation End", "End", "Planned End"
        ],
    },
    "Change_Status": {
        "label": "Change Status (optional)",
        "required": False,
        "aliases": ["Change_Status", "Status", "State", "Implementation Status"],
    },
    "Failed_Flag": {
        "label": "Failure Flag or Success Flag (optional)",
        "required": False,
        "aliases": ["Failed_Flag", "Failure_Flag", "Implementation_Success_Flag", "Success_Flag", "Successful Flag"],
    },
    "Risk": {
        "label": "Risk (optional)",
        "required": False,
        "aliases": ["Risk", "Risk_Level", "Risk Level"],
    },
    "Category": {
        "label": "Change Category / Type (optional)",
        "required": False,
        "aliases": ["Category", "Type", "Change_Type", "Change Category"],
    },
}

PROBLEM_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor for Problems",
        "required": True,
        "aliases": [
            "Service", "Business Service", "Application", "Application Name", "CI",
            "Configuration Item", "CI Name", "Service Offering", "Product", "System",
            "Application CI", "Configuration Item Name"
        ],
    },
    "Problem_ID": {
        "label": "Problem ID",
        "required": True,
        "aliases": ["Problem_ID", "Problem", "Problem Number", "PRB Number"],
    },
    "Opened_Date": {
        "label": "Problem Opened Date (optional)",
        "required": False,
        "aliases": ["Opened_Date", "Opened", "Created", "Opened At"],
    },
    "Resolved_Date": {
        "label": "Problem Resolved Date (optional)",
        "required": False,
        "aliases": ["Resolved_Date", "Resolved", "Resolved At"],
    },
    "Closed_Date": {
        "label": "Problem Closed Date (optional)",
        "required": False,
        "aliases": ["Closed_Date", "Closed", "Closed At"],
    },
    "State": {
        "label": "Problem State / Status (optional)",
        "required": False,
        "aliases": ["State", "Status", "Problem State"],
    },
    "RCA_Completed_Flag": {
        "label": "RCA Completed Flag (optional)",
        "required": False,
        "aliases": ["RCA_Completed_Flag", "RCA_Completed", "RCA Complete", "RCA Done"],
    },
    "Known_Error_Flag": {
        "label": "Known Error Flag (optional)",
        "required": False,
        "aliases": ["Known_Error_Flag", "Known_Error", "Known Error"],
    },
    "Priority": {
        "label": "Problem Priority (optional)",
        "required": False,
        "aliases": ["Priority", "Severity", "Urgency"],
    },
    "Category": {
        "label": "Problem Category (optional)",
        "required": False,
        "aliases": ["Category", "Type", "Problem Category"],
    },
    "Root_Cause_Text": {
        "label": "Root Cause Description (optional)",
        "required": False,
        "aliases": ["Root Cause", "Root_Cause", "RCA Details", "RCA Description", "Cause", "Primary Cause"],
    },
    "Contributing_Cause_Text": {
        "label": "Contributing Cause Description (optional)",
        "required": False,
        "aliases": ["Contributing Cause", "Contributing_Cause", "Sub Cause", "Secondary Cause"],
    },
}

def safe_read_csv(file_or_path):
    """Safely read CSV with multiple encoding attempts"""
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1", "utf-16"]
    last_error = None

    for enc in encodings:
        try:
            if isinstance(file_or_path, str):
                return pd.read_csv(file_or_path, encoding=enc, engine="python")
            file_or_path.seek(0)
            return pd.read_csv(file_or_path, encoding=enc, engine="python")
        except Exception as e:
            last_error = e
            continue

    raise ValueError(f"Unable to read CSV file. Unsupported encoding or malformed CSV. Last error: {last_error}")

def _safe_read_csv(path: str) -> pd.DataFrame:
    """Safe read with file existence check"""
    if os.path.exists(path):
        return safe_read_csv(path)
    return pd.DataFrame()

def _load_demo_triplet() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Load demo data files"""
    inc = _safe_read_csv(DEMO_INCIDENTS)
    chg = _safe_read_csv(DEMO_CHANGES)
    prb = _safe_read_csv(DEMO_PROBLEMS)
    return inc, chg, prb

def _required_template_text() -> str:
    """Get required column template"""
    cols = ",".join(INCIDENT_REQUIRED_COLUMNS)
    example = "Customer Portal,Tier 1,2026-01-05 08:00,P2"
    return f"{cols}\n{example}"

def _normalize_col_name(x: str) -> str:
    """Normalize column name for matching"""
    return str(x).strip().lower().replace("_", " ").replace("-", " ")

def _fuzzy_suggest(columns: List[str], aliases: List[str]) -> str:
    """Fuzzy match column names"""
    if not columns:
        return "None"

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

    candidates = []
    for alias in aliases:
        a = _normalize_col_name(alias)
        match = get_close_matches(a, norm_columns, n=1, cutoff=0.72)
        if match:
            candidates.append(match[0])

    if candidates:
        return norm_to_original[candidates[0]]

    return "None"

def _render_mapping_ui(df: pd.DataFrame, spec: Dict[str, Dict[str, object]], title: str, key_prefix: str) -> Dict[str, Optional[str]]:
    """Render column mapping UI"""
    st.markdown(f"#### {title}")
    st.caption("Auto detection is only a starting point. Confirm the best field for each canonical OSIL input before running analysis.")

    columns = list(df.columns)
    options = ["None"] + columns
    mapping: Dict[str, Optional[str]] = {}

    for canonical, cfg in spec.items():
        suggested = _fuzzy_suggest(columns, cfg["aliases"]) if columns else "None"
        if suggested not in options:
            suggested = "None"
        idx = options.index(suggested) if suggested in options else 0

        selected = st.selectbox(
            f"{cfg['label']}{' *' if cfg['required'] else ''}",
            options=options,
            index=idx,
            key=f"{key_prefix}_{canonical}",
        )
        mapping[canonical] = None if selected == "None" else selected

    return mapping

def _apply_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    """Apply column mapping to DataFrame"""
    out = df.copy()
    rename_map = {}

    for canonical, selected in mapping.items():
        if selected and selected in out.columns:
            rename_map[selected] = canonical

    out = out.rename(columns=rename_map)
    return out

def _validate_mapping(mapping: Dict[str, Optional[str]], spec: Dict[str, Dict[str, object]], dataset_name: str) -> List[str]:
    """Validate required mappings are present"""
    missing = []
    for canonical, cfg in spec.items():
        if cfg["required"] and not mapping.get(canonical):
            missing.append(f"{dataset_name}: {canonical}")
    return missing

def heatmap_chart(hm: pd.DataFrame):
    """Generate heatmap chart for Streamlit display"""
    fig = plt.figure(figsize=(10, 5), dpi=160)
    ax = plt.gca()
    im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(hm.columns)))
    ax.set_xticklabels(list(hm.columns), fontsize=9)
    ax.set_yticks(range(len(hm.index)))
    ax.set_yticklabels(list(hm.index), fontsize=9)
    ax.set_title("Service Stability Heatmap", fontsize=12)

    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            ax.text(j, i, f"{int(round(float(hm.iat[i, j]), 0))}", ha="center", va="center", fontsize=8)

    cbar = plt.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cbar.set_label("Risk Score", fontsize=8)
    cbar.ax.tick_params(labelsize=7)
    plt.tight_layout()
    return fig

def radar_chart(domain_scores: dict):
    """Generate radar chart for Streamlit display"""
    labels = list(domain_scores.keys())
    values = [float(domain_scores[k]) for k in labels]
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + [values[0]]
    angles_loop = angles + [angles[0]]

    fig = plt.figure(figsize=(6.2, 5.0), dpi=160)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.12)
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)
    ax.set_title("Operational Stability Radar", fontsize=12, pad=16)
    plt.tight_layout()
    return fig

def plot_pareto(df: pd.DataFrame):
    """Generate locked numeric axis Pareto chart for Root Cause Themes"""
    fig, ax1 = plt.subplots(figsize=(8.5, 5.0), dpi=120)
    
    labels = [textwrap.fill(str(x)[:22] + "..." if len(str(x)) > 22 else str(x), width=18) for x in df["Theme"]]
    x_pos = np.arange(len(df))
    
    ax1.bar(x_pos, df["Frequency"], color="#3B82F6", width=0.55)
    ax1.set_ylabel("Frequency of Root Cause", color="#0F172A", fontweight="bold", fontsize=9)
    ax1.tick_params(axis="y", labelcolor="#0F172A")
    
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(labels, rotation=35, ha="right", fontsize=9)

    ax2 = ax1.twinx()
    ax2.plot(x_pos, df["Cumulative_Pct"], color="#DC2626", marker="o", linewidth=2.5)
    ax2.set_ylabel("Cumulative Impact Percentage", color="#DC2626", fontweight="bold", fontsize=9)
    ax2.set_ylim(0, 110)
    
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.title("Eighty Twenty Rule: Top Structural Risk Themes", fontweight="bold", color="#0F172A", pad=15)
    
    plt.gcf().subplots_adjust(bottom=0.35)
    plt.tight_layout()
    return fig

def plot_impact_matrix(service_risk_df: pd.DataFrame):
    """Generate Dual Axis Chart for Disruption vs Recurrence using Service Risk Data directly"""
    if service_risk_df.empty or "Active_Disruption_P1_P2" not in service_risk_df.columns:
        fig, ax = plt.subplots(figsize=(7, 4.5))
        ax.text(0.5, 0.5, 'Insufficient Data for Impact Matrix', ha='center', va='center')
        return fig

    merged = service_risk_df.sort_values(by=["Active_Disruption_P1_P2", "Total_Service_Risk"], ascending=[False, False]).head(5).copy()
    
    merged["Active_Disruption_P1_P2"] = pd.to_numeric(merged["Active_Disruption_P1_P2"], errors="coerce").fillna(0)
    merged["Recurrence_Risk"] = pd.to_numeric(merged["Recurrence_Risk"], errors="coerce").fillna(0)
        
    fig, ax1 = plt.subplots(figsize=(8.0, 5.0), dpi=120)
    
    x_pos = np.arange(len(merged))
    labels = [textwrap.fill(str(x)[:20], width=12) for x in merged["Service"]]
    
    bars = ax1.bar(x_pos, merged["Active_Disruption_P1_P2"], color="#DC2626", width=0.45, alpha=0.9, label="Active Disruption (P1/P2 Count)")
    ax1.set_ylabel("Active Disruption Volume (P1 and P2)", color="#DC2626", fontweight="bold", fontsize=9)
    ax1.tick_params(axis="y", labelcolor="#DC2626")
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(labels, rotation=0, ha="center", fontsize=9, fontweight="bold")
    
    for bar in bars:
        yval = bar.get_height()
        if yval > 0:
            ax1.text(bar.get_x() + bar.get_width()/2.0, yval + 0.05, f"{int(yval)}", ha='center', va='bottom', color='#DC2626', fontweight='bold', fontsize=10)

    max_disruption = float(merged["Active_Disruption_P1_P2"].max())
    if max_disruption < 5:
        ax1.set_ylim(0, 5)
    else:
        ax1.set_ylim(0, max_disruption * 1.3)

    ax2 = ax1.twinx()
    line = ax2.plot(x_pos, merged["Recurrence_Risk"], color="#0F172A", marker="o", linewidth=2.5, markersize=8, label="Recurrence Risk Score")
    ax2.set_ylabel("Recurrence Risk Score (Zero to 100)", color="#0F172A", fontweight="bold", fontsize=9)
    ax2.tick_params(axis="y", labelcolor="#0F172A")
    ax2.set_ylim(0, 115)
    
    for i, val in enumerate(merged["Recurrence_Risk"]):
        ax2.text(x_pos[i], val + 3, f"{int(val)}", ha='center', va='bottom', color='#0F172A', fontweight='bold', fontsize=9)

    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.title("Executive Strike Zone: Top 5 Services (Disruption vs Recurrence)", fontweight="bold", color="#0F172A", pad=15)
    
    lines_1, labels_1 = ax1.get_legend_handles_labels()
    lines_2, labels_2 = ax2.get_legend_handles_labels()
    ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper center', bbox_to_anchor=(0.5, -0.15), frameon=False, ncol=2, fontsize=9)
    
    plt.gcf().subplots_adjust(bottom=0.25)
    plt.tight_layout()
    return fig

def render_service_instability_leaders(service_risk_df: pd.DataFrame) -> None:
    """Render service instability leaders section"""
    st.subheader("Service Instability Leaders (Top 5)")
    st.caption("Narrative view of the services currently driving the highest operational instability risk.")

    if service_risk_df is None or service_risk_df.empty:
        st.info("No service instability signals available.")
        return

    df = service_risk_df.sort_values("Total_Service_Risk", ascending=False).head(5).copy()

    for rank, (_, row) in enumerate(df.iterrows(), start=1):
        service = str(row.get("Service", "Unknown Service"))
        tier = str(row.get("Service_Tier", "Unknown Tier"))
        score = float(row.get("Total_Service_Risk", 0.0))

        risks = {
            "Recurrence": float(row.get("Recurrence_Risk", 0.0)),
            "MTTR Drag": float(row.get("MTTR_Drag_Risk", 0.0)),
            "Reopen Churn": float(row.get("Reopen_Churn_Risk", 0.0)),
            "Change Collision": float(row.get("Change_Collision_Risk", 0.0)),
        }

        primary = max(risks, key=risks.get)
        primary_score = risks.get(primary, 0.0)

        if primary == "Recurrence":
            meaning = "Recurring incidents suggest unresolved structural issues and repeat operational friction."
            action = "Start a SIP focused on recurrence elimination: clear problem statements, root cause pathways, and preventive controls."
        elif primary == "MTTR Drag":
            meaning = "Recovery times are longer than expected, indicating response coordination gaps, unclear ownership, or weak runbooks."
            action = "Start a SIP focused on recovery execution: playbooks, escalation pathways, and targeted automation to reduce recovery time."
        elif primary == "Reopen Churn":
            meaning = "High reopen rates suggest incomplete resolution or fixes that do not hold under operational load."
            action = "Start a SIP focused on fix quality: improve validation, tighten closure criteria, and drive problem investigations for repeat patterns."
        else:
            meaning = "Instability patterns frequently occur near change windows, suggesting governance gaps or insufficient pre release validation."
            action = "Start a SIP focused on change governance: Tier 1 controls, stronger validation, and post change monitoring."

        st.markdown(
            f'''
            <div style="border:1px solid #D1D5DB;background:#F5F7FA;padding:14px 16px;border-radius:10px;margin-bottom:10px;">
                <div style="font-size:16px;"><b>#{rank} {service}</b></div>
                <div style="margin-top:6px;">
                    <b>Tier:</b> {tier} &nbsp; | &nbsp; <b>Total Risk Score:</b> {score:.1f}
                </div>
                <div style="margin-top:10px;">
                    <b>Primary Instability Driver:</b> {primary} ({primary_score:.0f}/100)<br/>
                    {meaning}
                </div>
                <div style="margin-top:10px;">
                    <b>Recommended Action:</b><br/>
                    {action}
                </div>
            </div>
            ''',
            unsafe_allow_html=True,
        )

def _build_pdf_payload(results: dict, tenant_name: str) -> dict:
    """Build payload for PDF generator with all required fields"""
    return {
        "bvsi": results.get("bvsi", 0.0),
        "posture": results.get("posture", "Unknown"),
        "as_of": results.get("as_of", ""),
        "executive_interpretation": results.get("exec_text", ""),
        "trust_gap_narrative": results.get("trust_gap_narrative", ""),
        "trust_gap_df": results.get("trust_gap_df", pd.DataFrame()),
        "rca_themes_df": results.get("rca_themes_df", pd.DataFrame()),
        "rca_pareto_df": results.get("rca_pareto_df", pd.DataFrame()),
        "domain_scores": results.get("domain_scores", {}),
        "service_risk_top10": results.get("top10", pd.DataFrame()),
        "sip_candidates": results.get("sip_view", pd.DataFrame()),
        "data_readiness_score": results.get("readiness_score", 0.0),
        "service_anchor_used": results.get("anchor_used", "None"),
        "detected_dataset": results.get("practice_type", "unknown"),
        "tenant_name": tenant_name,
    }

def main():
    """Main application function"""
    st.title(APP_TITLE)
    st.caption(APP_SUB)

    tenant_name = st.text_input("Organization / Tenant Name", value="Default")

    st.subheader("Run Options")
    mode = st.radio(
        "Choose a run mode",
        ["Run with Demo Data", "Upload Incident / Change / Problem CSVs"],
        horizontal=True,
    )

    incidents_df: Optional[pd.DataFrame] = None
    changes_df: Optional[pd.DataFrame] = None
    problems_df: Optional[pd.DataFrame] = None
    source_label = None
    run_requested = False

    if mode == "Run with Demo Data":
        if st.button("Run Demo Analysis", use_container_width=True):
            incidents_df, changes_df, problems_df = _load_demo_triplet()
            if incidents_df.empty:
                st.error("Demo load failed: data/demo_incidents.csv was not found or is empty.")
                return
            source_label = "Demo Data"
            run_requested = True
    else:
        st.markdown("### Upload Files")
        c1, c2, c3 = st.columns(3)

        with c1:
            inc_file = st.file_uploader("Incident CSV", type=["csv"], key="inc")
        with c2:
            chg_file = st.file_uploader("Change CSV (optional)", type=["csv"], key="chg")
        with c3:
            prb_file = st.file_uploader("Problem CSV (optional)", type=["csv"], key="prb")

        if inc_file is not None:
            try:
                inc_preview = safe_read_csv(inc_file)
                inc_file.seek(0)
                st.markdown("### Incident Mapping")
                inc_mapping = _render_mapping_ui(inc_preview, INCIDENT_MAPPING_SPEC, "Map Incident Columns", "incmap")
            except Exception as e:
                st.error(f"Could not read Incident CSV: {e}")
                return
        else:
            inc_mapping = {}

        if chg_file is not None:
            try:
                chg_preview = safe_read_csv(chg_file)
                chg_file.seek(0)
                st.markdown("### Change Mapping")
                chg_mapping = _render_mapping_ui(chg_preview, CHANGE_MAPPING_SPEC, "Map Change Columns", "chgmap")
            except Exception as e:
                st.error(f"Could not read Change CSV: {e}")
                return
        else:
            chg_mapping = {}

        if prb_file is not None:
            try:
                prb_preview = safe_read_csv(prb_file)
                prb_file.seek(0)
                st.markdown("### Problem Mapping")
                prb_mapping = _render_mapping_ui(prb_preview, PROBLEM_MAPPING_SPEC, "Map Problem Columns", "prbmap")
            except Exception as e:
                st.error(f"Could not read Problem CSV: {e}")
                return
        else:
            prb_mapping = {}

        st.caption(
            "Use the best available operational anchor for each dataset. "
            "That may be a service, application, CI, service offering, product, or another sensible grouping field."
        )

        if st.button("Run Uploaded Analysis", use_container_width=True):
            if inc_file is None:
                st.error("Please upload an Incident CSV.")
                return

            missing = _validate_mapping(inc_mapping, INCIDENT_MAPPING_SPEC, "Incident")
            if chg_file is not None:
                missing += _validate_mapping(chg_mapping, CHANGE_MAPPING_SPEC, "Change")
            if prb_file is not None:
                missing += _validate_mapping(prb_mapping, PROBLEM_MAPPING_SPEC, "Problem")

            if missing:
                st.error("Missing required mappings: " + " | ".join(missing))
                return

            try:
                incidents_df = safe_read_csv(inc_file)
                incidents_df = _apply_mapping(incidents_df, inc_mapping)

                if chg_file is not None:
                    changes_df = safe_read_csv(chg_file)
                    changes_df = _apply_mapping(changes_df, chg_mapping)
                else:
                    changes_df = pd.DataFrame()

                if prb_file is not None:
                    problems_df = safe_read_csv(prb_file)
                    problems_df = _apply_mapping(problems_df, prb_mapping)
                else:
                    problems_df = pd.DataFrame()

                source_label = f"Upload ({inc_file.name})"
                run_requested = True

            except Exception as e:
                st.error(f"Upload load failed: {e}")
                return

    if not run_requested:
        st.subheader("Required Incident Template")
        st.code(_required_template_text(), language="csv")
        st.info(
            "Incidents are required. Changes and Problems are optional. "
            "For customer org uploads, map the best available operational anchor service, application, CI, or equivalent."
        )
        return

    st.success(f"Loaded: {source_label}")

    try:
        results = run_osil(
            incidents_df=incidents_df,
            changes_df=changes_df,
            problems_df=problems_df,
        )
        results["source_label"] = source_label
        results["tenant_name"] = tenant_name
    except Exception as e:
        st.error(f"Run failed: {e}")
        st.exception(e)
        return

    mc1, mc2, mc3 = st.columns(3)
    mc1.metric("BVSI™", f"{results['bvsi']:.1f}")
    mc2.metric("Operating Posture", results["posture"])
    mc3.metric("Data Readiness", f"{results['readiness_score']:.1f}%")

    st.caption(
        f"Dataset: {results['practice_type']} • "
        f"Service anchor used: {results['anchor_used']} • "
        f"As of: {results['as_of']}"
    )

    st.divider()
    st.subheader("Executive Interpretation")
    st.write(results["exec_text"])
    
    st.markdown("#### Xentrixus OSIL™ Trust Gap Analysis")
    st.info(results["trust_gap_narrative"])
    
    if not results["trust_gap_df"].empty and not results["service_risk_df"].empty:
        fig_impact = plot_impact_matrix(results["service_risk_df"])
        if fig_impact:
            st.pyplot(fig_impact, use_container_width=True)
            
            st.markdown("""
            <div style="background-color: #FEF2F2; border-left: 4px solid #DC2626; padding: 12px 16px; margin-bottom: 24px;">
                <strong>Executive Insight:</strong> Services showing high disruption volumes (red bars) combined with elevated recurrence risk (black lines) represent immediate executive danger zones. These services require immediate executive sponsorship and capital allocation to stop the bleeding.
            </div>
            """, unsafe_allow_html=True)
        
        st.dataframe(results["trust_gap_df"], use_container_width=True)
    else:
        st.warning("Insufficient priority data to generate Trust Gap Matrix. Ensure Priority mapping distinguishes between critical and low urgency events.")

    st.divider()
    st.subheader("Operational Stability Profile")
    rc1, rc2 = st.columns([1.15, 1.0])

    with rc1:
        fig_radar = radar_chart(results["domain_scores"])
        st.pyplot(fig_radar, use_container_width=True)
        st.caption("How to read: balanced shape indicates aligned governance; collapsed axis indicates maturity gap requiring targeted SIP focus.")

    with rc2:
        st.markdown("### Stability Domain Scores (Zero to 100)")
        st.dataframe(
            pd.DataFrame({"Domain": list(results["domain_scores"].keys()), "Score": list(results["domain_scores"].values())}),
            use_container_width=True,
        )

    st.divider()
    st.subheader("Structural Risk Debt™: Thematic Extraction")
    
    if not results["rca_pareto_df"].empty:
        fig_pareto = plot_pareto(results["rca_pareto_df"])
        st.pyplot(fig_pareto, use_container_width=True)
        
        st.markdown("""
        <div style="background-color: #F8FAFC; border-left: 4px solid #0F172A; padding: 12px 16px; margin-bottom: 24px;">
            <strong>Executive Insight:</strong> The Pareto principle dictates that roughly eighty percent of operational instability stems from twenty percent of the underlying causes. The themes clustered on the left side of this axis represent your highest leverage remediation targets. Funding service improvement programs focused exclusively on these top drivers will mathematically eliminate the vast majority of your structural risk debt while optimizing resource allocation.
        </div>
        """, unsafe_allow_html=True)
        
        st.dataframe(results["rca_themes_df"], use_container_width=True)
    else:
        st.error("No thematic root cause data detected. The engine requires text based root causes to extract operational themes. Relying purely on boolean checkboxes hides the true drivers of structural risk debt.")

    st.divider()

    service_risk_df = results["service_risk_df"]
    top10 = results["top10"].copy()

    if top10 is not None and not top10.empty:
        heatmap_df = top10.copy()
        if "Service" in heatmap_df.columns:
            if "Service_Tier" not in heatmap_df.columns:
                heatmap_df["Service_Tier"] = "Unspecified"

            required_risk_cols = [
                "Recurrence_Risk",
                "MTTR_Drag_Risk",
                "Reopen_Churn_Risk",
                "Change_Collision_Risk",
            ]

            for col in required_risk_cols:
                if col not in heatmap_df.columns:
                    heatmap_df[col] = 0.0

            try:
                hm = heatmap_df.set_index(
                    heatmap_df["Service"].astype(str) + " (" + heatmap_df["Service_Tier"].astype(str) + ")"
                )[required_risk_cols].rename(
                    columns={
                        "Recurrence_Risk": "Recurrence",
                        "MTTR_Drag_Risk": "MTTR Drag",
                        "Reopen_Churn_Risk": "Reopen Churn",
                        "Change_Collision_Risk": "Change Collision",
                    }
                )
                hm = hm.apply(pd.to_numeric, errors="coerce").replace([np.inf, -np.inf], np.nan).fillna(0.0)

                st.markdown("### Service Stability Heatmap (Top 10 Services by Risk)")
                hm_fig = heatmap_chart(hm)
                hc1, hc2, hc3 = st.columns([1, 2, 1])
                with hc2:
                    st.pyplot(hm_fig, use_container_width=False)
                st.caption(
                    "How to read: services with consistently high values across multiple columns usually represent the strongest candidates "
                    "for leadership attention or SIP execution."
                )
            except Exception as e:
                st.warning(f"Could not generate heatmap: {e}")

        st.markdown("**Top 10 Services | Risk Breakdown**")
        st.dataframe(top10, use_container_width=True)
    else:
        st.info("No service risk data available.")

    st.divider()
    render_service_instability_leaders(service_risk_df)

    st.divider()
    st.markdown("### Top SIP Candidates (Next Thirty Days)")
    st.dataframe(results["sip_view"], use_container_width=True)

    st.divider()

    with st.expander("Preview | Incidents", expanded=False):
        st.dataframe(incidents_df.head(20), use_container_width=True)

    if changes_df is not None and not changes_df.empty:
        with st.expander("Preview | Changes", expanded=False):
            st.dataframe(changes_df.head(20), use_container_width=True)

    if problems_df is not None and not problems_df.empty:
        with st.expander("Preview | Problems", expanded=False):
            st.dataframe(problems_df.head(20), use_container_width=True)

    st.divider()

    st.markdown("### Executive PDF Report")
    try:
        payload = _build_pdf_payload(results, tenant_name)
        pdf_bytes = build_osil_pdf_report(payload)

        st.download_button(
            "Download Executive PDF",
            data=pdf_bytes,
            file_name=f"OSIL_Executive_Report_{results['as_of']}_{tenant_name}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )
        st.success("PDF generated successfully!")
    except Exception as e:
        st.error(f"PDF generation failed: {e}")
        st.exception(e)

if __name__ == "__main__":
    main()
