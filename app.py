import os
from typing import Dict, List, Optional, Tuple

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
import streamlit as st

from osil_engine import INCIDENT_REQUIRED_COLUMNS, run_osil
from report_generator import build_osil_pdf_report


st.set_page_config(
    page_title="OSIL™ by Xentrixus",
    layout="wide",
)

APP_TITLE = "Xentrixus OSIL™ — Stability Intelligence MVP"
APP_SUB = (
    "Upload Incident / Change / Problem exports → get BVSI™, Structural Risk Debt™, "
    "SIP priorities, and executive interpretation."
)

DEMO_INCIDENTS = "data/demo_incidents.csv"
DEMO_CHANGES = "data/demo_changes.csv"
DEMO_PROBLEMS = "data/demo_problems.csv"


# ============================================================
# Mapping configuration
# ============================================================
INCIDENT_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor (Service / Application / CI / etc.)",
        "required": True,
        "aliases": [
            "Service", "Business Service", "business_service", "Application", "Application Name",
            "CI", "Configuration Item", "CI Name", "Affected Service", "Service Offering",
            "Product", "System", "App", "Application Service"
        ],
    },
    "Service_Tier": {
        "label": "Service Tier (optional)",
        "required": False,
        "aliases": ["Service_Tier", "Tier", "Criticality", "Support Tier"],
    },
    "Opened_Date": {
        "label": "Opened / Created Date",
        "required": True,
        "aliases": ["Opened_Date", "Opened", "Created", "Created_At", "Open Date", "Created Date"],
    },
    "Resolved_Date": {
        "label": "Resolved Date (optional)",
        "required": False,
        "aliases": ["Resolved_Date", "Resolved", "Resolved_At", "Resolution Date"],
    },
    "Closed_Date": {
        "label": "Closed Date (optional)",
        "required": False,
        "aliases": ["Closed_Date", "Closed", "Closed_At", "Close Date"],
    },
    "Priority": {
        "label": "Priority / Severity",
        "required": True,
        "aliases": ["Priority", "Severity", "Impact/Priority", "Incident Priority"],
    },
    "Reopened_Flag": {
        "label": "Reopened Flag (optional)",
        "required": False,
        "aliases": ["Reopened_Flag", "Reopened", "Reopen_Flag", "Reopened?"],
    },
    "Category": {
        "label": "Category / Type (optional)",
        "required": False,
        "aliases": ["Category", "Subcategory", "Type", "Assignment Category"],
    },
    "Change_Related_Flag": {
        "label": "Change-Related Flag (optional)",
        "required": False,
        "aliases": ["Change_Related_Flag", "RFC_Flag", "Change Related", "Was Change Related"],
    },
    "Problem_ID": {
        "label": "Problem ID / Link (optional)",
        "required": False,
        "aliases": ["Problem_ID", "Problem", "Problem Number", "Linked Problem"],
    },
}

CHANGE_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor for Changes",
        "required": True,
        "aliases": [
            "Service", "Business Service", "Application", "Application Name", "CI",
            "Configuration Item", "CI Name", "Service Offering", "Product", "System"
        ],
    },
    "Change_ID": {
        "label": "Change ID / RFC (optional)",
        "required": False,
        "aliases": ["Change_ID", "Change", "RFC", "Change Number"],
    },
    "Change_Start": {
        "label": "Change Start Date",
        "required": True,
        "aliases": [
            "Change_Start", "Change_Start_Date", "Start_Date", "Planned_Start",
            "Implemented_Date", "Implementation Start", "Start"
        ],
    },
    "Change_End": {
        "label": "Change End Date (optional)",
        "required": False,
        "aliases": [
            "Change_End", "Change_End_Date", "End_Date", "Planned_End",
            "Completed_Date", "Implementation End", "End"
        ],
    },
    "Change_Status": {
        "label": "Change Status (optional)",
        "required": False,
        "aliases": ["Change_Status", "Status", "State"],
    },
    "Failed_Flag": {
        "label": "Failure Flag or Success Flag (optional)",
        "required": False,
        "aliases": ["Failed_Flag", "Failure_Flag", "Implementation_Success_Flag", "Success_Flag"],
    },
    "Risk": {
        "label": "Risk (optional)",
        "required": False,
        "aliases": ["Risk", "Risk_Level", "Risk Level"],
    },
    "Category": {
        "label": "Change Category / Type (optional)",
        "required": False,
        "aliases": ["Category", "Type", "Change_Type"],
    },
}

PROBLEM_MAPPING_SPEC = {
    "Service": {
        "label": "Operational Anchor for Problems",
        "required": True,
        "aliases": [
            "Service", "Business Service", "Application", "Application Name", "CI",
            "Configuration Item", "CI Name", "Service Offering", "Product", "System"
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
        "aliases": ["Opened_Date", "Opened", "Created"],
    },
    "Resolved_Date": {
        "label": "Problem Resolved Date (optional)",
        "required": False,
        "aliases": ["Resolved_Date", "Resolved"],
    },
    "Closed_Date": {
        "label": "Problem Closed Date (optional)",
        "required": False,
        "aliases": ["Closed_Date", "Closed"],
    },
    "State": {
        "label": "Problem State / Status (optional)",
        "required": False,
        "aliases": ["State", "Status"],
    },
    "RCA_Completed_Flag": {
        "label": "RCA Completed Flag (optional)",
        "required": False,
        "aliases": ["RCA_Completed_Flag", "RCA_Completed", "RCA Complete"],
    },
    "Known_Error_Flag": {
        "label": "Known Error Flag (optional)",
        "required": False,
        "aliases": ["Known_Error_Flag", "Known_Error", "Known Error"],
    },
    "Priority": {
        "label": "Problem Priority (optional)",
        "required": False,
        "aliases": ["Priority", "Severity"],
    },
    "Category": {
        "label": "Problem Category (optional)",
        "required": False,
        "aliases": ["Category", "Type"],
    },
}


# ============================================================
# Helpers
# ============================================================
def _safe_read_csv(path: str) -> pd.DataFrame:
    if os.path.exists(path):
        return pd.read_csv(path)
    return pd.DataFrame()


def _load_demo_triplet() -> Tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    inc = _safe_read_csv(DEMO_INCIDENTS)
    chg = _safe_read_csv(DEMO_CHANGES)
    prb = _safe_read_csv(DEMO_PROBLEMS)
    return inc, chg, prb


def _required_template_text() -> str:
    cols = ",".join(INCIDENT_REQUIRED_COLUMNS)
    example = "Customer Portal,Tier 1,2026-01-05 08:00,P2"
    return f"{cols}\n{example}"


def _normalize_col_name(x: str) -> str:
    return str(x).strip().lower()


def _suggest_column(columns: List[str], aliases: List[str]) -> str:
    if not columns:
        return "-- None --"

    norm_map = {_normalize_col_name(c): c for c in columns}
    for alias in aliases:
        if _normalize_col_name(alias) in norm_map:
            return norm_map[_normalize_col_name(alias)]

    # fuzzy-ish contains fallback
    for alias in aliases:
        a = _normalize_col_name(alias)
        for c in columns:
            cn = _normalize_col_name(c)
            if a in cn or cn in a:
                return c

    return "-- None --"


def _render_mapping_ui(df: pd.DataFrame, spec: Dict[str, Dict[str, object]], title: str, key_prefix: str) -> Dict[str, Optional[str]]:
    st.markdown(f"#### {title}")
    st.caption("Confirm the best available column for each canonical OSIL field. Use the operational anchor that makes the most sense for the organization.")

    columns = list(df.columns)
    options = ["-- None --"] + columns
    mapping: Dict[str, Optional[str]] = {}

    for canonical, cfg in spec.items():
        suggested = _suggest_column(columns, cfg["aliases"]) if columns else "-- None --"
        if suggested not in options:
            suggested = "-- None --"

        idx = options.index(suggested) if suggested in options else 0
        selected = st.selectbox(
            f"{cfg['label']}{' *' if cfg['required'] else ''}",
            options=options,
            index=idx,
            key=f"{key_prefix}_{canonical}",
        )
        mapping[canonical] = None if selected == "-- None --" else selected

    return mapping


def _apply_mapping(df: pd.DataFrame, mapping: Dict[str, Optional[str]]) -> pd.DataFrame:
    out = df.copy()
    rename_map = {}

    for canonical, selected in mapping.items():
        if selected and selected in out.columns:
            rename_map[selected] = canonical

    out = out.rename(columns=rename_map)
    return out


def _validate_mapping(mapping: Dict[str, Optional[str]], spec: Dict[str, Dict[str, object]], dataset_name: str) -> List[str]:
    missing = []
    for canonical, cfg in spec.items():
        if cfg["required"] and not mapping.get(canonical):
            missing.append(f"{dataset_name}: {canonical}")
    return missing


def heatmap_chart(hm: pd.DataFrame):
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


def render_service_instability_leaders(service_risk_df: pd.DataFrame) -> None:
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
            meaning = "Instability patterns frequently occur near change windows, suggesting governance gaps or insufficient pre-release validation."
            action = "Start a SIP focused on change governance: Tier-1 controls, stronger validation, and post-change monitoring."

        st.markdown(
            f"""
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
""",
            unsafe_allow_html=True,
        )


def _build_pdf_payload(results: dict, tenant_name: str) -> dict:
    return {
        "bvsi": results["bvsi"],
        "posture": results["posture"],
        "as_of": results["as_of"],
        "executive_interpretation": results["exec_text"],
        "domain_scores": results["domain_scores"],
        "service_risk_top10": results["top10"],
        "sip_candidates": results["sip_view"],
        "data_readiness_score": results["readiness_score"],
        "service_anchor_used": results["anchor_used"],
        "detected_dataset": results["practice_type"],
        "tenant_name": tenant_name,
    }


# ============================================================
# Main
# ============================================================
def main():
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
                inc_preview = pd.read_csv(inc_file)
                inc_file.seek(0)
                st.markdown("### Incident Mapping")
                inc_mapping = _render_mapping_ui(inc_preview, INCIDENT_MAPPING_SPEC, "Map Incident Columns", "incmap")
            except Exception as e:
                st.error(f"Could not read Incident CSV: {e}")
                return
        else:
            inc_preview = None
            inc_mapping = {}

        if chg_file is not None:
            try:
                chg_preview = pd.read_csv(chg_file)
                chg_file.seek(0)
                st.markdown("### Change Mapping")
                chg_mapping = _render_mapping_ui(chg_preview, CHANGE_MAPPING_SPEC, "Map Change Columns", "chgmap")
            except Exception as e:
                st.error(f"Could not read Change CSV: {e}")
                return
        else:
            chg_preview = None
            chg_mapping = {}

        if prb_file is not None:
            try:
                prb_preview = pd.read_csv(prb_file)
                prb_file.seek(0)
                st.markdown("### Problem Mapping")
                prb_mapping = _render_mapping_ui(prb_preview, PROBLEM_MAPPING_SPEC, "Map Problem Columns", "prbmap")
            except Exception as e:
                st.error(f"Could not read Problem CSV: {e}")
                return
        else:
            prb_preview = None
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
                incidents_df = pd.read_csv(inc_file)
                incidents_df = _apply_mapping(incidents_df, inc_mapping)

                if chg_file is not None:
                    changes_df = pd.read_csv(chg_file)
                    changes_df = _apply_mapping(changes_df, chg_mapping)
                else:
                    changes_df = pd.DataFrame()

                if prb_file is not None:
                    problems_df = pd.read_csv(prb_file)
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
            "For customer/org uploads, map the best available operational anchor — service, application, CI, or equivalent."
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
        return

    # Top metrics
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

    st.divider()

    st.subheader("Operational Stability Profile")
    rc1, rc2 = st.columns([1.15, 1.0])

    with rc1:
        fig_radar = radar_chart(results["domain_scores"])
        st.pyplot(fig_radar, use_container_width=True)
        st.caption("How to read: balanced shape = aligned governance; collapsed axis = maturity gap requiring targeted SIP focus.")

    with rc2:
        st.markdown("### Stability Domain Scores (0–100)")
        st.dataframe(
            pd.DataFrame({"Domain": list(results["domain_scores"].keys()), "Score": list(results["domain_scores"].values())}),
            use_container_width=True,
        )

    st.divider()

    service_risk_df = results["service_risk_df"]
    top10 = results["top10"].copy()

    if top10 is not None and not top10.empty:
        hm = top10.set_index(top10["Service"] + " (" + top10["Service_Tier"] + ")")[
            ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
        ].rename(columns={
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision",
        })
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

        st.markdown("**Top 10 Services — Risk Breakdown**")
        st.dataframe(top10, use_container_width=True)
    else:
        st.info("No service risk data available.")

    st.divider()

    render_service_instability_leaders(service_risk_df)

    st.divider()

    st.markdown("### Top SIP Candidates (Next 30 Days)")
    st.dataframe(results["sip_view"], use_container_width=True)

    st.divider()

    with st.expander("Preview — Incidents", expanded=False):
        st.dataframe(incidents_df.head(20), use_container_width=True)

    if changes_df is not None and not changes_df.empty:
        with st.expander("Preview — Changes", expanded=False):
            st.dataframe(changes_df.head(20), use_container_width=True)

    if problems_df is not None and not problems_df.empty:
        with st.expander("Preview — Problems", expanded=False):
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
    except Exception as e:
        st.error(f"PDF generation failed: {e}")


if __name__ == "__main__":
    main()
