
# ===============================
# OSIL Streamlit Application
# app.py
# ===============================

import streamlit as st
import pandas as pd
import numpy as np
from difflib import get_close_matches
from osil_engine import run_osil
from report_generator import build_osil_pdf_report

st.set_page_config(page_title="OSIL Stability Intelligence", layout="wide")

# -------------------------------
# CSV SAFE LOADER (handles Excel encodings)
# -------------------------------
def safe_read_csv(file):
    encodings = ["utf-8", "utf-8-sig", "cp1252", "latin1", "utf-16"]
    for enc in encodings:
        try:
            file.seek(0)
            return pd.read_csv(file, encoding=enc)
        except:
            continue
    raise Exception("Unable to read CSV file encoding.")

# -------------------------------
# SIMPLE COLUMN NORMALIZATION
# -------------------------------
def norm(x):
    return str(x).strip().lower().replace("_"," ").replace("-"," ")

# -------------------------------
# FUZZY COLUMN MATCHING
# -------------------------------
def suggest_column(columns, aliases):
    cols_norm = {norm(c):c for c in columns}
    for a in aliases:
        if norm(a) in cols_norm:
            return cols_norm[norm(a)]

    for a in aliases:
        match = get_close_matches(norm(a), list(cols_norm.keys()), n=1, cutoff=0.7)
        if match:
            return cols_norm[match[0]]

    return None

# -------------------------------
# INCIDENT COLUMN SPEC
# -------------------------------
INCIDENT_SPEC = {
"Service":["service","application","system","ci","configuration item"],
"Opened_Date":["opened","created","open date"],
"Priority":["priority","severity","impact"],
"Resolved_Date":["resolved","resolution"],
"Reopened_Flag":["reopened"],
"Category":["category","type"],
"Change_Related_Flag":["change related"],
"Problem_ID":["problem","problem id"]
}

# -------------------------------
# CHANGE COLUMN SPEC
# -------------------------------
CHANGE_SPEC = {
"Service":["service","application","system","ci"],
"Change_Start":["start","planned start"],
"Change_End":["end","completed"],
"Failed_Flag":["failed","failure","success flag"],
"Risk":["risk"],
"Category":["type","category"]
}

# -------------------------------
# PROBLEM COLUMN SPEC
# -------------------------------
PROBLEM_SPEC = {
"Service":["service","application","system","ci"],
"Problem_ID":["problem","problem id"],
"Opened_Date":["opened","created"],
"Closed_Date":["closed"],
"Root_Cause":["root cause","rca","cause"],
"RCA_Completed_Flag":["rca complete"],
"Workaround_Available":["workaround"]
}

# -------------------------------
# GENERIC COLUMN MAPPING UI
# -------------------------------
def mapping_ui(df, spec, key_prefix):

    mapping = {}

    for field,aliases in spec.items():
        suggestion = suggest_column(df.columns, aliases)

        options = ["-- None --"] + list(df.columns)

        idx = 0
        if suggestion and suggestion in options:
            idx = options.index(suggestion)

        sel = st.selectbox(
            f"{field}",
            options,
            index=idx,
            key=f"{key_prefix}_{field}"
        )

        mapping[field] = None if sel=="-- None --" else sel

    return mapping

# -------------------------------
# APPLY MAPPING
# -------------------------------
def apply_mapping(df,mapping):

    rename={}

    for k,v in mapping.items():
        if v and v in df.columns:
            rename[v]=k

    return df.rename(columns=rename)

# -------------------------------
# MAIN APP
# -------------------------------
def main():

    st.title("OSIL Stability Intelligence")
    st.caption("Operational Stability Intelligence — Executive ITSM Analytics")

    tenant = st.text_input("Organization Name","Default")

    st.header("Upload Data")

    c1,c2,c3 = st.columns(3)

    with c1:
        inc_file = st.file_uploader("Incident CSV", type=["csv"])

    with c2:
        chg_file = st.file_uploader("Change CSV", type=["csv"])

    with c3:
        prb_file = st.file_uploader("Problem CSV", type=["csv"])

    if inc_file:

        inc_df = safe_read_csv(inc_file)

        st.subheader("Incident Column Mapping")
        inc_map = mapping_ui(inc_df, INCIDENT_SPEC, "inc")

        chg_df=None
        prb_df=None

        if chg_file:
            chg_df = safe_read_csv(chg_file)
            st.subheader("Change Column Mapping")
            chg_map = mapping_ui(chg_df, CHANGE_SPEC, "chg")

        if prb_file:
            prb_df = safe_read_csv(prb_file)
            st.subheader("Problem Column Mapping")
            prb_map = mapping_ui(prb_df, PROBLEM_SPEC, "prb")

        if st.button("Run OSIL Analysis"):

            inc_df = apply_mapping(inc_df,inc_map)

            if chg_file:
                chg_df = apply_mapping(chg_df,chg_map)

            if prb_file:
                prb_df = apply_mapping(prb_df,prb_map)

            results = run_osil(
                incidents_df = inc_df,
                changes_df = chg_df,
                problems_df = prb_df
            )

            st.success("Analysis Completed")

            c1,c2,c3 = st.columns(3)

            c1.metric("BVSI", round(results["bvsi"],1))
            c2.metric("Posture", results["posture"])
            c3.metric("Data Readiness", f"{round(results['readiness_score'],1)}%")

            st.subheader("Executive Interpretation")
            st.write(results["exec_text"])

            st.subheader("Service Risk")
            st.dataframe(results["top10"])

            st.subheader("SIP Candidates")
            st.dataframe(results["sip_view"])

            if st.button("Download Executive PDF"):

                pdf = build_osil_pdf_report(results)

                st.download_button(
                    "Download PDF",
                    pdf,
                    file_name="OSIL_Executive_Report.pdf"
                )

if __name__ == "__main__":
    main()
