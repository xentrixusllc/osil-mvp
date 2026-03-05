import pandas as pd

# ---------- PRACTICE DETECTION ----------

def detect_practice_type(df):

    cols = [c.lower() for c in df.columns]

    if "incident" in " ".join(cols) or "opened" in " ".join(cols):
        return "incident"

    if "change" in " ".join(cols) or "planned_start" in " ".join(cols):
        return "change"

    if "problem" in " ".join(cols) or "root_cause" in " ".join(cols):
        return "problem"

    if "request" in " ".join(cols):
        return "request"

    return "unknown"


# ---------- SERVICE ANCHOR NORMALIZATION ----------

def normalize_service_anchor(df):

    anchor_columns = [
        "Business_Service",
        "Service",
        "Service_Name",
        "Application",
        "Application_Name",
        "CI",
        "CI_Name",
        "Configuration_Item",
        "Assignment_Group",
        "Category"
    ]

    for col in anchor_columns:
        if col in df.columns:
            df["Service_Anchor"] = df[col]
            return df, col

    df["Service_Anchor"] = "Unknown"
    return df, "None"


# ---------- DATA READINESS SCORE ----------

def calculate_data_readiness(df):

    score = 0
    checks = 5

    if "Opened_Date" in df.columns or "opened_at" in df.columns:
        score += 1

    if "Closed_Date" in df.columns or "resolved_at" in df.columns:
        score += 1

    if "Priority" in df.columns:
        score += 1

    if "Service_Anchor" in df.columns:
        score += 1

    if "Reopened_Flag" in df.columns:
        score += 1

    readiness = round((score / checks) * 100, 1)

    return readiness
