import pandas as pd


def _lower_cols(df: pd.DataFrame) -> list[str]:
    return [str(c).strip().lower() for c in df.columns]


def detect_practice_type(df: pd.DataFrame) -> str:
    """
    Option A: simple column-pattern detection.
    """
    cols = _lower_cols(df)
    joined = " | ".join(cols)

    # Incident
    incident_signals = [
        "incident", "inc number", "number", "opened", "opened at", "priority",
        "reopened", "incident state"
    ]
    # Change
    change_signals = [
        "change", "chg", "planned start", "planned end", "actual start", "actual end",
        "change type", "change state", "risk"
    ]
    # Problem
    problem_signals = [
        "problem", "root cause", "contributing cause", "known error", "problem state"
    ]
    # Request
    request_signals = [
        "request", "req", "requested for", "fulfillment", "catalog item"
    ]
    # Telemetry / monitoring
    telemetry_signals = [
        "alert", "event", "latency", "cpu", "memory", "availability", "threshold"
    ]

    def has_any(signals: list[str]) -> bool:
        return any(s in joined for s in signals)

    if has_any(change_signals):
        return "change"
    if has_any(problem_signals):
        return "problem"
    if has_any(request_signals):
        return "request"
    if has_any(telemetry_signals):
        return "telemetry"
    if has_any(incident_signals):
        return "incident"

    return "unknown"


def normalize_service_anchor(df: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    """
    Create a canonical Service_Anchor column using the approved fallback order:
    1) Business Service / Service Offering
    2) Application / Product
    3) CI
    4) Assignment Group
    5) Category
    """
    out = df.copy()

    preferred_order = [
        # Business service / service offering
        "Business_Service",
        "Business Service",
        "Service_Offering",
        "Service Offering",
        "Service",
        "Service_Name",
        "Service Name",

        # Application / product
        "Application",
        "Application_Name",
        "Application Name",
        "Product",
        "Product_Name",
        "Product Name",

        # CI
        "CI",
        "CI_Name",
        "CI Name",
        "Configuration_Item",
        "Configuration Item",
        "Configuration Item (CI)",

        # Assignment
        "Assignment_Group",
        "Assignment Group",

        # Category fallback
        "Category",
    ]

    cols_actual = list(out.columns)
    cols_norm = {str(c).strip().lower(): c for c in cols_actual}

    for candidate in preferred_order:
        norm = candidate.strip().lower()
        if norm in cols_norm:
            actual_col = cols_norm[norm]
            out["Service_Anchor"] = out[actual_col].astype(str).replace({"nan": "Unknown", "": "Unknown"}).fillna("Unknown")
            return out, actual_col

    out["Service_Anchor"] = "Unknown"
    return out, "None"


def calculate_data_readiness(df: pd.DataFrame) -> float:
    """
    Simple readiness score based on presence of basic operational fields.
    """
    cols_norm = {str(c).strip().lower() for c in df.columns}

    checks = {
        "opened": any(c in cols_norm for c in {
            "opened_date", "opened_at", "opened", "created", "created_at", "open_date"
        }),
        "closed": any(c in cols_norm for c in {
            "closed_date", "closed_at", "closed", "resolved", "resolved_at", "updated"
        }),
        "priority": "priority" in cols_norm,
        "service_anchor": "service_anchor" in cols_norm or any(c in cols_norm for c in {
            "business_service", "business service", "service", "application", "application_name",
            "configuration item", "configuration_item", "ci", "ci_name"
        }),
        "reopened": any(c in cols_norm for c in {
            "reopened_flag", "reopened", "reopen_count"
        }),
        "tier": any(c in cols_norm for c in {
            "service_tier", "tier", "criticality"
        }),
        "change_flag": any(c in cols_norm for c in {
            "change_related_flag", "change related flag", "rfc", "request for change (rfc) that caused problem"
        }),
    }

    score = round((sum(checks.values()) / len(checks)) * 100, 1)
    return score
