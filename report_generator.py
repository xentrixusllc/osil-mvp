import io
import re
from typing import Any, Dict, List, Optional

import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

from reportlab.lib import colors
from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import (
    Image,
    PageBreak,
    Paragraph,
    SimpleDocTemplate,
    Spacer,
    Table,
    TableStyle,
)


PAGE_WIDTH, PAGE_HEIGHT = LETTER
CONTENT_WIDTH = 7.4 * inch

NAVY = colors.HexColor("#0A192F")
TEAL = colors.HexColor("#64FFDA")
DARK = colors.HexColor("#111827")
MID = colors.HexColor("#4B5563")
LIGHT = colors.HexColor("#F5F7FA")
LIGHT_BLUE = colors.HexColor("#EAF4FF")
BORDER = colors.HexColor("#D1D5DB")
WHITE = colors.white


def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None:
            return default
        return float(val)
    except Exception:
        return default


def _safe_str(val: Any, default: str = "") -> str:
    if val is None:
        return default
    return str(val)


def _safe_df(val: Any) -> pd.DataFrame:
    if isinstance(val, pd.DataFrame):
        return val.copy()
    if isinstance(val, list):
        try:
            return pd.DataFrame(val)
        except Exception:
            return pd.DataFrame()
    return pd.DataFrame()


def _first_present(payload: Dict[str, Any], keys: List[str], default: Any = None) -> Any:
    for key in keys:
        if key in payload and payload.get(key) is not None:
            return payload.get(key)
    return default


def _clean_text(text: Any) -> str:
    if text is None:
        return ""

    s = str(text)
    s = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", s)
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    s = s.replace("\n\n", "<br/><br/>")
    s = s.replace("\n", "<br/>")
    s = s.replace("<br>", "<br/>").replace("<br />", "<br/>")

    allowed = {"b", "i", "u", "br"}

    def repl_tag(match):
        full = match.group(0)
        tag = match.group(1).lower().replace("/", "")
        return full if tag in allowed else ""

    s = re.sub(r"</?([a-zA-Z0-9]+).*?>", repl_tag, s)
    return s.strip()


def _styles():
    styles = getSampleStyleSheet()

    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=NAVY,
            spaceAfter=4,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=12,
            textColor=MID,
            spaceAfter=8,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13,
            leading=16,
            textColor=NAVY,
            spaceBefore=2,
            spaceAfter=7,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.2,
            leading=12.5,
            textColor=DARK,
            wordWrap="LTR",
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=MID,
            wordWrap="LTR",
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Table",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=7.8,
            leading=9.3,
            textColor=DARK,
            wordWrap="CJK",
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=DARK,
            wordWrap="LTR",
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Note",
            parent=styles["BodyText"],
            fontName="Helvetica-Oblique",
            fontSize=8,
            leading=10,
            textColor=MID,
            wordWrap="LTR",
        )
    )

    return styles


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(MID)
    canvas.drawString(0.55 * inch, 0.35 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
    canvas.drawRightString(7.95 * inch, 0.35 * inch, f"Page {doc.page}")
    canvas.restoreState()


def _header_band(text: str, width: float = CONTENT_WIDTH) -> Table:
    t = Table([[text]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), NAVY),
                ("TEXTCOLOR", (0, 0), (-1, -1), WHITE),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 11.5),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return t


def _accent_rule(width: float = CONTENT_WIDTH, color_hex: str = "#64FFDA", height: int = 3) -> Table:
    t = Table([[""]], colWidths=[width], rowHeights=[height])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(color_hex)),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return t


def _kpi_box(label: str, value: str, styles, width: float = 2.45 * inch, is_numeric: bool = False) -> Table:
    label_para = Paragraph(_clean_text(label), styles["OSIL_TableHeader"])

    value_style = ParagraphStyle(
        name=f"KPIValue_{'Num' if is_numeric else 'Text'}",
        parent=styles["OSIL_Body"],
        fontName="Helvetica-Bold",
        fontSize=17 if is_numeric else 11.5,
        leading=19 if is_numeric else 13,
        textColor=DARK,
        wordWrap="LTR",
    )

    value_para = Paragraph(_clean_text(value), value_style)

    t = Table([[label_para], [value_para]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                ("TOPPADDING", (0, 1), (-1, 1), 2),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
            ]
        )
    )
    return t


def _info_box(title: str, body: str, styles, width: float = CONTENT_WIDTH) -> Table:
    title_para = Paragraph(_clean_text(title), styles["OSIL_TableHeader"])
    body_para = Paragraph(_clean_text(body), styles["OSIL_Body"])
    t = Table([[title_para], [body_para]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LIGHT_BLUE),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F8FBFF")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B8D6F2")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 1), (-1, 1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
            ]
        )
    )
    return t


def _full_width_paragraph(text: str, styles, style_name: str = "OSIL_Body") -> Table:
    para = Paragraph(_clean_text(text), styles[style_name])
    tbl = Table([[para]], colWidths=[CONTENT_WIDTH])
    tbl.setStyle(
        TableStyle(
            [
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ]
        )
    )
    return tbl


def _paragraph_table(
    df: pd.DataFrame,
    col_widths: Optional[List[float]],
    styles,
    header_bg: str = "#E9EDF3",
) -> Table:
    df = _safe_df(df)

    if df.empty:
        data = [[Paragraph("No data available", styles["OSIL_Table"])]]
        tbl = Table(data, colWidths=[CONTENT_WIDTH])
        tbl.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        return tbl

    if col_widths is None or len(col_widths) != len(df.columns):
        col_widths = [CONTENT_WIDTH / len(df.columns)] * len(df.columns)

    header = [Paragraph(_clean_text(c), styles["OSIL_TableHeader"]) for c in df.columns]
    rows = []
    for row in df.values.tolist():
        rows.append([Paragraph(_clean_text(x), styles["OSIL_Table"]) for x in row])

    data = [header] + rows
    tbl = Table(data, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [WHITE, colors.HexColor("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("WORDWRAP", (0, 0), (-1, -1), True),
                ("SPLITBYROW", (0, 0), (-1, -1), True),
            ]
        )
    )
    return tbl


def _posture_signal_text(bvsi: float, posture: str) -> str:
    if bvsi >= 85:
        return (
            f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. "
            "Technology operations are broadly supporting business confidence, though targeted prevention can strengthen resilience further."
        )
    if bvsi >= 70:
        return (
            f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. "
            "Control exists across most operational dimensions, though localized weaknesses remain that could limit reliability at scale."
        )
    if bvsi >= 55:
        return (
            f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. "
            "Governance mechanisms are functioning, but recurring instability patterns remain across higher-impact services."
        )
    if bvsi >= 40:
        return (
            f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. "
            "Operational exposure is visible and likely affecting service confidence, efficiency, or customer experience."
        )
    return (
        f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. "
        "Immediate executive attention is warranted to contain structural instability and protect business performance."
    )


def _bvsi_scale_table(styles) -> Table:
    df = pd.DataFrame(
        [
            ["85–100", "High Confidence Operations", "Technology stability supports business growth and executive confidence."],
            ["70–84", "Controlled and Improving", "Core controls exist; targeted improvement can increase resilience and scale-readiness."],
            ["55–69", "Controlled but Exposed", "Operational control exists, but recurring instability still creates material exposure."],
            ["40–54", "Reactive and Exposed", "Instability is visible and may be impacting reliability, cost, or customer experience."],
            ["<40", "Fragile Operations", "Operational fragility is high and stabilization should be treated as an executive priority."],
        ],
        columns=["BVSI™ Range", "Operating Condition", "Executive Meaning"],
    )
    return _paragraph_table(df, [1.0 * inch, 2.0 * inch, 4.4 * inch], styles)


def _domain_definitions_table(domain_scores: Dict[str, float], styles) -> Table:
    rows = [
        [
            "Service Resilience",
            f"{round(_safe_float(domain_scores.get('Service Resilience', 0.0)), 1)}",
            "Ability to recover quickly and consistently from incidents.",
        ],
        [
            "Change Governance",
            f"{round(_safe_float(domain_scores.get('Change Governance', 0.0)), 1)}",
            "Effectiveness of controls preventing operational instability during change activity.",
        ],
        [
            "Structural Risk Debt™",
            f"{round(_safe_float(domain_scores.get('Structural Risk Debt™', 0.0)), 1)}",
            "Accumulated instability caused by unresolved recurring weaknesses and operational debt.",
        ],
        [
            "Reliability Momentum",
            f"{round(_safe_float(domain_scores.get('Reliability Momentum', 0.0)), 1)}",
            "Direction of operational reliability based on recurring instability and recovery behavior.",
        ],
    ]
    df = pd.DataFrame(rows, columns=["Domain", "Score", "What It Means"])
    return _paragraph_table(df, [2.0 * inch, 0.7 * inch, 4.7 * inch], styles)


def _prepare_sip_candidates_table(sip_candidates: pd.DataFrame) -> pd.DataFrame:
    df = _safe_df(sip_candidates)
    if df.empty:
        return pd.DataFrame(columns=["Service", "Tier", "Theme", "Priority", "Why Flagged", "Score"])

    mappings = [
        ("Service", ["Service", "Service_Name", "service"]),
        ("Tier", ["Service_Tier", "Tier", "service_tier"]),
        ("Theme", ["Suggested_Theme", "Theme", "suggested_theme"]),
        ("Priority", ["Priority_Label", "Priority", "priority", "Priority_Level"]),
        ("Why Flagged", ["Why_Flagged", "Why Leadership Should Care", "Reason", "reason"]),
        ("Score", ["SIP_Priority_Score", "Risk_Score", "Risk", "risk_score", "Score"]),
    ]

    out = pd.DataFrame()
    for target, options in mappings:
        found = next((c for c in options if c in df.columns), None)
        if found:
            out[target] = df[found]
        else:
            out[target] = ""

    if "Score" in out.columns:
        out["Score"] = out["Score"].apply(lambda x: f"{_safe_float(x, 0.0):.1f}" if str(x).strip() != "" else "")

    return out.head(10)


def _top_three_briefs(sip_candidates: pd.DataFrame) -> pd.DataFrame:
    df = _safe_df(sip_candidates)
    if df.empty:
        return pd.DataFrame(columns=["Initiative", "Why Leadership Should Care"])

    rows = []
    for _, row in df.head(3).iterrows():
        svc = _safe_str(row.get("Service", row.get("Service_Name", "Unknown Service")))
        tier = _safe_str(row.get("Service_Tier", row.get("Tier", "Unknown Tier")))
        theme = _safe_str(row.get("Suggested_Theme", row.get("Theme", "Stability Improvement")))
        why = _safe_str(
            row.get(
                "Why_Flagged",
                row.get("Why Leadership Should Care", "Elevated operational instability"),
            )
        )
        rows.append(
            {
                "Initiative": f"{svc} ({tier}) — {theme}",
                "Why Leadership Should Care": why,
            }
        )
    return pd.DataFrame(rows)


def _derive_key_takeaways(domain_scores: Dict[str, float], sip_candidates: pd.DataFrame) -> pd.DataFrame:
    ordered = sorted(
        [(k, _safe_float(v)) for k, v in (domain_scores or {}).items()],
        key=lambda x: x[1],
        reverse=True,
    )

    strongest = ordered[0][0] if ordered else "Change Governance"
    weakest = ordered[-1][0] if ordered else "Structural Risk Debt™"

    sip_df = _safe_df(sip_candidates)
    if not sip_df.empty:
        top_service = _safe_str(sip_df.iloc[0].get("Service", "Top service"))
        priority = _safe_str(sip_df.iloc[0].get("Priority_Label", sip_df.iloc[0].get("Priority", "Priority watch")))
        third = f"{top_service} is the leading service candidate for action and is currently classified as <b>{priority}</b>."
    else:
        third = "No SIP candidates were available in the current dataset, which may indicate limited service-level instability visibility."

    rows = [
        {
            "Key Takeaway": f"Operational strength is most visible in <b>{strongest}</b>, which currently provides the strongest stability support."
        },
        {
            "Key Takeaway": f"The primary exposure area is <b>{weakest}</b>, and it should anchor near-term stabilization decisions."
        },
        {
            "Key Takeaway": third
        },
    ]
    return pd.DataFrame(rows)


def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    expected_order = [
        "Service Resilience",
        "Change Governance",
        "Structural Risk Debt™",
        "Reliability Momentum",
    ]

    labels = expected_order
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]

    fig = plt.figure(figsize=(4.8, 4.25), dpi=180)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.12)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
    ax.set_title("Operational Stability Radar", pad=16, fontsize=11)

    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png", bbox_inches="tight", facecolor="white")
    plt.close(fig)
    img.seek(0)
    return img


def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    styles = _styles()
    out = io.BytesIO()

    doc = SimpleDocTemplate(
        out,
        pagesize=LETTER,
        rightMargin=0.55 * inch,
        leftMargin=0.55 * inch,
        topMargin=0.45 * inch,
        bottomMargin=0.55 * inch,
    )

    tenant_name = _safe_str(_first_present(payload, ["tenant_name", "organization", "client_name"], "Default"))
    as_of = _safe_str(_first_present(payload, ["as_of", "report_date", "generated_at"], ""))
    bvsi = _safe_float(_first_present(payload, ["bvsi", "BVSI", "overall_score"], 0.0))
    posture = _safe_str(_first_present(payload, ["posture", "operating_posture", "maturity_posture"], "Unknown"))
    executive_interpretation = _clean_text(
        _first_present(
            payload,
            ["executive_interpretation", "executive_summary", "summary", "leadership_summary"],
            "",
        )
    )

    domain_scores = _first_present(payload, ["domain_scores", "domains"], {}) or {}
    sip_candidates = _safe_df(_first_present(payload, ["sip_candidates", "sips", "sip_table"], pd.DataFrame()))
    detected_dataset = _safe_str(_first_present(payload, ["detected_dataset", "dataset_type"], "unknown")).upper()
    service_anchor_used = _safe_str(_first_present(payload, ["service_anchor_used", "service_anchor"], "None"))
    data_readiness_score = _safe_float(_first_present(payload, ["data_readiness_score", "readiness_score"], 0.0))

    story = []

    story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
    story.append(Paragraph(f"Operational Stability Intelligence Report — {tenant_name}", styles["OSIL_Subtitle"]))
    if as_of:
        story.append(Paragraph(f"As of {as_of}", styles["OSIL_Small"]))
    story.append(Spacer(1, 8))

    story.append(_header_band("Executive Brief"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    metric_row = Table(
        [[
            _kpi_box("BVSI™ Score", f"{bvsi:.1f}", styles, is_numeric=True),
            _kpi_box("Operating Posture", posture, styles, is_numeric=False),
            _kpi_box("Data Readiness", f"{data_readiness_score:.1f}%", styles, is_numeric=True),
        ]],
        colWidths=[2.45 * inch, 2.45 * inch, 2.45 * inch],
    )
    metric_row.setStyle(
        TableStyle(
            [
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    story.append(metric_row)

    story.append(Spacer(1, 10))
    story.append(Paragraph("BVSI™ Interpretation", styles["OSIL_Section"]))
    story.append(_bvsi_scale_table(styles))

    story.append(Spacer(1, 10))
    story.append(_info_box("Executive Signal", _posture_signal_text(bvsi, posture), styles))

    if executive_interpretation:
        story.append(Spacer(1, 10))
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(_full_width_paragraph(executive_interpretation, styles, "OSIL_Body"))

    story.append(Spacer(1, 10))
    story.append(Paragraph("Key Takeaways", styles["OSIL_Section"]))
    key_takeaways_df = _derive_key_takeaways(domain_scores, sip_candidates)
    story.append(_paragraph_table(key_takeaways_df, None, styles))

    story.append(Spacer(1, 8))
    context_text = (
        f"Detected dataset: <b>{detected_dataset}</b> • "
        f"Service anchor used: <b>{service_anchor_used}</b> • "
        f"Organization: <b>{tenant_name}</b>"
    )
    story.append(_full_width_paragraph(context_text, styles, "OSIL_Small"))

    story.append(PageBreak())
    story.append(_header_band("Operational Stability Profile"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    radar_img = _build_radar_image(domain_scores)
    story.append(Image(radar_img, width=4.95 * inch, height=4.35 * inch))
    story.append(Spacer(1, 8))

    story.append(
        _full_width_paragraph(
            "How to read this chart: higher scores indicate stronger operational stability. "
            "The radar shows current performance across service resilience, change governance, "
            "structural risk debt, and reliability momentum.",
            styles,
            "OSIL_Note",
        )
    )

    story.append(Spacer(1, 10))
    story.append(Paragraph("Domain Scores", styles["OSIL_Section"]))
    story.append(_domain_definitions_table(domain_scores, styles))

    story.append(Spacer(1, 8))
    story.append(
        _full_width_paragraph(
            "Score guide: 80–100 = strong maturity; 60–79 = controlled but improving; "
            "40–59 = operational weakness; below 40 = structural instability.",
            styles,
            "OSIL_Small",
        )
    )

    story.append(PageBreak())
    story.append(_header_band("Service Improvement Priorities"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    story.append(
        _full_width_paragraph(
            "This page highlights the highest-priority Service Improvement Programs (SIPs) "
            "based on recurrence, MTTR drag, reopen churn, and change-related instability.",
            styles,
            "OSIL_Body",
        )
    )
    story.append(Spacer(1, 10))

    story.append(Paragraph("Top 3 Initiatives to Brief Leadership", styles["OSIL_Section"]))
    story.append(_paragraph_table(_top_three_briefs(sip_candidates), [3.1 * inch, 4.3 * inch], styles))

    story.append(Spacer(1, 12))
    story.append(Paragraph("Service Improvement Priorities (SIPs)", styles["OSIL_Section"]))
    detailed = _prepare_sip_candidates_table(sip_candidates)
    story.append(
        _paragraph_table(
            detailed,
            [1.45 * inch, 0.72 * inch, 1.3 * inch, 0.9 * inch, 2.3 * inch, 0.73 * inch],
            styles,
        )
    )

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    out.seek(0)
    return out.getvalue()
