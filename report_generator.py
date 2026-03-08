# ===============================
# OSIL Executive PDF Generator
# report_generator.py
# ===============================

import io
import re
from typing import Any, Dict, List

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


# -------------------------------
# Helpers
# -------------------------------
def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None:
            return default
        return float(val)
    except Exception:
        return default


def _safe_df(val: Any) -> pd.DataFrame:
    return val.copy() if isinstance(val, pd.DataFrame) else pd.DataFrame()


def _clean_text(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    s = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", s)
    s = s.replace("<br>", "<br/>").replace("<br />", "<br/>")
    return s


def _styles():
    styles = getSampleStyleSheet()

    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=6,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#555555"),
            spaceAfter=10,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=17,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=8,
            spaceBefore=2,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.5,
            leading=13,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#666666"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Table",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=7.8,
            leading=9.4,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.black,
        )
    )

    return styles


def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(0.55 * inch, 0.35 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
    canvas.drawRightString(7.95 * inch, 0.35 * inch, "Confidential")
    canvas.restoreState()


def _header_band(text: str, width: float = 7.35 * inch) -> Table:
    t = Table([[text]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0A192F")),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 12),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return t


def _accent_rule(width: float = 7.35 * inch, color_hex: str = "#64FFDA", height: int = 3) -> Table:
    t = Table([[""]], colWidths=[width], rowHeights=[height])
    t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(color_hex))]))
    return t


def _paragraph_table(df: pd.DataFrame, col_widths: List[float], styles, header_bg: str = "#E9EDF3") -> Table:
    df = _safe_df(df)

    if df.empty:
        data = [[Paragraph("No data available", styles["OSIL_Table"])]]
        tbl = Table(data, colWidths=[sum(col_widths)])
        tbl.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.4, colors.grey)]))
        return tbl

    header = [Paragraph(_clean_text(c), styles["OSIL_TableHeader"]) for c in df.columns]
    rows = [[Paragraph(_clean_text(x), styles["OSIL_Table"]) for x in row] for row in df.values.tolist()]
    tbl = Table([header] + rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return tbl


def _kpi_box(label: str, value: str, width: float = 2.35 * inch) -> Table:
    t = Table([[label], [value]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F7FA")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, 1), 17),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#666666")),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#111111")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
            ]
        )
    )
    return t


# -------------------------------
# Visuals
# -------------------------------
def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    labels = list(domain_scores.keys())
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]

    if not labels:
        labels = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
        values = [0, 0, 0, 0]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]

    fig = plt.figure(figsize=(4.4, 4.1), dpi=180)
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
    ax.set_title("Operational Stability Profile", pad=16, fontsize=11)

    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png", bbox_inches="tight")
    plt.close(fig)
    img.seek(0)
    return img


def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> io.BytesIO | None:
    df = _safe_df(service_risk_top10)
    if df.empty:
        return None

    if "Service_Tier" not in df.columns:
        df["Service_Tier"] = "Unspecified"

    required = ["Service", "Service_Tier", "Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    for col in required:
        if col not in df.columns:
            return None

    hm = df.head(10).copy()
    hm.index = hm["Service"].astype(str) + " (" + hm["Service_Tier"].astype(str) + ")"
    hm = hm[
        ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    ].rename(
        columns={
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision",
        }
    )
    hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)

    fig = plt.figure(figsize=(6.8, 3.9), dpi=180)
    ax = plt.gca()
    im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100)
    ax.set_xticks(range(len(hm.columns)))
    ax.set_xticklabels(list(hm.columns), fontsize=8)
    ax.set_yticks(range(len(hm.index)))
    ax.set_yticklabels(list(hm.index), fontsize=8)
    ax.set_title("Service Stability Heatmap", fontsize=11)

    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            ax.text(j, i, f"{int(round(float(hm.iat[i, j]), 0))}", ha="center", va="center", fontsize=8)

    cbar = plt.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
    cbar.set_label("Risk Score", fontsize=8)
    cbar.ax.tick_params(labelsize=7)

    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png", bbox_inches="tight")
    plt.close(fig)
    img.seek(0)
    return img


# -------------------------------
# PDF Builder
# -------------------------------
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

    # Support both direct engine payload and app-built payload
    if "executive_interpretation" not in payload and "exec_text" in payload:
        payload = {
            "bvsi": payload.get("bvsi", 0.0),
            "posture": payload.get("posture", "Unknown"),
            "as_of": payload.get("as_of", ""),
            "executive_interpretation": payload.get("exec_text", ""),
            "domain_scores": payload.get("domain_scores", {}),
            "service_risk_top10": payload.get("top10", pd.DataFrame()),
            "sip_candidates": payload.get("sip_view", pd.DataFrame()),
            "data_readiness_score": payload.get("readiness_score", 0.0),
            "service_anchor_used": payload.get("anchor_used", "Service"),
            "detected_dataset": payload.get("practice_type", "INCIDENT"),
            "tenant_name": payload.get("tenant_name", "Default"),
            "cmdb_alignment_text": payload.get("cmdb_alignment_text", ""),
            "metric_definitions": payload.get("metric_definitions", []),
            "score_interpretation": payload.get("score_interpretation", []),
            "action_guidance": payload.get("action_guidance", []),
        }

    tenant_name = str(payload.get("tenant_name", "Default"))
    as_of = str(payload.get("as_of", ""))
    bvsi = _safe_float(payload.get("bvsi", 0.0))
    posture = str(payload.get("posture", "Unknown"))
    executive_interpretation = _clean_text(payload.get("executive_interpretation", ""))

    domain_scores = payload.get("domain_scores", {}) or {}
    service_risk_top10 = _safe_df(payload.get("service_risk_top10"))
    sip_candidates = _safe_df(payload.get("sip_candidates"))

    detected_dataset = str(payload.get("detected_dataset", "unknown")).upper()
    service_anchor_used = str(payload.get("service_anchor_used", "None"))
    data_readiness_score = _safe_float(payload.get("data_readiness_score", 0.0))

    cmdb_alignment_text = _clean_text(payload.get("cmdb_alignment_text", ""))
    metric_definitions = pd.DataFrame(payload.get("metric_definitions", []))
    score_interpretation = pd.DataFrame(payload.get("score_interpretation", []))
    action_guidance = pd.DataFrame(payload.get("action_guidance", []))

    story = []

    # Page 1
    story.append(Paragraph("Operational Stability Intelligence (OSIL™)", styles["OSIL_Title"]))
    story.append(Paragraph(f"Executive Report — {tenant_name}", styles["OSIL_Subtitle"]))
    story.append(Paragraph(f"As of {as_of}", styles["OSIL_Small"]))
    story.append(Spacer(1, 6))

    story.append(_header_band("Executive Stability Brief"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    story.append(
        Table(
            [[
                _kpi_box("BVSI™ Score", f"{bvsi:.1f}"),
                _kpi_box("Operating Posture", posture),
                _kpi_box("Data Readiness", f"{data_readiness_score:.1f}%"),
            ]],
            colWidths=[2.45 * inch, 2.45 * inch, 2.45 * inch],
        )
    )

    story.append(Spacer(1, 10))
    story.append(Paragraph("Executive Interpretation", styles["OSIL_Section"]))
    story.append(Paragraph(executive_interpretation, styles["OSIL_Body"]))

    story.append(Spacer(1, 10))
    context_df = pd.DataFrame(
        [
            ["Detected Dataset", detected_dataset],
            ["Service Anchor Used", service_anchor_used],
            ["Organization", tenant_name],
        ],
        columns=["Field", "Value"],
    )
    story.append(Paragraph("Assessment Context", styles["OSIL_Section"]))
    story.append(_paragraph_table(context_df, [2.2 * inch, 5.0 * inch], styles))

    # Page 2
    story.append(PageBreak())
    story.append(_header_band("Operational Stability Diagnostics"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    story.append(Image(_build_radar_image(domain_scores), width=4.75 * inch, height=4.4 * inch))
    story.append(Spacer(1, 8))

    domain_df = pd.DataFrame({"Domain": list(domain_scores.keys()), "Score": list(domain_scores.values())})
    story.append(_paragraph_table(domain_df, [4.7 * inch, 2.5 * inch], styles))

    story.append(Spacer(1, 8))
    story.append(Paragraph("How to Read the Scores", styles["OSIL_Section"]))
    story.append(_paragraph_table(metric_definitions, [1.8 * inch, 2.6 * inch, 3.0 * inch], styles))

    story.append(Spacer(1, 8))
    story.append(Paragraph("Score Interpretation Guide", styles["OSIL_Section"]))
    story.append(_paragraph_table(score_interpretation, [1.2 * inch, 2.0 * inch, 4.2 * inch], styles))

    # Page 3
    story.append(PageBreak())
    story.append(_header_band("Service Model & Action Guidance"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    story.append(Paragraph("Service Model & CMDB Alignment", styles["OSIL_Section"]))
    story.append(
        Paragraph(
            cmdb_alignment_text or "No service model alignment advisory available.",
            styles["OSIL_Body"],
        )
    )

    story.append(Spacer(1, 10))
    story.append(Paragraph("Executive Action Lens", styles["OSIL_Section"]))
    story.append(_paragraph_table(action_guidance, [1.8 * inch, 5.4 * inch], styles))

    # Page 4
    story.append(PageBreak())
    story.append(_header_band("Service Stability Heatmap"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    hm_img = _build_heatmap_image(service_risk_top10)
    if hm_img is not None:
        story.append(Image(hm_img, width=6.8 * inch, height=4.2 * inch))
    else:
        story.append(Paragraph("No service risk data available.", styles["OSIL_Body"]))

    story.append(Spacer(1, 8))
    story.append(
        Paragraph(
            "Heatmap Reading Guide: services with elevated values across multiple columns typically represent concentrated stability exposure and should be prioritized for leadership attention or SIP action.",
            styles["OSIL_Small"],
        )
    )

    if not service_risk_top10.empty:
        story.append(Spacer(1, 10))
        story.append(Paragraph("Top 10 Services — Risk Breakdown", styles["OSIL_Section"]))
        cols = list(service_risk_top10.columns)
        width = 7.2 / max(len(cols), 1)
        story.append(_paragraph_table(service_risk_top10, [width * inch] * len(cols), styles))

    # Page 5
    story.append(PageBreak())
    story.append(_header_band("Service Improvement Programs (SIPs)"))
    story.append(_accent_rule())
    story.append(Spacer(1, 10))

    story.append(_paragraph_table(sip_candidates, [1.4 * inch, 0.9 * inch, 1.5 * inch, 1.0 * inch, 0.8 * inch, 2.75 * inch], styles))

    doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
    out.seek(0)
    return out.getvalue()
