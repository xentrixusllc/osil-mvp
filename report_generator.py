from __future__ import annotations

from io import BytesIO
from typing import Dict, Any, List, Tuple

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
    Image,
    PageBreak,
)
from reportlab.lib.enums import TA_RIGHT


# -----------------------------
# Brand
# -----------------------------
NAVY = colors.HexColor("#0A192F")
TEAL = colors.HexColor("#64FFDA")
DARK = colors.HexColor("#111827")
MID = colors.HexColor("#374151")
LIGHT = colors.HexColor("#F5F7FA")
BORDER = colors.HexColor("#D1D5DB")
WHITE = colors.white


# -----------------------------
# Utilities
# -----------------------------
def _safe_df(obj) -> pd.DataFrame:
    if obj is None:
        return pd.DataFrame()
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    return pd.DataFrame(obj)


def _posture_chip_colors(posture: str) -> Tuple[colors.Color, colors.Color]:
    posture = (posture or "").strip()

    if posture == "High Confidence Operations":
        return colors.HexColor("#D1FAE5"), colors.HexColor("#065F46")

    if posture == "Controlled but Exposed":
        return colors.HexColor("#FEF3C7"), colors.HexColor("#92400E")

    if posture == "Fragile Stability":
        return colors.HexColor("#FFE4E6"), colors.HexColor("#9F1239")

    return colors.HexColor("#E5E7EB"), colors.HexColor("#111827")


def _bvsi_scale_lines(bvsi: float) -> List[str]:
    return [
        "<b>BVSI™ Scale (Business Value Stability Index)</b>",
        "80–100: High Confidence Operations",
        "60–79: Controlled but Exposed",
        "40–59: Fragile Stability",
        "0–39: Reactive Instability",
        f"<b>Current BVSI™:</b> {bvsi:.1f}",
    ]


# -----------------------------
# Radar Chart
# -----------------------------
def _radar_png(domain_scores: Dict[str, float]) -> BytesIO:

    labels = [
        "Service Resilience",
        "Change Governance",
        "Structural Risk Debt™",
        "Reliability Momentum",
    ]

    values = [
        float(domain_scores.get("Overall Service Resilience", 0)),
        float(domain_scores.get("Overall Change Governance", 0)),
        float(domain_scores.get("Overall Structural Risk Debt", 0)),
        float(domain_scores.get("Overall Reliability Momentum", 0)),
    ]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()

    values_loop = values + values[:1]
    angles_loop = angles + angles[:1]

    fig = plt.figure(figsize=(6.5, 5), dpi=260)
    ax = plt.subplot(111, polar=True)

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)

    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)

    ax.set_ylim(0, 100)

    ax.grid(color="#D1D5DB", linewidth=0.8)

    ax.plot(angles_loop, values_loop, linewidth=2.4, color="#0A192F")
    ax.fill(angles_loop, values_loop, alpha=0.12, color="#64FFDA")

    ax.set_title("Operational Stability Profile", fontsize=11, pad=14)

    buf = BytesIO()

    plt.tight_layout()

    fig.savefig(buf, format="png", bbox_inches="tight", dpi=260)

    plt.close(fig)

    buf.seek(0)

    return buf


# -----------------------------
# Header / Footer
# -----------------------------
def _draw_header(canvas, doc, title: str, subtitle: str, as_of: str):

    w, h = LETTER
    band_h = 0.85 * inch

    canvas.saveState()

    canvas.setFillColor(NAVY)
    canvas.rect(0, h - band_h, w, band_h, fill=1, stroke=0)

    canvas.setFillColor(TEAL)
    canvas.rect(0, h - band_h, w, 0.06 * inch, fill=1, stroke=0)

    canvas.setFillColor(WHITE)

    canvas.setFont("Helvetica-Bold", 18)
    canvas.drawString(0.75 * inch, h - 0.55 * inch, title)

    canvas.setFont("Helvetica", 10)
    canvas.drawString(0.75 * inch, h - 0.72 * inch, subtitle)

    canvas.setFont("Helvetica", 9)
    canvas.drawRightString(w - 0.75 * inch, h - 0.72 * inch, f"As-of: {as_of}")

    canvas.setFillColor(MID)

    canvas.setFont("Helvetica", 8.5)
    canvas.drawString(
        0.75 * inch,
        0.55 * inch,
        "Operational Stability Intelligence (OSIL™) • BVSI™ • Structural Risk Debt™",
    )

    canvas.drawRightString(
        w - 0.75 * inch,
        0.55 * inch,
        "© Xentrixus • Confidential",
    )

    canvas.restoreState()


# -----------------------------
# Report Builder
# -----------------------------
def build_osil_pdf_report(results: Dict[str, Any]) -> bytes:

    overall = results.get("overall", {})
    bvsi = float(overall.get("BVSI", 0))

    posture = (results.get("posture", "") or "").strip()

    as_of = results.get("as_of", "")

    sip_df = _safe_df(results.get("sip_table"))

    if not sip_df.empty:
        sip_df = sip_df.head(10)

    styles = getSampleStyleSheet()

    base = styles["BodyText"]
    base.fontSize = 10.5
    base.leading = 13.5
    base.textColor = DARK

    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        fontSize=12.5,
        textColor=DARK,
        spaceAfter=6,
    )

    small = ParagraphStyle(
        "Small",
        parent=base,
        fontSize=9.5,
        textColor=MID,
    )

    right = ParagraphStyle(
        "Right",
        parent=base,
        alignment=TA_RIGHT,
    )

    score_right = ParagraphStyle(
        "ScoreRight",
        parent=right,
        fontName="Helvetica-Bold",
    )

    out = BytesIO()

    doc = SimpleDocTemplate(
        out,
        pagesize=LETTER,
        leftMargin=0.75 * inch,
        rightMargin=0.75 * inch,
        topMargin=1.20 * inch,
        bottomMargin=0.85 * inch,
    )

    story: List[Any] = []

    chip_fill, chip_text = _posture_chip_colors(posture)

    # BVSI Block

    bvsi_block = Paragraph(
        f"<b style='font-size:28px'>{bvsi:.1f}</b><br/><font size='10'>BVSI™</font>",
        base,
    )

    posture_style = ParagraphStyle("chip", parent=base, textColor=chip_text)

    posture_block = Table(
        [[Paragraph(f"<b>Operating Posture:</b> {posture}", posture_style)]],
        colWidths=[3.6 * inch],
    )

    posture_block.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), chip_fill),
                ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )

    top_row = Table([[bvsi_block, posture_block]], colWidths=[2.2 * inch, 4.55 * inch])

    story.append(top_row)
    story.append(Spacer(1, 10))

    # BVSI interpretation

    interp_lines = "<br/>".join(_bvsi_scale_lines(bvsi))

    interp_tbl = Table([[Paragraph(interp_lines, base)]], colWidths=[6.75 * inch])

    interp_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ]
        )
    )

    story.append(Paragraph("BVSI™ Interpretation", h2))
    story.append(interp_tbl)
    story.append(Spacer(1, 12))

    # -----------------------------
    # Executive Signal (short)
    # -----------------------------

    signal_text = (
        f"Operational stability is <b>{posture}</b>, with recurring disruptions in "
        "Tier-1 services limiting reliability gains despite strong governance. "
        "Immediate Service Improvement Programs targeting the highest-impact "
        "services will deliver the fastest improvement in operational stability."
    )

    story.append(Paragraph("Executive Signal", h2))
    story.append(Paragraph(signal_text, base))
    story.append(Spacer(1, 12))

    # -----------------------------
    # Executive Summary (expanded)
    # -----------------------------

    summary_text = (
        f"Your organization is operating in a <b>{posture}</b> stability posture "
        f"with a BVSI™ score of <b>{bvsi:.1f}</b>. While governance mechanisms "
        "such as change management appear strong, recurring incident patterns "
        "and structural risk debt across Tier-1 services continue to create "
        "operational exposure. These patterns suggest that instability is driven "
        "less by governance breakdowns and more by underlying reliability "
        "challenges within critical services. Targeted Service Improvement "
        "Programs focused on the highest-impact services can materially reduce "
        "recurrence, improve recovery performance, and strengthen operational "
        "confidence over the next <b>30–60 days</b>."
    )

    story.append(Paragraph("Executive Summary", h2))
    story.append(Paragraph(summary_text, base))
    story.append(Spacer(1, 10))

    bullets = [
        "Treat Tier-1 services as the primary stability surface area.",
        "Convert recurring incidents into prevention programs.",
        "Strong governance should now translate into reliability gains.",
    ]

    story.append(Paragraph("Key Takeaways", h2))

    for b in bullets:
        story.append(Paragraph(f"• {b}", base))

    story.append(PageBreak())

    # Radar Page

    story.append(Paragraph("Operational Stability Profile", h2))

    radar_buf = _radar_png(overall)

    radar_img = Image(radar_buf, width=6.5 * inch, height=4.5 * inch)

    story.append(radar_img)

    story.append(
        Paragraph(
            "Balanced shape indicates aligned governance. Collapsed axis highlights stability gaps.",
            small,
        )
    )

    story.append(PageBreak())

    # SIP Table

    story.append(Paragraph("Service Improvement Priorities (SIPs)", h2))

    if sip_df.empty:

        story.append(Paragraph("No SIP candidates available.", small))

    else:

        header = ["Service", "Tier", "Theme", "Priority", "Why Flagged", "Score"]

        rows = [header]

        for _, r in sip_df.iterrows():

            rows.append(
                [
                    str(r.get("Service", "")),
                    str(r.get("Service_Tier", "")),
                    str(r.get("Suggested_Theme", "")),
                    str(r.get("Priority_Label", "")),
                    str(r.get("Why_Flagged", "")),
                    str(r.get("SIP_Priority_Score", "")),
                ]
            )

        sip_table = Table(rows, repeatRows=1)

        sip_table.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
                    ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                    ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
                ]
            )
        )

        story.append(sip_table)

    def on_page(canvas, doc_):

        page = canvas.getPageNumber()

        if page == 1:
            subtitle = "Operational Stability Intelligence Report"

        elif page == 2:
            subtitle = "Operational Stability Profile"

        else:
            subtitle = "Service Improvement Priorities"

        _draw_header(canvas, doc_, "OSIL™ by Xentrixus", subtitle, as_of)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    out.seek(0)

    return out.getvalue()
