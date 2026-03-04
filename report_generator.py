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
# Brand (Consulting Style)
# -----------------------------
NAVY = colors.HexColor("#0A192F")
TEAL = colors.HexColor("#64FFDA")
DARK = colors.HexColor("#111827")
MID = colors.HexColor("#374151")
LIGHT = colors.HexColor("#F5F7FA")
BORDER = colors.HexColor("#D1D5DB")
WHITE = colors.white


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


def _radar_png(domain_scores: Dict[str, float]) -> BytesIO:
    """
    Crisp, non-stretched radar chart for PDF embedding.
    - Higher DPI for clarity
    - Taller aspect ratio for a rounder look
    - White background
    """
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

    fig = plt.figure(figsize=(6.6, 5.1), dpi=260)
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

    ax.grid(color="#D1D5DB", linewidth=0.8, alpha=0.95)

    ax.plot(angles_loop, values_loop, linewidth=2.4, color="#0A192F")
    ax.fill(angles_loop, values_loop, alpha=0.12, color="#64FFDA")

    ax.set_title("Operational Stability Profile", fontsize=11, pad=14)

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight", dpi=260)
    plt.close(fig)
    buf.seek(0)
    return buf


def _draw_header(canvas, doc, title: str, subtitle: str, as_of: str) -> None:
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

    # Footer
    canvas.setFillColor(MID)
    canvas.setFont("Helvetica", 8.5)
    canvas.drawString(
        0.75 * inch,
        0.55 * inch,
        "Operational Stability Intelligence (OSIL™)  •  BVSI™  •  Structural Risk Debt™",
    )
    canvas.drawRightString(w - 0.75 * inch, 0.55 * inch, "© Xentrixus • Confidential")

    canvas.restoreState()


def build_osil_pdf_report(results: Dict[str, Any]) -> bytes:
    overall = results.get("overall", {})
    bvsi = float(overall.get("BVSI", 0))
    posture = (results.get("posture", "") or "").strip()
    as_of = results.get("as_of", "")

    sip_df = _safe_df(results.get("sip_table"))
    if not sip_df.empty:
        sip_df = sip_df.head(10).copy()

    # -----------------------------
    # Styles
    # -----------------------------
    styles = getSampleStyleSheet()

    base = styles["BodyText"]
    base.fontName = "Helvetica"
    base.fontSize = 10.5
    base.leading = 13.5
    base.textColor = DARK

    h2 = ParagraphStyle(
        "H2",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=12.5,
        textColor=DARK,
        spaceAfter=6,
    )

    small = ParagraphStyle(
        "Small",
        parent=base,
        fontSize=9.5,
        leading=12,
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

    # -----------------------------
    # Document
    # -----------------------------
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

    # =========================
    # Page 1 — Executive Brief
    # =========================
    chip_fill, chip_text = _posture_chip_colors(posture)

    bvsi_block = Paragraph(
        f"<b style='font-size:28px'>{bvsi:.1f}</b><br/><font size='10'>BVSI™ (Business Value Stability Index)</font>",
        base,
    )

    posture_style = ParagraphStyle("chip", parent=base, textColor=chip_text)
    posture_block = Table(
        [[Paragraph(f"<b>Operating Posture:</b> {posture}", posture_style)]],
        colWidths=[3.6 * inch],
        rowHeights=[0.45 * inch],
    )
    posture_block.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), chip_fill),
                ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ]
        )
    )

    top_row = Table([[bvsi_block, posture_block]], colWidths=[2.2 * inch, 4.55 * inch])
    top_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
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
                ("TOPPADDING", (0, 0), (-1, -1), 8),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ]
        )
    )
    story.append(Paragraph("BVSI™ Interpretation", h2))
    story.append(interp_tbl)
    story.append(Spacer(1, 12))

    # Executive Signal (NEW)
    signal_text = (
        f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of "
        f"<b>{bvsi:.1f}</b> indicating that governance mechanisms are functioning but "
        "recurring instability patterns remain across Tier-1 services. Focused "
        "Service Improvement Programs targeting the highest-impact services will "
        "deliver the greatest stability improvement over the next <b>30–60 days</b>."
    )
    story.append(Paragraph("Executive Signal", h2))
    story.append(Paragraph(signal_text, base))
    story.append(Spacer(1, 12))

    # Executive Summary
    summary_text = (
        f"Your organization is operating in a <b>{posture}</b> stability posture based on a BVSI™ score of "
        f"<b>{bvsi:.1f}</b>. Operational control exists, but recurring instability patterns still create exposure—"
        "especially across higher-impact services. Targeted Service Improvement Programs (SIPs) in the next 30–60 "
        "days can increase executive confidence and preserve customer trust during disruption events."
    )
    story.append(Paragraph("Executive Summary", h2))
    story.append(Paragraph(summary_text, base))
    story.append(Spacer(1, 10))

    # Key Takeaways
    bullets = [
        "Treat Tier-1 services as the primary stability surface area due to customer and business impact concentration.",
        "Use SIPs to convert recurring incident themes into prevention work with owners, due dates, and verification steps.",
        "Strong change governance enables reliability improvements to drive the next BVSI™ uplift.",
    ]
    story.append(Paragraph("Key Takeaways (Leadership View)", h2))
    for b in bullets:
        story.append(Paragraph(f"• {b}", base))

    story.append(PageBreak())

    # =========================
    # Page 2 — Radar + Diagnostics
    # =========================
    story.append(Paragraph("Operational Stability Profile (Radar)", h2))

    radar_buf = _radar_png(overall)

    # Aspect-preserving sizing (prevents stretch)
    RADAR_W = 6.55 * inch
    RADAR_ASPECT = 0.78
    radar_img = Image(radar_buf, width=RADAR_W, height=RADAR_W * RADAR_ASPECT)

    story.append(radar_img)
    story.append(Spacer(1, 6))
    story.append(
        Paragraph(
            "How to read: a balanced shape suggests aligned governance; a collapsed axis indicates a concentrated stability gap.",
            small,
        )
    )
    story.append(Spacer(1, 12))

    story.append(Paragraph("Domain Scores (0–100)", h2))
    domains = [
        ("Service Resilience", float(overall.get("Overall Service Resilience", 0))),
        ("Change Governance", float(overall.get("Overall Change Governance", 0))),
        ("Structural Risk Debt™", float(overall.get("Overall Structural Risk Debt", 0))),
        ("Reliability Momentum", float(overall.get("Overall Reliability Momentum", 0))),
    ]

    dom_rows = [[Paragraph("<b>Domain</b>", base), Paragraph("<b>Score (0–100)</b>", right)]]
    for name, score in domains:
        dom_rows.append([Paragraph(name, base), Paragraph(f"{score:.1f}", score_right)])

    dom_tbl = Table(dom_rows, colWidths=[4.9 * inch, 1.85 * inch])
    dom_tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
                ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
                ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]
        )
    )
    story.append(dom_tbl)

    story.append(PageBreak())

    # =========================
    # Page 3 — Findings + Actions
    # =========================
    story.append(Paragraph("Key Stability Findings", h2))
    findings = [
        "Tier-1 services should be treated as the primary stability surface area due to business impact concentration.",
        "Structural Risk Debt™ indicates recurring instability patterns that are not being fully eliminated through prevention work.",
        "Strong change governance reduces release-driven volatility, enabling targeted reliability improvements to drive the next BVSI™ lift.",
    ]
    for f in findings:
        story.append(Paragraph(f"• {f}", base))
    story.append(Spacer(1, 10))

    story.append(Paragraph("Recommended Next Actions (Next 30 Days)", h2))
    actions = [
        "Launch SIPs against the top recurring instability clusters across Tier-1 services (Service + Theme).",
        "Reduce reopen churn by strengthening closure verification and linking repeat themes to prevention work with owners and due dates.",
        "Improve recovery readiness for Tier-1 services through runbooks, alert quality tuning, and escalation path clarity.",
    ]
    for i, a in enumerate(actions, start=1):
        story.append(Paragraph(f"{i}. {a}", base))

    story.append(PageBreak())

    # =========================
    # Page 4 — SIP Priorities
    # =========================
    story.append(Paragraph("Service Improvement Priorities (SIPs)", h2))
    story.append(
        Paragraph(
            "SIPs below are prioritized by Service + Theme and tier exposure. These represent the most actionable stability initiatives "
            "to improve BVSI™ and reduce avoidable disruption.",
            base,
        )
    )
    story.append(Spacer(1, 10))

    if sip_df.empty:
        story.append(Paragraph("<i>No SIP candidates available from the current dataset.</i>", small))
    else:
        required_cols = [
            "Service",
            "Service_Tier",
            "Suggested_Theme",
            "Priority_Label",
            "Why_Flagged",
            "SIP_Priority_Score",
        ]
        for col in required_cols:
            if col not in sip_df.columns:
                sip_df[col] = ""

        top3 = sip_df.head(3)
        story.append(Paragraph("Top 3 Initiatives to Brief Leadership", h2))
        for _, r in top3.iterrows():
            story.append(
                Paragraph(
                    f"• <b>{r['Service']}</b> — {r['Suggested_Theme']} (<b>{r['Priority_Label']}</b>)",
                    base,
                )
            )
        story.append(Spacer(1, 10))

        story.append(Paragraph("SIP Candidates (Top 10)", h2))

        header = ["Service", "Tier", "Theme", "Priority", "Why Flagged", "Score"]
        rows = [[Paragraph(f"<b>{h}</b>", base) for h in header]]

        def priority_bg(label: str) -> colors.Color:
            label = (label or "").strip()
            if label.startswith("Immediate"):
                return colors.HexColor("#FFE4E6")
            if label.startswith("Next"):
                return colors.HexColor("#FEF3C7")
            if label.startswith("Monitor"):
                return colors.HexColor("#F3F4F6")
            return WHITE

        row_bg = [LIGHT]  # header background

        for _, r in sip_df.head(10).iterrows():
            row_bg.append(priority_bg(str(r["Priority_Label"])))

            score_val = ""
            try:
                if str(r["SIP_Priority_Score"]).strip() != "":
                    score_val = f"{float(r['SIP_Priority_Score']):.2f}"
            except Exception:
                score_val = str(r["SIP_Priority_Score"])

            rows.append(
                [
                    Paragraph(str(r["Service"]), base),
                    Paragraph(str(r["Service_Tier"]), base),
                    Paragraph(str(r["Suggested_Theme"]), base),
                    Paragraph(str(r["Priority_Label"]), base),
                    Paragraph(str(r["Why_Flagged"]), base),  # wraps
                    Paragraph(score_val, score_right),
                ]
            )

        col_widths = [1.55 * inch, 0.65 * inch, 0.95 * inch, 1.05 * inch, 2.05 * inch, 0.70 * inch]
        sip_tbl = Table(rows, colWidths=col_widths, repeatRows=1)

        ts = [
            ("BOX", (0, 0), (-1, -1), 0.6, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("BACKGROUND", (0, 0), (-1, 0), LIGHT),
        ]
        for i in range(1, len(row_bg)):
            ts.append(("BACKGROUND", (0, i), (-1, i), row_bg[i]))

        sip_tbl.setStyle(TableStyle(ts))
        story.append(sip_tbl)

    # -----------------------------
    # Render with headers/footers
    # -----------------------------
    def on_page(canvas, doc_):
        page = canvas.getPageNumber()
        if page == 1:
            subtitle = "Operational Stability Intelligence Report"
        elif page == 2:
            subtitle = "Operational Stability Profile & Diagnostics"
        elif page == 3:
            subtitle = "Stability Findings & Recommended Actions"
        else:
            subtitle = "Service Improvement Priorities (SIPs)"
        _draw_header(canvas, doc_, "OSIL™ by Xentrixus", subtitle, as_of)

    doc.build(story, onFirstPage=on_page, onLaterPages=on_page)

    out.seek(0)
    return out.getvalue()
