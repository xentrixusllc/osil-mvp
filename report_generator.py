from __future__ import annotations

from io import BytesIO
from typing import Dict, Any, List

import numpy as np
import pandas as pd

# PDF
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader  # <-- FIX

# Chart image
import matplotlib.pyplot as plt


def _wrap_text(c: canvas.Canvas, text: str, x: float, y: float, max_width: float, leading: float = 14):
    """Simple text wrapping for ReportLab canvas. Returns the new y after drawing."""
    words = (text or "").split()
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, "Helvetica", 11) <= max_width:
            line = test
        else:
            c.drawString(x, y, line)
            y -= leading
            line = w
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def _radar_png(domain_scores: Dict[str, float]) -> BytesIO:
    """
    Creates a static radar chart PNG in memory using matplotlib.
    Returns BytesIO buffer.
    """
    labels = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
    values = [
        float(domain_scores.get("Overall Service Resilience", 0)),
        float(domain_scores.get("Overall Change Governance", 0)),
        float(domain_scores.get("Overall Structural Risk Debt", 0)),
        float(domain_scores.get("Overall Reliability Momentum", 0)),
    ]

    # Close the loop for radar
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + values[:1]
    angles_loop = angles + angles[:1]

    fig = plt.figure(figsize=(5.2, 3.8), dpi=160)
    ax = plt.subplot(111, polar=True)

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)

    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)
    ax.set_ylim(0, 100)

    ax.plot(angles_loop, values_loop, linewidth=2)
    ax.fill(angles_loop, values_loop, alpha=0.12)

    ax.set_title("Operational Stability Profile", fontsize=11, pad=14)

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _bvsi_scale_block(bvsi: float) -> List[str]:
    return [
        "BVSI Scale (Business Value Stability Index)",
        "80–100: High Confidence Operations",
        "60–79: Controlled but Exposed",
        "40–59: Fragile Stability",
        "0–39: Reactive Instability",
        f"Current BVSI: {bvsi:.1f}",
    ]


def build_osil_pdf_report(results: Dict[str, Any]) -> bytes:
    """
    Build a consulting-style executive PDF report.
    Returns PDF bytes.
    Expects results from run_osil():
      - as_of, posture, overall, sip_table
      - (optional) notes
    """
    overall = results["overall"]
    bvsi = float(overall["BVSI"])
    posture = results.get("posture", "")

    # SIP table
    sip_table = results.get("sip_table")
    if isinstance(sip_table, pd.DataFrame):
        sip_df = sip_table.copy()
    else:
        sip_df = pd.DataFrame(sip_table) if sip_table is not None else pd.DataFrame()

    if not sip_df.empty:
        sip_df = sip_df.head(10)

    # Radar image -> ImageReader (FIX)
    radar_buf = _radar_png(overall)
    radar_reader = ImageReader(radar_buf)

    # Create PDF
    out = BytesIO()
    c = canvas.Canvas(out, pagesize=LETTER)
    width, height = LETTER

    margin_x = 0.75 * inch
    y = height - 0.85 * inch
    max_w = width - 2 * margin_x

    # -------------------------
    # Page 1: Executive Summary
    # -------------------------
    c.setFont("Helvetica-Bold", 20)
    c.drawString(margin_x, y, "OSIL by Xentrixus")
    y -= 22
    c.setFont("Helvetica", 12)
    c.drawString(margin_x, y, "Operational Stability Intelligence Report")
    y -= 18

    c.setFont("Helvetica", 10)
    c.drawString(margin_x, y, f"As-of Date: {results.get('as_of','')}")
    y -= 18

    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_x, y, f"BVSI: {bvsi:.1f}")
    y -= 14
    c.drawString(margin_x, y, f"Operating Posture: {posture}")
    y -= 18

    c.setFont("Helvetica", 10)
    for line in _bvsi_scale_block(bvsi):
        c.drawString(margin_x, y, line)
        y -= 12

    y -= 6
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_x, y, "Executive Interpretation")
    y -= 16

    c.setFont("Helvetica", 11)
    interpretation = (
        "This report summarizes operational stability across resilience, governance, structural risk exposure, "
        "and reliability momentum. The BVSI posture indicates how consistently services deliver outcomes "
        "without disruption to customer experience and internal execution confidence."
    )
    y = _wrap_text(c, interpretation, margin_x, y, max_w, leading=14)

    # Draw radar (FIX: use ImageReader)
    y -= 8
    img_x = margin_x
    img_y = max(1.3 * inch, y - 3.3 * inch)
    c.drawImage(radar_reader, img_x, img_y, width=6.6 * inch, height=3.2 * inch, mask="auto")

    c.setFont("Helvetica-Oblique", 9)
    c.drawString(
        margin_x,
        img_y - 12,
        "How to read: a balanced shape suggests aligned governance; a collapsed axis indicates a concentrated stability gap."
    )
    c.showPage()

    # -------------------------
    # Page 2: Diagnostics
    # -------------------------
    y = height - 0.85 * inch
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, "Operational Stability Diagnostics")
    y -= 22

    c.setFont("Helvetica", 11)
    diag_text = (
        "Domain scores reflect tier-weighted operational stability. Lower-scoring domains represent stability constraints "
        "that most directly impact executive confidence and customer trust under disruption."
    )
    y = _wrap_text(c, diag_text, margin_x, y, max_w, leading=14)
    y -= 10

    domains = [
        ("Service Resilience", float(overall.get("Overall Service Resilience", 0))),
        ("Change Governance", float(overall.get("Overall Change Governance", 0))),
        ("Structural Risk Debt™", float(overall.get("Overall Structural Risk Debt", 0))),
        ("Reliability Momentum", float(overall.get("Overall Reliability Momentum", 0))),
    ]

    c.setFont("Helvetica-Bold", 11)
    c.drawString(margin_x, y, "Domain Scores")
    y -= 14

    c.setFont("Helvetica-Bold", 10)
    c.drawString(margin_x, y, "Domain")
    c.drawRightString(margin_x + max_w, y, "Score (0–100)")
    y -= 10
    c.setFont("Helvetica", 10)
    c.line(margin_x, y, margin_x + max_w, y)
    y -= 12

    for name, score in domains:
        c.drawString(margin_x, y, name)
        c.drawRightString(margin_x + max_w, y, f"{score:.1f}")
        y -= 14

    y -= 8
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_x, y, "Recommended Next Actions (Next 30 Days)")
    y -= 16
    c.setFont("Helvetica", 11)

    actions = [
        "Launch SIPs against recurring instability clusters in Tier 1 services (reduce repeat incidents and reopen churn).",
        "Strengthen closure-to-prevention: convert repeat themes into corrective actions with owners, due dates, and verification steps.",
        "Tighten change discipline for high-impact services to reduce instability introduced by releases and configuration shifts.",
    ]
    for i, a in enumerate(actions, start=1):
        y = _wrap_text(c, f"{i}. {a}", margin_x, y, max_w, leading=14)

    c.showPage()

    # -------------------------
    # Page 3: SIP Priorities
    # -------------------------
    y = height - 0.85 * inch
    c.setFont("Helvetica-Bold", 16)
    c.drawString(margin_x, y, "Service Improvement Priorities (SIPs)")
    y -= 20

    c.setFont("Helvetica", 11)
    intro = (
        "Top SIP candidates are prioritized by Service + Theme and tier exposure. "
        "These represent the most actionable stability improvements to increase BVSI and reduce avoidable disruption."
    )
    y = _wrap_text(c, intro, margin_x, y, max_w, leading=14)
    y -= 10

    if sip_df.empty:
        c.setFont("Helvetica-Oblique", 11)
        c.drawString(margin_x, y, "No SIP candidates available from the current dataset.")
    else:
        # Safe columns (Phase 6 compatible)
        cols = ["Service", "Service_Tier", "Suggested_Theme", "Priority_Label", "Why_Flagged", "SIP_Priority_Score"]
        cols = [col for col in cols if col in sip_df.columns]
        table = sip_df[cols].copy()

        c.setFont("Helvetica-Bold", 9)
        x_positions = [
            margin_x,               # Service
            margin_x + 1.9 * inch,  # Tier
            margin_x + 2.8 * inch,  # Theme
            margin_x + 3.9 * inch,  # Priority
            margin_x + 5.1 * inch,  # Why
        ]

        c.drawString(x_positions[0], y, "Service")
        c.drawString(x_positions[1], y, "Tier")
        c.drawString(x_positions[2], y, "Theme")
        c.drawString(x_positions[3], y, "Priority")
        c.drawString(x_positions[4], y, "Why Flagged")
        c.drawRightString(margin_x + max_w, y, "Score")
        y -= 10

        c.setFont("Helvetica", 9)
        c.line(margin_x, y, margin_x + max_w, y)
        y -= 12

        for _, r in table.iterrows():
            service = str(r.get("Service", ""))[:26]
            tier = str(r.get("Service_Tier", ""))[:7]
            theme = str(r.get("Suggested_Theme", ""))[:16]
            priority = str(r.get("Priority_Label", ""))[:12]
            why = str(r.get("Why_Flagged", ""))[:28]
            score = r.get("SIP_Priority_Score", "")

            c.drawString(x_positions[0], y, service)
            c.drawString(x_positions[1], y, tier)
            c.drawString(x_positions[2], y, theme)
            c.drawString(x_positions[3], y, priority)
            c.drawString(x_positions[4], y, why)
            try:
                c.drawRightString(margin_x + max_w, y, f"{float(score):.2f}")
            except Exception:
                c.drawRightString(margin_x + max_w, y, str(score))

            y -= 12
            if y < 1.0 * inch:
                c.showPage()
                y = height - 0.85 * inch
                c.setFont("Helvetica-Bold", 12)
                c.drawString(margin_x, y, "Service Improvement Priorities (continued)")
                y -= 18
                c.setFont("Helvetica", 9)

    c.showPage()
    c.save()

    out.seek(0)
    return out.getvalue()
