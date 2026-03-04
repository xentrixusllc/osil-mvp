from __future__ import annotations

from io import BytesIO
from typing import Dict, Any, List, Tuple

import numpy as np
import pandas as pd

# PDF
from reportlab.lib.pagesizes import LETTER
from reportlab.pdfgen import canvas
from reportlab.lib.units import inch
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors

# Chart image
import matplotlib.pyplot as plt


# -----------------------------
# Brand (Consulting Style)
# -----------------------------
NAVY = colors.HexColor("#0A192F")
TEAL = colors.HexColor("#64FFDA")
DARK = colors.HexColor("#1F2937")
MID = colors.HexColor("#374151")
LIGHT = colors.HexColor("#F5F7FA")
BORDER = colors.HexColor("#D1D5DB")
WHITE = colors.white


def _wrap_text(
    c: canvas.Canvas,
    text: str,
    x: float,
    y: float,
    max_width: float,
    font_name: str = "Helvetica",
    font_size: int = 11,
    leading: float = 14,
) -> float:
    """Draw wrapped text and return new y."""
    if not text:
        return y
    c.setFont(font_name, font_size)
    words = text.split()
    line = ""
    for w in words:
        test = (line + " " + w).strip()
        if c.stringWidth(test, font_name, font_size) <= max_width:
            line = test
        else:
            c.drawString(x, y, line)
            y -= leading
            line = w
    if line:
        c.drawString(x, y, line)
        y -= leading
    return y


def _draw_header(c: canvas.Canvas, width: float, height: float, title: str, subtitle: str, as_of: str):
    """Top header band with minimal consulting branding."""
    band_h = 0.85 * inch
    c.setFillColor(NAVY)
    c.rect(0, height - band_h, width, band_h, fill=1, stroke=0)

    # Accent line
    c.setFillColor(TEAL)
    c.rect(0, height - band_h, width, 0.06 * inch, fill=1, stroke=0)

    # Title text
    c.setFillColor(WHITE)
    c.setFont("Helvetica-Bold", 18)
    c.drawString(0.75 * inch, height - 0.55 * inch, title)

    c.setFont("Helvetica", 10)
    c.drawString(0.75 * inch, height - 0.72 * inch, subtitle)

    c.setFont("Helvetica", 9)
    c.drawRightString(width - 0.75 * inch, height - 0.72 * inch, f"As-of: {as_of}")


def _posture_chip_colors(posture: str) -> Tuple[colors.Color, colors.Color]:
    """
    Returns (fill, text) colors for posture chip.
    Keep muted/consulting style.
    """
    posture = (posture or "").strip()
    if posture == "High Confidence Operations":
        return colors.HexColor("#D1FAE5"), colors.HexColor("#065F46")  # soft green / dark green
    if posture == "Controlled but Exposed":
        return colors.HexColor("#FEF3C7"), colors.HexColor("#92400E")  # soft amber / dark amber
    if posture == "Fragile Stability":
        return colors.HexColor("#FFE4E6"), colors.HexColor("#9F1239")  # soft rose / dark rose
    return colors.HexColor("#E5E7EB"), colors.HexColor("#111827")      # neutral


def _bvsi_scale_lines(bvsi: float) -> List[str]:
    return [
        "BVSI™ Scale (Business Value Stability Index)",
        "80–100: High Confidence Operations",
        "60–79: Controlled but Exposed",
        "40–59: Fragile Stability",
        "0–39: Reactive Instability",
        f"Current BVSI™: {bvsi:.1f}",
    ]


def _radar_png(domain_scores: Dict[str, float]) -> BytesIO:
    """
    Consulting-style radar (subtle fill, navy line).
    """
    labels = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
    values = [
        float(domain_scores.get("Overall Service Resilience", 0)),
        float(domain_scores.get("Overall Change Governance", 0)),
        float(domain_scores.get("Overall Structural Risk Debt", 0)),
        float(domain_scores.get("Overall Reliability Momentum", 0)),
    ]

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + values[:1]
    angles_loop = angles + angles[:1]

    fig = plt.figure(figsize=(5.6, 4.0), dpi=170)
    ax = plt.subplot(111, polar=True)

    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)

    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)

    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)
    ax.set_ylim(0, 100)

    # Line and fill (consulting style: subtle)
    ax.plot(angles_loop, values_loop, linewidth=2.2, color="#0A192F")
    ax.fill(angles_loop, values_loop, alpha=0.12, color="#64FFDA")

    ax.grid(color="#D1D5DB", linewidth=0.7, alpha=0.9)
    ax.set_title("Operational Stability Profile", fontsize=11, pad=14)

    buf = BytesIO()
    plt.tight_layout()
    fig.savefig(buf, format="png", bbox_inches="tight")
    plt.close(fig)
    buf.seek(0)
    return buf


def _section_box(c: canvas.Canvas, x: float, y_top: float, w: float, h: float, title: str | None = None):
    """Light gray section band."""
    c.setFillColor(LIGHT)
    c.setStrokeColor(BORDER)
    c.rect(x, y_top - h, w, h, fill=1, stroke=1)
    if title:
        c.setFillColor(MID)
        c.setFont("Helvetica-Bold", 11)
        c.drawString(x + 0.18 * inch, y_top - 0.32 * inch, title)


def _safe_df(obj) -> pd.DataFrame:
    if obj is None:
        return pd.DataFrame()
    if isinstance(obj, pd.DataFrame):
        return obj.copy()
    return pd.DataFrame(obj)


def build_osil_pdf_report(results: Dict[str, Any]) -> bytes:
    overall = results["overall"]
    bvsi = float(overall["BVSI"])
    posture = results.get("posture", "")
    as_of = results.get("as_of", "")

    sip_df = _safe_df(results.get("sip_table"))
    if not sip_df.empty:
        sip_df = sip_df.head(10)

    # Radar image (ImageReader)
    radar_reader = ImageReader(_radar_png(overall))

    # Page setup
    out = BytesIO()
    c = canvas.Canvas(out, pagesize=LETTER)
    width, height = LETTER

    margin_x = 0.75 * inch
    max_w = width - 2 * margin_x

    # -------------------------
    # Page 1 — Executive Brief
    # -------------------------
    _draw_header(
        c,
        width,
        height,
        title="OSIL™ by Xentrixus",
        subtitle="Operational Stability Intelligence Report",
        as_of=as_of,
    )

    y = height - (0.85 * inch) - 0.55 * inch

    # BVSI big number
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 32)
    c.drawString(margin_x, y, f"{bvsi:.1f}")

    c.setFont("Helvetica", 11)
    c.setFillColor(MID)
    c.drawString(margin_x, y - 18, "BVSI™ (Business Value Stability Index)")

    # Posture chip
    chip_fill, chip_text = _posture_chip_colors(posture)
    chip_x = margin_x + 3.35 * inch
    chip_y = y + 8
    chip_w = 3.35 * inch
    chip_h = 0.46 * inch

    c.setFillColor(chip_fill)
    c.setStrokeColor(BORDER)
    c.roundRect(chip_x, chip_y - chip_h, chip_w, chip_h, radius=10, fill=1, stroke=1)

    c.setFillColor(chip_text)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(chip_x + 0.18 * inch, chip_y - 0.32 * inch, f"Operating Posture: {posture}")

    # BVSI scale in a light box
    box_top = y - 0.55 * inch
    _section_box(c, margin_x, box_top, max_w, 1.08 * inch, title="BVSI™ Interpretation")
    c.setFillColor(DARK)
    c.setFont("Helvetica", 10)
    ty = box_top - 0.55 * inch
    for line in _bvsi_scale_lines(bvsi):
        c.drawString(margin_x + 0.18 * inch, ty, line)
        ty -= 12

    # Executive narrative (stronger, consulting tone)
    y2 = box_top - 1.28 * inch
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_x, y2, "Executive Summary")
    y2 -= 16

    c.setFillColor(DARK)
    summary = (
        f"Your organization is operating in a **{posture}** stability posture based on a BVSI™ score of {bvsi:.1f}. "
        "This indicates that operational control exists, but recurring instability patterns still create exposure—especially "
        "across higher-impact services. Targeted Service Improvement Programs (SIPs) in the next 30–60 days can materially "
        "increase executive confidence and preserve customer trust during disruption events."
    ).replace("**", "")  # plain canvas text
    y2 = _wrap_text(c, summary, margin_x, y2, max_w, font_name="Helvetica", font_size=11, leading=14)

    # Radar chart + key observation
    # place chart
    img_w = 6.9 * inch
    img_h = 3.35 * inch
    img_x = margin_x
    img_y = 1.85 * inch
    c.drawImage(radar_reader, img_x, img_y, width=img_w, height=img_h, mask="auto")

    # Key observation line
    c.setFillColor(MID)
    c.setFont("Helvetica-Oblique", 9)
    c.drawString(
        margin_x,
        img_y - 12,
        "How to read: a balanced shape suggests aligned governance; a collapsed axis indicates a concentrated stability gap."
    )

    c.showPage()

    # -------------------------
    # Page 2 — Diagnostics
    # -------------------------
    _draw_header(
        c,
        width,
        height,
        title="OSIL™ by Xentrixus",
        subtitle="Operational Stability Diagnostics",
        as_of=as_of,
    )

    y = height - (0.85 * inch) - 0.55 * inch

    # Diagnostics intro
    c.setFillColor(DARK)
    c.setFont("Helvetica", 11)
    intro = (
        "Domain scores below reflect tier-weighted operational stability. Lower-scoring domains represent stability constraints "
        "that most directly affect executive confidence and customer experience under disruption."
    )
    y = _wrap_text(c, intro, margin_x, y, max_w, font_name="Helvetica", font_size=11, leading=14)
    y -= 6

    # Domain table with light banding + simple color cue
    domains = [
        ("Service Resilience", float(overall.get("Overall Service Resilience", 0))),
        ("Change Governance", float(overall.get("Overall Change Governance", 0))),
        ("Structural Risk Debt™", float(overall.get("Overall Structural Risk Debt", 0))),
        ("Reliability Momentum", float(overall.get("Overall Reliability Momentum", 0))),
    ]

    # Table header
    table_x = margin_x
    table_w = max_w
    row_h = 0.36 * inch

    c.setFillColor(LIGHT)
    c.setStrokeColor(BORDER)
    c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=1)

    c.setFillColor(MID)
    c.setFont("Helvetica-Bold", 10)
    c.drawString(table_x + 0.18 * inch, y - 0.25 * inch, "Domain")
    c.drawRightString(table_x + table_w - 0.18 * inch, y - 0.25 * inch, "Score (0–100)")
    y -= row_h

    def score_color(s: float) -> colors.Color:
        if s >= 80:
            return colors.HexColor("#065F46")  # green
        if s >= 60:
            return colors.HexColor("#92400E")  # amber
        if s >= 40:
            return colors.HexColor("#9F1239")  # rose
        return colors.HexColor("#111827")      # near-black

    # rows
    c.setFont("Helvetica", 10)
    for i, (name, score) in enumerate(domains):
        fill = colors.white if i % 2 == 0 else colors.HexColor("#FBFCFE")
        c.setFillColor(fill)
        c.setStrokeColor(BORDER)
        c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=1)

        c.setFillColor(DARK)
        c.drawString(table_x + 0.18 * inch, y - 0.25 * inch, name)

        c.setFillColor(score_color(score))
        c.setFont("Helvetica-Bold", 10)
        c.drawRightString(table_x + table_w - 0.18 * inch, y - 0.25 * inch, f"{score:.1f}")
        c.setFont("Helvetica", 10)

        y -= row_h

    y -= 0.25 * inch

    # Key findings (consulting bullets)
    c.setFillColor(DARK)
    c.setFont("Helvetica-Bold", 12)
    c.drawString(margin_x, y, "Key Stability Findings")
    y -= 16

    # Lightweight, always-safe findings (no over-claim)
    findings = [
        "Tier-1 services should be treated as the primary stability surface area due to business impact concentration.",
        "Structural Risk Debt™ indicates recurring instability patterns that are not being fully eliminated through prevention work.",
        "Strong change governance reduces release-driven volatility, enabling targeted reliability improvements to drive the next BVSI™ lift.",
    ]
    c.setFont("Helvetica", 11)
    for f in findings:
        y = _wrap_text(c, f"• {f}", margin_x, y, max_w, font_name="Helvetica", font_size=11, leading=14)
    y -= 6

    # Recommended actions
    c.setFont("Helvetica-Bold", 12)
    c.setFillColor(DARK)
    c.drawString(margin_x, y, "Recommended Next Actions (Next 30 Days)")
    y -= 16

    actions = [
        "Launch SIPs against the top recurring instability clusters across Tier-1 services (Service + Theme).",
        "Reduce reopen churn by strengthening closure verification and linking repeat themes to prevention work with owners and due dates.",
        "Improve recovery readiness for Tier-1 services through runbooks, alert quality tuning, and escalation path clarity.",
    ]
    c.setFont("Helvetica", 11)
    for i, a in enumerate(actions, start=1):
        y = _wrap_text(c, f"{i}. {a}", margin_x, y, max_w, font_name="Helvetica", font_size=11, leading=14)

    c.showPage()

    # -------------------------
    # Page 3 — SIP Priorities
    # -------------------------
    _draw_header(
        c,
        width,
        height,
        title="OSIL™ by Xentrixus",
        subtitle="Service Improvement Priorities (SIPs)",
        as_of=as_of,
    )

    y = height - (0.85 * inch) - 0.55 * inch
    c.setFillColor(DARK)
    c.setFont("Helvetica", 11)
    sip_intro = (
        "SIPs below are prioritized by Service + Theme and tier exposure. These represent the most actionable stability initiatives "
        "to improve BVSI™ and reduce avoidable disruption."
    )
    y = _wrap_text(c, sip_intro, margin_x, y, max_w, font_name="Helvetica", font_size=11, leading=14)
    y -= 10

    if sip_df.empty:
        c.setFillColor(MID)
        c.setFont("Helvetica-Oblique", 11)
        c.drawString(margin_x, y, "No SIP candidates available from the current dataset.")
    else:
        # Ensure expected columns exist
        for col in ["Service", "Service_Tier", "Suggested_Theme", "Priority_Label", "Why_Flagged", "SIP_Priority_Score"]:
            if col not in sip_df.columns:
                sip_df[col] = ""

        sip_df = sip_df[["Service", "Service_Tier", "Suggested_Theme", "Priority_Label", "Why_Flagged", "SIP_Priority_Score"]].copy()

        # Highlight top 3
        c.setFillColor(DARK)
        c.setFont("Helvetica-Bold", 12)
        c.drawString(margin_x, y, "Top 3 Initiatives to Brief Leadership")
        y -= 14

        top3 = sip_df.head(3)
        c.setFont("Helvetica", 11)
        for idx, r in top3.iterrows():
            line = f"• {str(r['Service'])} — {str(r['Suggested_Theme'])} ({str(r['Priority_Label'])})"
            y = _wrap_text(c, line, margin_x, y, max_w, font_name="Helvetica", font_size=11, leading=14)

        y -= 8

        # SIP table
        c.setFont("Helvetica-Bold", 10)
        c.setFillColor(MID)
        c.drawString(margin_x, y, "SIP Candidates (Top 10)")
        y -= 12

        # Table layout
        table_x = margin_x
        table_w = max_w
        row_h = 0.34 * inch

        col_defs = [
            ("Service", 2.10 * inch),
            ("Tier", 0.65 * inch),
            ("Theme", 1.15 * inch),
            ("Priority", 1.05 * inch),
            ("Why Flagged", 2.45 * inch),
        ]
        # Score is right-aligned in remaining space
        used_w = sum(w for _, w in col_defs)
        score_w = table_w - used_w

        # Header row
        c.setFillColor(LIGHT)
        c.setStrokeColor(BORDER)
        c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=1)

        cx = table_x
        c.setFillColor(MID)
        c.setFont("Helvetica-Bold", 9)
        for name, w in col_defs:
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, name)
            cx += w
        c.drawRightString(table_x + table_w - 0.10 * inch, y - 0.23 * inch, "Score")
        y -= row_h

        def priority_fill(label: str) -> colors.Color:
            # Muted consulting highlights
            label = (label or "").strip()
            if label.startswith("Immediate"):
                return colors.HexColor("#FFE4E6")  # soft rose
            if label.startswith("Next"):
                return colors.HexColor("#FEF3C7")  # soft amber
            if label.startswith("Monitor"):
                return colors.HexColor("#F3F4F6")  # soft gray
            return colors.white

        c.setFont("Helvetica", 9)
        for i, (_, r) in enumerate(sip_df.iterrows()):
            if i >= 10:
                break

            fill = priority_fill(str(r["Priority_Label"]))
            c.setFillColor(fill)
            c.setStrokeColor(BORDER)
            c.rect(table_x, y - row_h, table_w, row_h, fill=1, stroke=1)

            service = str(r["Service"])[:28]
            tier = str(r["Service_Tier"])[:7]
            theme = str(r["Suggested_Theme"])[:16]
            prio = str(r["Priority_Label"])[:14]
            why = str(r["Why_Flagged"])[:36]
            score = r["SIP_Priority_Score"]

            cx = table_x
            c.setFillColor(DARK)
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, service); cx += col_defs[0][1]
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, tier);    cx += col_defs[1][1]
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, theme);   cx += col_defs[2][1]
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, prio);    cx += col_defs[3][1]
            c.drawString(cx + 0.10 * inch, y - 0.23 * inch, why)

            try:
                c.setFillColor(MID)
                c.drawRightString(table_x + table_w - 0.10 * inch, y - 0.23 * inch, f"{float(score):.2f}")
            except Exception:
                c.drawRightString(table_x + table_w - 0.10 * inch, y - 0.23 * inch, str(score))

            y -= row_h
            if y < 1.2 * inch:
                break

    # Footer
    c.setFillColor(MID)
    c.setFont("Helvetica", 8.5)
    c.drawString(margin_x, 0.65 * inch, "Operational Stability Intelligence (OSIL™)  •  BVSI™  •  Structural Risk Debt™")
    c.drawRightString(width - margin_x, 0.65 * inch, "© Xentrixus • Confidential")

    c.showPage()
    c.save()

    out.seek(0)
    return out.getvalue()
