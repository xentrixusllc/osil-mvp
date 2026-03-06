import io
import pandas as pd
import matplotlib.pyplot as plt
import numpy as np

from reportlab.lib.pagesizes import LETTER
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
from reportlab.lib import colors
from reportlab.lib.units import inch


# =========================================================
# Styles
# =========================================================

styles = getSampleStyleSheet()

TITLE = ParagraphStyle(
    "Title",
    parent=styles["Heading1"],
    fontSize=22,
    spaceAfter=12
)

HEADER = ParagraphStyle(
    "Header",
    parent=styles["Heading2"],
    fontSize=16,
    spaceAfter=8
)

BODY = ParagraphStyle(
    "Body",
    parent=styles["BodyText"],
    fontSize=10,
    leading=14
)

SMALL = ParagraphStyle(
    "Small",
    parent=styles["BodyText"],
    fontSize=8,
    textColor=colors.grey
)


# =========================================================
# Radar Chart
# =========================================================

def build_radar(domain_scores):

    labels = list(domain_scores.keys())
    values = list(domain_scores.values())

    if len(values) == 0:
        return None

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False)
    values = np.concatenate((values, [values[0]]))
    angles = np.concatenate((angles, [angles[0]]))

    fig = plt.figure(figsize=(4.5,4.5), dpi=200)

    ax = plt.subplot(111, polar=True)
    ax.plot(angles, values, linewidth=2)
    ax.fill(angles, values, alpha=0.15)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels, fontsize=8)

    ax.set_ylim(0,100)

    buf = io.BytesIO()
    plt.tight_layout()
    plt.savefig(buf, format="png")
    plt.close()

    buf.seek(0)
    return buf


# =========================================================
# Heatmap
# =========================================================

def build_heatmap(service_risk):

    if service_risk is None or service_risk.empty:
        return None

    cols = [
        "Recurrence_Risk",
        "MTTR_Drag_Risk",
        "Reopen_Churn_Risk",
        "Change_Collision_Risk"
    ]

    df = service_risk.copy()

    df = df[["Service","Service_Tier"] + cols]

    df.index = df["Service"] + " (" + df["Service_Tier"] + ")"

    hm = df[cols]

    fig = plt.figure(figsize=(7,4), dpi=200)

    ax = plt.gca()

    im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100)

    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(["Recurrence","MTTR Drag","Reopen","Change"])

    ax.set_yticks(range(len(hm.index)))
    ax.set_yticklabels(hm.index)

    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            val = int(hm.values[i,j])
            ax.text(j,i,val,ha="center",va="center",fontsize=8)

    plt.colorbar(im)

    buf = io.BytesIO()

    plt.tight_layout()
    plt.savefig(buf, format="png")
    plt.close()

    buf.seek(0)

    return buf


# =========================================================
# Domain Score Table
# =========================================================

def domain_table(domain_scores):

    data = [["Domain","Score"]]

    for k,v in domain_scores.items():
        data.append([k, round(v,1)])

    table = Table(data)

    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey)
    ]))

    return table


# =========================================================
# SIP Table
# =========================================================

def sip_table(df):

    if df is None or df.empty:
        return Paragraph("No SIP candidates detected.",BODY)

    data = [list(df.columns)]

    for row in df.values:
        data.append([str(x) for x in row])

    table = Table(data)

    table.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.grey),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey)
    ]))

    return table


# =========================================================
# Main Report Builder
# =========================================================

def build_osil_pdf_report(payload):

    bvsi = payload["bvsi"]
    posture = payload["posture"]
    as_of = payload["as_of"]

    interpretation = payload["executive_interpretation"]

    domain_scores = payload["domain_scores"]

    sip = payload["sip_candidates"]

    service_risk = payload["service_risk_top10"]

    tenant = payload["tenant_name"]

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=LETTER,
        rightMargin=50,
        leftMargin=50,
        topMargin=40,
        bottomMargin=40
    )

    story = []

    # =====================================================
    # PAGE 1
    # =====================================================

    story.append(Paragraph("Operational Stability Intelligence (OSIL™)", TITLE))
    story.append(Paragraph(f"Executive Report – {tenant}", BODY))
    story.append(Paragraph(f"As of {as_of}", SMALL))
    story.append(Spacer(1,20))

    story.append(Paragraph("Operational Stability Snapshot", HEADER))

    story.append(Paragraph(
        f"<b>BVSI™ Score:</b> {bvsi:.1f}<br/>"
        f"<b>Operating Posture:</b> {posture}",
        BODY
    ))

    story.append(Spacer(1,12))

    story.append(Paragraph("Executive Interpretation", HEADER))

    story.append(Paragraph(interpretation, BODY))

    story.append(PageBreak())

    # =====================================================
    # PAGE 2
    # =====================================================

    story.append(Paragraph("Operational Stability Profile", HEADER))

    radar = build_radar(domain_scores)

    if radar:
        story.append(Image(radar, width=4.5*inch, height=4.5*inch))

    story.append(Spacer(1,10))

    story.append(domain_table(domain_scores))

    story.append(PageBreak())

    # =====================================================
    # PAGE 3
    # =====================================================

    story.append(Paragraph("Service Improvement Programs (SIPs)", HEADER))

    story.append(Paragraph(
        "These represent the highest impact stability improvements for the next 30–60 days.",
        BODY
    ))

    story.append(Spacer(1,10))

    story.append(sip_table(sip))

    story.append(PageBreak())

    # =====================================================
    # PAGE 4
    # =====================================================

    story.append(Paragraph("Service Stability Heatmap", HEADER))

    heat = build_heatmap(service_risk)

    if heat:
        story.append(Image(heat, width=6.5*inch, height=4*inch))

    doc.build(story)

    pdf = buffer.getvalue()

    buffer.close()

    return pdf
