
# ===============================
# OSIL Executive PDF Generator
# report_generator.py
# ===============================

import io
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

from reportlab.lib.pagesizes import LETTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, Image
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch


# -------------------------------
# Radar chart for PDF
# -------------------------------
def build_radar_chart(domain_scores):

    labels = list(domain_scores.keys())
    values = list(domain_scores.values())

    if len(labels) == 0:
        return None

    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values += values[:1]
    angles += angles[:1]

    fig = plt.figure(figsize=(4,4))
    ax = plt.subplot(111, polar=True)

    ax.plot(angles, values)
    ax.fill(angles, values, alpha=0.1)

    ax.set_xticks(angles[:-1])
    ax.set_xticklabels(labels)

    ax.set_ylim(0,100)

    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png")
    plt.close()

    img.seek(0)
    return img


# -------------------------------
# Heatmap chart
# -------------------------------
def build_heatmap(service_df):

    if service_df is None or service_df.empty:
        return None

    cols = [
        "Recurrence_Risk",
        "MTTR_Drag_Risk",
        "Reopen_Churn_Risk",
        "Change_Collision_Risk"
    ]

    missing = [c for c in cols if c not in service_df.columns]
    if missing:
        return None

    df = service_df.copy().head(10)

    labels = df["Service"]

    matrix = df[cols].values

    fig = plt.figure(figsize=(6,3))
    ax = plt.gca()

    im = ax.imshow(matrix, aspect="auto")

    ax.set_xticks(range(len(cols)))
    ax.set_xticklabels(cols, rotation=45)

    ax.set_yticks(range(len(labels)))
    ax.set_yticklabels(labels)

    plt.colorbar(im)

    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png")
    plt.close()

    img.seek(0)

    return img


# -------------------------------
# Build PDF
# -------------------------------
def build_osil_pdf_report(results):

    buffer = io.BytesIO()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=LETTER
    )

    styles = getSampleStyleSheet()

    story = []

    story.append(Paragraph("Operational Stability Intelligence (OSIL)", styles["Title"]))
    story.append(Spacer(1,10))

    story.append(Paragraph("Executive Summary", styles["Heading2"]))

    story.append(Paragraph(results.get("exec_text",""), styles["BodyText"]))

    story.append(Spacer(1,20))

    # KPI TABLE

    kpi_data = [
        ["BVSI", round(results.get("bvsi",0),1)],
        ["Operating Posture", results.get("posture","")],
        ["Data Readiness", str(results.get("readiness_score",""))+"%"]
    ]

    story.append(Table(kpi_data))

    story.append(Spacer(1,20))

    # Radar

    story.append(Paragraph("Operational Stability Radar", styles["Heading2"]))

    radar = build_radar_chart(results.get("domain_scores",{}))

    if radar:
        story.append(Image(radar, 4*inch, 4*inch))

    story.append(Spacer(1,20))

    # Heatmap

    story.append(Paragraph("Service Stability Heatmap", styles["Heading2"]))

    heatmap = build_heatmap(results.get("top10"))

    if heatmap:
        story.append(Image(heatmap, 6*inch, 3*inch))

    story.append(Spacer(1,20))

    # Top services

    story.append(Paragraph("Top Service Risks", styles["Heading2"]))

    if isinstance(results.get("top10"), pd.DataFrame):

        df = results["top10"].copy()

        table = [df.columns.tolist()] + df.values.tolist()

        story.append(Table(table))

    story.append(Spacer(1,20))

    story.append(Paragraph("Service Improvement Candidates", styles["Heading2"]))

    if isinstance(results.get("sip_view"), pd.DataFrame):

        df = results["sip_view"].copy()

        table = [df.columns.tolist()] + df.values.tolist()

        story.append(Table(table))

    doc.build(story)

    buffer.seek(0)

    return buffer.getvalue()
