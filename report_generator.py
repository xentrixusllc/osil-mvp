"""PDF Generator for Xentrixus OSIL"""

import io
import textwrap
import numpy as np
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, Image, PageBreak
)

def radar_chart(domain_scores: dict):
    labels = list(domain_scores.keys())
    values = [float(domain_scores[k]) for k in labels]
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    values_loop = values + [values[0]]
    angles_loop = angles + [angles[0]]

    fig = plt.figure(figsize=(4, 4), dpi=150)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.plot(angles_loop, values_loop, linewidth=2, color="#2563EB")
    ax.fill(angles_loop, values_loop, alpha=0.15, color="#2563EB")
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8, fontweight="bold")
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
    
    ax.spines['polar'].set_visible(False)
    plt.tight_layout()
    
    img_data = io.BytesIO()
    plt.savefig(img_data, format='png', bbox_inches='tight')
    plt.close(fig)
    img_data.seek(0)
    return img_data

def heatmap_chart(hm: pd.DataFrame):
    fig = plt.figure(figsize=(8, 4), dpi=150)
    ax = plt.gca()
    
    im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap="RdYlGn_r")
    
    ax.set_xticks(range(len(hm.columns)))
    ax.set_xticklabels(list(hm.columns), fontsize=8, fontweight='bold', rotation=15, ha="right")
    
    ax.set_yticks(range(len(hm.index)))
    ax.set_yticklabels(list(hm.index), fontsize=8, fontweight='bold')
    
    for i in range(hm.shape[0]):
        for j in range(hm.shape[1]):
            val = int(round(float(hm.iat[i, j]), 0))
            text_color = "white" if (val > 80 or val < 20) else "black"
            ax.text(j, i, f"{val}", ha="center", va="center", fontsize=8, fontweight='bold', color=text_color)

    plt.tight_layout()
    
    img_data = io.BytesIO()
    plt.savefig(img_data, format='png', bbox_inches='tight')
    plt.close(fig)
    img_data.seek(0)
    return img_data

def plot_pareto(df: pd.DataFrame):
    fig, ax1 = plt.subplots(figsize=(7, 4), dpi=150)
    
    labels = [textwrap.fill(str(x)[:22], width=18) for x in df["Theme"]]
    x_pos = np.arange(len(df))
    
    ax1.bar(x_pos, df["Frequency"], color="#3B82F6", width=0.5)
    ax1.set_ylabel("Frequency", color="black", fontweight="bold", fontsize=8)
    ax1.tick_params(axis="y", labelsize=8)
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(labels, rotation=35, ha="right", fontsize=7)

    ax2 = ax1.twinx()
    ax2.plot(x_pos, df["Cumulative_Pct"], color="#DC2626", marker="o", linewidth=2)
    ax2.set_ylabel("Cumulative %", color="#DC2626", fontweight="bold", fontsize=8)
    ax2.set_ylim(0, 110)
    ax2.tick_params(axis="y", labelsize=8)
    
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.tight_layout()
    
    img_data = io.BytesIO()
    plt.savefig(img_data, format='png', bbox_inches='tight')
    plt.close(fig)
    img_data.seek(0)
    return img_data

def plot_impact_matrix(df: pd.DataFrame):
    if df.empty or "Active_Disruption_P1_P2" not in df.columns:
        return None
        
    merged = df.sort_values(by=["Active_Disruption_P1_P2", "Total_Service_Risk"], ascending=[False, False]).head(5).copy()
    merged["Active_Disruption_P1_P2"] = pd.to_numeric(merged["Active_Disruption_P1_P2"], errors="coerce").fillna(0)
    merged["Recurrence_Risk"] = pd.to_numeric(merged["Recurrence_Risk"], errors="coerce").fillna(0)
        
    fig, ax1 = plt.subplots(figsize=(7, 4), dpi=150)
    x_pos = np.arange(len(merged))
    labels = [textwrap.fill(str(x)[:15], width=10) for x in merged["Service"]]
    
    ax1.bar(x_pos, merged["Active_Disruption_P1_P2"], color="#DC2626", width=0.4, alpha=0.9)
    ax1.set_ylabel("Active Disruption (P1/P2)", color="#DC2626", fontweight="bold", fontsize=8)
    ax1.tick_params(axis="y", labelcolor="#DC2626", labelsize=8)
    ax1.set_xticks(x_pos)
    ax1.set_xticklabels(labels, fontsize=8, fontweight="bold")

    max_val = float(merged["Active_Disruption_P1_P2"].max())
    ax1.set_ylim(0, max(5, max_val * 1.3))

    ax2 = ax1.twinx()
    ax2.plot(x_pos, merged["Recurrence_Risk"], color="#0F172A", marker="o", linewidth=2)
    ax2.set_ylabel("Recurrence Risk Score", color="#0F172A", fontweight="bold", fontsize=8)
    ax2.tick_params(axis="y", labelcolor="#0F172A", labelsize=8)
    ax2.set_ylim(0, 115)
    
    ax1.spines['top'].set_visible(False)
    ax2.spines['top'].set_visible(False)
    
    plt.tight_layout()
    
    img_data = io.BytesIO()
    plt.savefig(img_data, format='png', bbox_inches='tight')
    plt.close(fig)
    img_data.seek(0)
    return img_data

def df_to_table(df: pd.DataFrame, col_widths=None) -> Table:
    if df.empty:
        return Table([["No Data Available"]])
        
    data = [df.columns.tolist()] + df.astype(str).values.tolist()
    
    t = Table(data, colWidths=col_widths)
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
        ('BACKGROUND', (0, 1), (-1, -1), colors.HexColor("#F8FAFC")),
        ('TEXTCOLOR', (0, 1), (-1, -1), colors.HexColor("#334155")),
        ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
        ('FONTSIZE', (0, 1), (-1, -1), 8),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ]))
    return t

def build_osil_pdf_report(payload: dict) -> bytes:
    bvsi = payload.get("bvsi", 0.0)
    posture = payload.get("posture", "Unknown")
    as_of = payload.get("as_of", "")
    exec_text = payload.get("executive_interpretation", "")
    domain_scores = payload.get("domain_scores", {})
    top10 = payload.get("service_risk_top10", pd.DataFrame())
    pareto_df = payload.get("rca_pareto_df", pd.DataFrame())
    trust_gap_narrative = payload.get("trust_gap_narrative", "")
    trust_gap_df = payload.get("trust_gap_df", pd.DataFrame())
    automation_df = payload.get("automation_df", pd.DataFrame())
    readiness_score = payload.get("data_readiness_score", 0.0)
    tenant_name = payload.get("tenant_name", "Organization")

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        buffer, 
        pagesize=letter,
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
    )

    styles = getSampleStyleSheet()
    
    styles.add(ParagraphStyle(
        name='Title_Custom',
        fontName='Helvetica-Bold',
        fontSize=24,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=6
    ))
    
    styles.add(ParagraphStyle(
        name='Subtitle_Custom',
        fontName='Helvetica',
        fontSize=12,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=24
    ))

    styles.add(ParagraphStyle(
        name='MetricBox',
        fontName='Helvetica-Bold',
        fontSize=14,
        textColor=colors.HexColor("#2563EB"),
        spaceAfter=12
    ))

    story = []

    # --- TITLE PAGE / HEADER ---
    story.append(Paragraph(f"Xentrixus OSIL™ • {tenant_name}", styles['Title_Custom']))
    story.append(Paragraph(f"Business Value Stability & Structural Risk Analysis • As of {as_of}", styles['Subtitle_Custom']))

    # --- CORE METRICS ---
    metric_text = (
        f"<b>Global Stability (BVSI™):</b> {bvsi:.1f} / 100 <br/>"
        f"<b>Operating Posture:</b> {posture} <br/>"
        f"<b>Data Readiness Score:</b> {readiness_score:.1f}%"
    )
    story.append(Paragraph(metric_text, styles['MetricBox']))
    story.append(Spacer(1, 12))

    # --- EXECUTIVE INTERPRETATION ---
    story.append(Paragraph("Executive Interpretation", styles['Heading2']))
    story.append(Paragraph(exec_text.replace('\n', '<br/>'), styles['Normal']))
    story.append(Spacer(1, 20))
        
    story.append(PageBreak())

    # --- TRUST GAP MATRIX ---
    story.append(Paragraph("Business Trust and Disruption Matrix", styles['Heading1']))
    story.append(Paragraph(trust_gap_narrative, styles['Normal']))
    story.append(Spacer(1, 12))
    
    if not top10.empty:
        impact_img = plot_impact_matrix(top10)
        if impact_img:
            story.append(Image(impact_img, width=400, height=230))
            story.append(Spacer(1, 12))
            
    if not trust_gap_df.empty:
        story.append(df_to_table(trust_gap_df.head(10)))
        
    story.append(PageBreak())

    # --- RADAR & DOMAINS ---
    story.append(Paragraph("Operational Stability Profile", styles['Heading1']))
    if domain_scores:
        radar_img = radar_chart(domain_scores)
        story.append(Image(radar_img, width=300, height=300))
        
        domain_df = pd.DataFrame({
            "Domain Capability": list(domain_scores.keys()), 
            "Index Score (0-100)": [f"{v:.1f}" for v in domain_scores.values()]
        })
        story.append(Spacer(1, 12))
        story.append(df_to_table(domain_df, col_widths=[250, 150]))
        
    story.append(PageBreak())

    # --- PARETO ---
    story.append(Paragraph("Structural Risk Debt™: Root Cause Themes", styles['Heading1']))
    if not pareto_df.empty:
        pareto_img = plot_pareto(pareto_df)
        story.append(Image(pareto_img, width=400, height=230))
        story.append(Spacer(1, 12))
        
        display_pareto = pareto_df.copy()
        display_pareto["Cumulative_Pct"] = display_pareto["Cumulative_Pct"].apply(lambda x: f"{x:.1f}%")
        story.append(df_to_table(display_pareto, col_widths=[200, 100, 100]))
    else:
        story.append(Paragraph("No thematic root cause data available.", styles['Normal']))
        
    story.append(PageBreak())

    # --- SERVICE RISK CONCENTRATION ---
    story.append(Paragraph("Service Risk Concentration (Top 10)", styles['Heading1']))
    if not top10.empty:
        heatmap_df = top10.copy()
        if "Service" in heatmap_df.columns:
            if "Service_Tier" not in heatmap_df.columns:
                heatmap_df["Service_Tier"] = "Unspecified"

            req_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Execution_Churn_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
            for c in req_cols:
                if c not in heatmap_df.columns:
                    heatmap_df[c] = 0.0
            
            hm_idx = heatmap_df["Service"].astype(str) + " (" + heatmap_df["Service_Tier"].astype(str) + ")"
            hm_data = heatmap_df[req_cols].copy()
            hm_data.index = hm_idx
            hm_data = hm_data.apply(pd.to_numeric, errors="coerce").fillna(0.0)
            
            hm_img = heatmap_chart(hm_data)
            story.append(Image(hm_img, width=450, height=225))
            story.append(Spacer(1, 12))
            
        display_top10 = top10[["Service", "Total_Service_Risk", "Data_Confidence"]].copy()
        display_top10["Total_Service_Risk"] = display_top10["Total_Service_Risk"].apply(lambda x: f"{float(x):.1f}")
        story.append(df_to_table(display_top10, col_widths=[200, 100, 100]))
    else:
        story.append(Paragraph("No service risk data available.", styles['Normal']))
        
    # --- AUTOMATION STRIKE ZONE ---
    story.append(Spacer(1, 24))
    story.append(Paragraph("Automation Strike Zone (AIOps Readiness)", styles['Heading1']))
    if not automation_df.empty:
        story.append(df_to_table(automation_df, col_widths=[150, 100, 150]))
    else:
        story.append(Paragraph("Insufficient data to calculate automation deficits.", styles['Normal']))

    story.append(PageBreak())

    # --- ROOT CAUSE LEDGER ---
    story.append(Paragraph("Root Cause Ledger", styles['Heading1']))
    story.append(Paragraph("Unfiltered extraction of raw root causes tied directly to operational anchors.", styles['Normal']))
    story.append(Spacer(1, 12))
    
    rca_raw = payload.get("rca_themes_df", pd.DataFrame())
    if not rca_raw.empty and "Underlying_Causes" in rca_raw.columns:
        for _, row in rca_raw.iterrows():
            theme = row.get("Theme", "Unknown Theme")
            freq = row.get("Frequency", 0)
            causes = row.get("Underlying_Causes", [])
            
            story.append(Paragraph(f"<b>{theme}</b> (Frequency: {freq})", styles['Heading3']))
            
            if isinstance(causes, list):
                for cause in causes:
                    bullet = f"• {cause}"
                    story.append(Paragraph(bullet, styles['Normal']))
            story.append(Spacer(1, 8))
    else:
         story.append(Paragraph("No text-based root causes logged in system.", styles['Normal']))

    doc.build(story)
    buffer.seek(0)
    return buffer.read()
