"""OSIL Executive Report - Board-Ready Professional Design"""
import io
import re
from typing import Any, Dict, List, Optional
import matplotlib
matplotlib.use('Agg')
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
    KeepTogether,
)

# Disable TeX to avoid artifacts
plt.rcParams['text.usetex'] = False
plt.rcParams['mathtext.default'] = 'regular'

def _safe_float(val: Any, default: float = 0.0) -> float:
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return default
        return float(val)
    except Exception:
        return default

def _safe_df(val: Any) -> pd.DataFrame:
    if isinstance(val, pd.DataFrame):
        return val.copy()
    return pd.DataFrame()

def _clean_text(text: Any) -> str:
    if text is None:
        return ""
    s = str(text)
    s = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", s)
    s = s.replace("<br>", "<br/>").replace("<br />", "<br/>")
    s = s.replace("™", "&trade;")
    return s

def _get_posture_from_bvsi(bvsi: float) -> str:
    if bvsi >= 85:
        return "High Confidence Operations"
    elif bvsi >= 70:
        return "Controlled and Improving"
    elif bvsi >= 55:
        return "Controlled but Exposed"
    elif bvsi >= 40:
        return "Reactive and Exposed"
    else:
        return "Fragile Operations"

def _styles():
    styles = getSampleStyleSheet()
    
    # Executive Title - Clean, authoritative
    styles.add(ParagraphStyle(
        name="ExecutiveTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=24,
        leading=28,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=6,
    ))
    
    # Subtitle - Metadata
    styles.add(ParagraphStyle(
        name="ExecutiveSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=24,
    ))
    
    # Section Headers - Executive level
    styles.add(ParagraphStyle(
        name="SectionHeader",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=12,
        spaceBefore=18,
    ))
    
    # Sub-section for page headers
    styles.add(ParagraphStyle(
        name="PageHeader",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=16,
    ))
    
    # Executive summary text - Readable at distance
    styles.add(ParagraphStyle(
        name="ExecutiveBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#334155"),
        spaceAfter=12,
    ))
    
    # Caption/annotation text
    styles.add(ParagraphStyle(
        name="Caption",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=8,
    ))
    
    # Table Header - Professional dark
    styles.add(ParagraphStyle(
        name="TableHeader",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.white,
        alignment=1,  # Center
    ))
    
    # Table Cell Body
    styles.add(ParagraphStyle(
        name="TableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=13,
        textColor=colors.HexColor("#334155"),
    ))
    
    # Table Cell for metrics (centered)
    styles.add(ParagraphStyle(
        name="TableMetric",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=13,
        textColor=colors.HexColor("#1E293B"),
        alignment=1,
    ))
    
    # Score indicators
    styles.add(ParagraphStyle(
        name="ScoreGood",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#059669"),
        alignment=1,
    ))
    
    styles.add(ParagraphStyle(
        name="ScoreWarning",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#D97706"),
        alignment=1,
    ))
    
    styles.add(ParagraphStyle(
        name="ScoreCritical",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#DC2626"),
        alignment=1,
    ))
    
    # Footer text
    styles.add(ParagraphStyle(
        name="FooterText",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=8,
        leading=10,
        textColor=colors.HexColor("#94A3B8"),
    ))
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawString(0.75 * inch, 0.5 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence • Confidential")
    canvas.drawRightString(7.75 * inch, 0.5 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _get_score_style(score: float):
    """Return appropriate style name based on score"""
    if score >= 70:
        return "ScoreGood"
    elif score >= 55:
        return "ScoreWarning"
    else:
        return "ScoreCritical"

def _build_radar_chart(domain_scores: Dict[str, float]) -> Optional[io.BytesIO]:
    """Professional radar chart for executives"""
    try:
        domains = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
        labels = ["Service\nResilience", "Change\nGovernance", "Structural\nRisk Debt", "Reliability\nMomentum"]
        
        values = []
        for d in domains:
            if d in domain_scores:
                values.append(_safe_float(domain_scores[d], 0))
            elif d + "™" in domain_scores:
                values.append(_safe_float(domain_scores[d + "™"], 0))
            else:
                values.append(0)
        
        angles = np.linspace(0, 2 * np.pi, len(domains), endpoint=False).tolist()
        values += values[:1]
        angles += angles[:1]
        
        fig, ax = plt.subplots(figsize=(6.5, 5.5), subplot_kw=dict(projection='polar'))
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        # Professional color scheme
        ax.fill(angles, values, color='#3B82F6', alpha=0.2)
        ax.plot(angles, values, color='#2563EB', linewidth=3, marker='o', markersize=6)
        
        # Styling
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(labels, fontsize=11, fontweight='bold', color='#1E293B')
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(['20', '40', '60', '80', '100'], fontsize=9, color='#64748B')
        ax.grid(True, linestyle='--', alpha=0.4, color='#94A3B8')
        
        ax.set_title("Operational Stability Profile", fontsize=14, fontweight='bold', 
                    color='#1E293B', pad=25, y=1.08)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.3)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Radar error: {e}")
        return None

def _build_heatmap(service_risk_df: pd.DataFrame) -> Optional[io.BytesIO]:
    """Executive heatmap with clear risk indicators"""
    df = _safe_df(service_risk_df)
    if df.empty:
        return None
    
    required = ["Service", "Service_Tier"]
    if not all(c in df.columns for c in required):
        return None
    
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_names = {
        "Recurrence_Risk": "Recurrence",
        "MTTR_Drag_Risk": "MTTR Drag", 
        "Reopen_Churn_Risk": "Reopen",
        "Change_Collision_Risk": "Change"
    }
    
    available = [c for c in risk_cols if c in df.columns]
    if not available:
        return None
    
    try:
        hm = df.head(8).copy()  # Limit to top 8 for readability
        hm["Display"] = hm["Service"].astype(str).str[:18] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display")[available]
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0)
        
        fig, ax = plt.subplots(figsize=(7.5, 5.5))
        
        # RdYlGn_r: Red (high risk) to Green (low risk)
        im = ax.imshow(hm.values, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
        
        # Labels
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(hm.columns, fontsize=10, fontweight='bold', color='#1E293B')
        ax.set_yticklabels(hm.index, fontsize=9, color='#334155')
        
        # Colorbar
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label('Risk Score', fontsize=10, fontweight='bold', color='#1E293B')
        cbar.ax.tick_params(labelsize=9, colors='#64748B')
        
        # Annotations
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(hm.iloc[i, j], 0))
                color = 'white' if val > 60 else '#1E293B'
                ax.text(j, i, str(val), ha='center', va='center', 
                       fontsize=9, color=color, fontweight='bold')
        
        ax.set_title('Service Risk Concentration', fontsize=13, fontweight='bold', 
                    color='#1E293B', pad=15)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.2)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap error: {e}")
        return None

def _create_metric_table(data, col_widths, styles_obj):
    """Create professional metric table with color coding"""
    table_data = []
    for i, row in enumerate(data):
        new_row = []
        for j, cell in enumerate(row):
            if i == 0:
                new_row.append(Paragraph(str(cell), styles_obj["TableHeader"]))
            else:
                # Center align metrics, left align text
                if j == 1 and isinstance(cell, str) and any(x in cell for x in ['.', '%']):
                    # Try to parse as number for styling
                    try:
                        val = float(cell.replace('%', ''))
                        style_name = _get_score_style(val) if val <= 100 else "TableMetric"
                        new_row.append(Paragraph(cell, styles_obj[style_name]))
                    except:
                        new_row.append(Paragraph(cell, styles_obj["TableCell"]))
                else:
                    new_row.append(Paragraph(str(cell), styles_obj["TableCell"]))
        table_data.append(new_row)
    
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    style_cmds = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
        ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
        ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
    ]
    
    # Add subtle row separators
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            style_cmds.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")))
        style_cmds.append(('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")))
    
    t.setStyle(TableStyle(style_cmds))
    return t

def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    try:
        styles = _styles()
        out = io.BytesIO()
        
        doc = SimpleDocTemplate(
            out,
            pagesize=LETTER,
            rightMargin=0.75 * inch,
            leftMargin=0.75 * inch,
            topMargin=0.75 * inch,
            bottomMargin=0.75 * inch,
        )

        # Extract data
        tenant_name = str(payload.get("tenant_name", "Organization"))
        as_of = str(payload.get("as_of", ""))
        bvsi = _safe_float(payload.get("bvsi", 0))
        posture = str(payload.get("posture", _get_posture_from_bvsi(bvsi)))
        exec_summary = _clean_text(payload.get("executive_interpretation", ""))
        domain_scores = payload.get("domain_scores", {}) or {}
        service_risk_df = _safe_df(payload.get("service_risk_top10"))
        sip_candidates = _safe_df(payload.get("sip_candidates"))
        detected_dataset = str(payload.get("detected_dataset", "INCIDENT")).upper()
        service_anchor = str(payload.get("service_anchor_used", "Service"))
        readiness = _safe_float(payload.get("data_readiness_score", 0))

        story = []

        # ==================== PAGE 1: EXECUTIVE DASHBOARD ====================
        story.append(Paragraph("OSIL™ Executive Briefing", styles["ExecutiveTitle"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report — {tenant_name}<br/>As of {as_of}", 
                              styles["ExecutiveSubtitle"]))
        
        # Key Metrics - 3 Column Layout
        metric_data = [
            ["BVSI™ Score", "Operating Posture", "Data Quality"],
            [f"{bvsi:.1f}", posture, f"{readiness:.0f}%"]
        ]
        
        # Style the score cell based on value
        score_style = _get_score_style(bvsi)
        metric_table = Table(metric_data, colWidths=[2.2*inch, 3.0*inch, 1.3*inch])
        metric_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, 1), 12),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 12),
            ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (0, 1), 18),
            ('TEXTCOLOR', (0, 1), (0, 1), colors.HexColor("#2563EB") if bvsi >= 70 else colors.HexColor("#D97706") if bvsi >= 55 else colors.HexColor("#DC2626")),
            ('FONTNAME', (1, 1), (1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (1, 1), (1, 1), 12),
            ('FONTNAME', (2, 1), (2, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (2, 1), (2, 1), 14),
        ]))
        story.append(metric_table)
        story.append(Spacer(1, 20))

        # BVSI Scale Reference
        story.append(Paragraph("BVSI™ Interpretation Framework", styles["SectionHeader"]))
        
        scale_data = [
            ["Range", "Posture", "Executive Action"],
            ["85 – 100", "High Confidence", "Invest in scaling operations; stability enables growth"],
            ["70 – 84", "Controlled & Improving", "Maintain course; selective SIPs in weak domains"],
            ["55 – 69", "Controlled but Exposed", "Immediate SIP activation in Structural Risk Debt"],
            ["40 – 54", "Reactive & Exposed", "Executive intervention required; halt non-critical changes"],
            ["< 40", "Fragile Operations", "Emergency stabilization; board-level crisis protocol"],
        ]
        
        scale_table = _create_metric_table(scale_data, [1.0*inch, 1.8*inch, 3.7*inch], styles)
        story.append(scale_table)
        story.append(Spacer(1, 20))

        # Executive Signal Box
        story.append(Paragraph("Executive Signal", styles["SectionHeader"]))
        signal = f"Current operational posture is <b>{posture}</b> (BVSI™ {bvsi:.1f}). "
        signal += "Governance mechanisms are active, but concentrated instability in specific service domains requires targeted executive attention and SIP (Service Improvement Program) activation."
        story.append(Paragraph(signal, styles["ExecutiveBody"]))
        
        story.append(Spacer(1, 12))
        
        # Key Takeaways (Critical for board)
        kt_header = Paragraph("Strategic Imperatives", styles["SectionHeader"])
        
        strongest = max(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Change Governance"
        weakest = min(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Structural Risk Debt"
        top_svc = str(sip_candidates.iloc[0].get("Service", "Priority Service")) if not sip_candidates.empty else "Critical Service"
        
        kt_data = [
            ["Strategic Insight", "Board Action"],
            [f"<b>Strength:</b> {strongest} demonstrates operational maturity", "Leverage as model for organization-wide standards"],
            [f"<b>Exposure:</b> {weakest} represents highest risk concentration", f"Immediate SIP funding for {top_svc} and peer services"],
            [f"<b>Investment Priority:</b> Top-quartile instability in {top_svc}", "Assign executive sponsor; 30-day remediation sprint"],
        ]
        
        kt_table = Table(kt_data, colWidths=[3.4*inch, 3.1*inch])
        kt_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Changed to TOP for multi-line text
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 1), (-1, 1), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEBELOW', (0, 2), (-1, 2), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEBELOW', (0, 3), (-1, 3), 2, colors.HexColor("#1E293B")),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#F8FAFC")),
        ]))
        
        # Keep Strategic Imperatives together
        story.append(KeepTogether([kt_header, Spacer(1, 6), kt_table]))
        
        story.append(Spacer(1, 16))
        story.append(Paragraph(f"<i>Data Sources: {detected_dataset} | Classification: Confidential</i>", styles["Caption"]))

        # Force page break after page 1 content
        story.append(PageBreak())

        # ==================== PAGE 2: STABILITY PROFILE ====================
        story.append(Paragraph("Operational Stability Profile", styles["PageHeader"]))
        
        # Two-column layout: Chart on left, Domain breakdown on right
        radar_img = _build_radar_chart(domain_scores)
        
        domain_rows = [["Domain", "Score", "Assessment"]]
        domain_defs = {
            "Service Resilience": "Recovery capability & incident response",
            "Change Governance": "Change control & collision prevention",
            "Structural Risk Debt": "Accumulated unresolved weaknesses",
            "Reliability Momentum": "Trend direction of stability"
        }
        
        for key in ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]:
            display = "Structural Risk Debt™" if key == "Structural Risk Debt" else key
            score = _safe_float(domain_scores.get(key, domain_scores.get(key + "™", 0)), 0)
            assessment = "Strong" if score >= 80 else "Improving" if score >= 60 else "At Risk" if score >= 40 else "Critical"
            domain_rows.append([display, f"{score:.1f}", assessment])
        
        # Create table for domains
        domain_table = Table(domain_rows, colWidths=[2.4*inch, 0.9*inch, 1.8*inch])
        domain_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
        ] + [
            # Color code the score cells
            ('TEXTCOLOR', (1, i), (1, i), 
             colors.HexColor("#059669") if _safe_float(domain_rows[i][1], 0) >= 70 
             else colors.HexColor("#D97706") if _safe_float(domain_rows[i][1], 0) >= 55 
             else colors.HexColor("#DC2626"))
            for i in range(1, len(domain_rows))
        ] + [
            ('FONTNAME', (1, i), (1, i), 'Helvetica-Bold') for i in range(1, len(domain_rows))
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(domain_rows))
        ]))
        
        # Place chart and table side by side conceptually, or stack for clarity
        if radar_img:
            story.append(Image(radar_img, width=5.5*inch, height=4.5*inch))
            story.append(Spacer(1, 8))
        
        story.append(Paragraph("Domain Scorecard", styles["SectionHeader"]))
        story.append(domain_table)
        story.append(Spacer(1, 10))
        
        story.append(Paragraph(
            "<i>Scoring: 80-100 Strong Maturity | 60-79 Controlled | 40-59 Weakness | <40 Structural Fragility</i>", 
            styles["Caption"]
        ))

        story.append(PageBreak())

        # ==================== PAGE 3: SERVICE IMPROVEMENT ====================
        story.append(Paragraph("Service Improvement Priorities", styles["PageHeader"]))
        story.append(Paragraph(
            "High-impact SIPs for executive sponsorship. These services represent the highest concentration of operational risk and require immediate intervention.", 
            styles["ExecutiveBody"]
        ))
        story.append(Spacer(1, 12))
        
        if not sip_candidates.empty:
            # Top 3 Executive Brief
            story.append(Paragraph("Immediate Executive Actions", styles["SectionHeader"]))
            
            top3_data = [["Service", "Tier", "Risk Driver", "Executive Decision"]]
            for idx, (_, row) in enumerate(sip_candidates.head(3).iterrows(), 1):
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "-"))
                why = str(row.get("Why_Flagged", "Risk Concentration"))
                # Truncate why for table
                why_short = why[:40] + "..." if len(why) > 40 else why
                action = "Activate SIP" if idx == 1 else "Monitor Closely" if idx == 2 else "Include in Q Plan"
                top3_data.append([svc, tier, why_short, action])
            
            top3_table = _create_metric_table(top3_data, [1.8*inch, 0.7*inch, 2.5*inch, 1.5*inch], styles)
            story.append(top3_table)
            story.append(Spacer(1, 20))
            
            # Full SIP Portfolio
            story.append(Paragraph("SIP Portfolio Overview", styles["SectionHeader"]))
            
            sip_data = [["Service", "Tier", "Theme", "Priority", "Risk Score"]]
            for _, row in sip_candidates.iterrows():
                svc = str(row.get("Service", "-"))
                tier = str(row.get("Service_Tier", "-"))
                theme = str(row.get("Suggested_Theme", "Stability"))[:15]
                priority = str(row.get("Priority_Label", "-"))
                score = _safe_float(row.get("SIP_Priority_Score", row.get("sip_priority_score", 0)), 0)
                sip_data.append([svc, tier, theme, priority, f"{score:.1f}"])
            
            sip_table = Table(sip_data, colWidths=[1.8*inch, 0.6*inch, 1.4*inch, 0.9*inch, 0.9*inch])
            sip_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 9),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(sip_data))
            ] + [
                ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#FEF3C7")) 
                for i in range(1, len(sip_data)) if sip_data[i][3] == "Next SIP"
            ]))
            story.append(sip_table)

        story.append(PageBreak())

        # ==================== PAGE 4: RISK HEATMAP ====================
        story.append(Paragraph("Service Risk Concentration Matrix", styles["PageHeader"]))
        story.append(Paragraph(
            "Cross-dimensional risk analysis across the top 8 services. Dark red indicates critical risk concentration requiring immediate executive intervention.", 
            styles["ExecutiveBody"]
        ))
        story.append(Spacer(1, 10))
        
        hm_img = _build_heatmap(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=6.5*inch, height=4.8*inch))
            story.append(Spacer(1, 16))
        
        # Heatmap interpretation
        interp_header = Paragraph("Risk Dimension Definitions", styles["SectionHeader"])
        interp_text = Paragraph(
            "<b>Recurrence:</b> Frequency of repeated incidents indicating unresolved root causes.<br/>"
            "<b>MTTR Drag:</b> Mean Time To Recovery exceeds organizational SLAs.<br/>"
            "<b>Reopen:</b> Incidents closed prematurely, indicating quality gaps in resolution validation.<br/>"
            "<b>Change:</b> Instability correlated with change windows, indicating governance gaps.", 
            styles["ExecutiveBody"]
        )
        
        story.append(KeepTogether([interp_header, Spacer(1, 6), interp_text]))

        # ==================== PAGE 5: APPENDIX ====================
        story.append(PageBreak())
        story.append(Paragraph("Appendix: Risk Scoring Methodology", styles["PageHeader"]))
        story.append(Spacer(1, 12))
        
        story.append(Paragraph("Risk Score Calibration", styles["SectionHeader"]))
        
        risk_data = [
            ["Score Range", "Risk Level", "Executive Response", "Investment Priority"],
            ["0 – 30", "Green / Stable", "Standard operations, trend monitoring", "Maintenance funding"],
            ["31 – 60", "Yellow / Elevated", "Operational review within 30 days", "Preventive investment"],
            ["61 – 80", "Orange / High", "Executive sponsor assignment, SIP activation", "Priority resourcing"],
            ["81 – 100", "Red / Critical", "Crisis protocol, board notification, change freeze", "Emergency budget allocation"],
        ]
        
        risk_table = _create_metric_table(risk_data, [1.1*inch, 1.3*inch, 2.2*inch, 1.9*inch], styles)
        story.append(risk_table)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("About This Report", styles["SectionHeader"]))
        story.append(Paragraph(
            "The Operational Stability Intelligence Layer (OSIL™) synthesizes incident, change, and problem data to provide "
            "executive-level visibility into technology operational stability. The BVSI™ (Business Value Stability Index) is a "
            "composite metric measuring organizational capability to maintain service stability while executing change.<br/><br/>"
            f"<b>Data Foundation:</b> This analysis is based on {detected_dataset} data using {service_anchor} as the "
            "operational anchor for service classification.", 
            styles["ExecutiveBody"]
        ))

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
