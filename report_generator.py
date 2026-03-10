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
    
    styles.add(ParagraphStyle(
        name="ExecutiveTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=24,
        leading=28,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=6,
    ))
    
    styles.add(ParagraphStyle(
        name="ExecutiveSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=14,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=24,
    ))
    
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
    
    styles.add(ParagraphStyle(
        name="PageHeader",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#1E293B"),
        spaceAfter=16,
    ))
    
    styles.add(ParagraphStyle(
        name="ExecutiveBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=16,
        textColor=colors.HexColor("#334155"),
        spaceAfter=12,
    ))
    
    styles.add(ParagraphStyle(
        name="Caption",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=8,
    ))
    
    styles.add(ParagraphStyle(
        name="TableHeader",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=12,
        textColor=colors.white,
    ))
    
    styles.add(ParagraphStyle(
        name="TableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#334155"),
    ))
    
    styles.add(ParagraphStyle(
        name="TableCellBold",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#1E293B"),
    ))
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawString(0.75 * inch, 0.5 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence • Confidential")
    canvas.drawRightString(7.75 * inch, 0.5 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _build_radar_chart(domain_scores: Dict[str, float]) -> Optional[io.BytesIO]:
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
        
        ax.fill(angles, values, color='#3B82F6', alpha=0.2)
        ax.plot(angles, values, color='#2563EB', linewidth=3, marker='o', markersize=6)
        
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
        hm = df.head(8).copy()
        hm["Display"] = hm["Service"].astype(str).str[:18] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display")[available]
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0)
        
        fig, ax = plt.subplots(figsize=(7.5, 5.5))
        
        im = ax.imshow(hm.values, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
        
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(hm.columns, fontsize=10, fontweight='bold', color='#1E293B')
        ax.set_yticklabels(hm.index, fontsize=9, color='#334155')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label('Risk Score', fontsize=10, fontweight='bold', color='#1E293B')
        cbar.ax.tick_params(labelsize=9, colors='#64748B')
        
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

        tenant_name = str(payload.get("tenant_name", "Organization"))
        as_of = str(payload.get("as_of", ""))
        bvsi = _safe_float(payload.get("bvsi", 0))
        posture = str(payload.get("posture", _get_posture_from_bvsi(bvsi)))
        domain_scores = payload.get("domain_scores", {}) or {}
        service_risk_df = _safe_df(payload.get("service_risk_top10"))
        sip_candidates = _safe_df(payload.get("sip_candidates"))
        detected_dataset = str(payload.get("detected_dataset", "INCIDENT")).upper()
        service_anchor = str(payload.get("service_anchor_used", "Service"))
        readiness = _safe_float(payload.get("data_readiness_score", 0))

        story = []

        # ==================== PAGE 1 ====================
        story.append(Paragraph("OSIL™ Executive Briefing", styles["ExecutiveTitle"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report — {tenant_name}<br/>As of {as_of}", 
                              styles["ExecutiveSubtitle"]))
        
        # Key Metrics
        metric_data = [
            ["BVSI™ Score", "Operating Posture", "Data Quality"],
            [f"{bvsi:.1f}", posture, f"{readiness:.0f}%"]
        ]
        
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

        # BVSI Scale
        story.append(Paragraph("BVSI™ Interpretation Framework", styles["SectionHeader"]))
        
        scale_data = [
            ["Range", "Posture", "Executive Action"],
            ["85 – 100", "High Confidence", "Invest in scaling operations; stability enables growth"],
            ["70 – 84", "Controlled & Improving", "Maintain course; selective SIPs in weak domains"],
            ["55 – 69", "Controlled but Exposed", "Immediate SIP activation in Structural Risk Debt"],
            ["40 – 54", "Reactive & Exposed", "Executive intervention required; halt non-critical changes"],
            ["< 40", "Fragile Operations", "Emergency stabilization; board-level crisis protocol"],
        ]
        
        scale_table = Table(scale_data, colWidths=[1.0*inch, 1.8*inch, 3.7*inch])
        scale_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(scale_data))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(scale_data), 2)
        ]))
        story.append(scale_table)
        story.append(Spacer(1, 20))

        # Executive Signal
        story.append(Paragraph("Executive Signal", styles["SectionHeader"]))
        signal = f"Current operational posture is {posture} (BVSI™ {bvsi:.1f}). "
        signal += "Governance mechanisms are active, but concentrated instability in specific service domains requires targeted executive attention and SIP activation."
        story.append(Paragraph(signal, styles["ExecutiveBody"]))
        story.append(Spacer(1, 16))
        
        # ==================== FIXED STRATEGIC IMPERATIVES ====================
        story.append(Paragraph("Strategic Imperatives", styles["SectionHeader"]))
        
        strongest = max(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Change Governance"
        weakest = min(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Structural Risk Debt"
        top_svc = str(sip_candidates.iloc[0].get("Service", "Priority Service")) if not sip_candidates.empty else "Critical Service"
        
        # Create a clean 2-column layout with proper spacing
        imp_data = [
            [Paragraph("Strategic Insight", styles["TableHeader"]), 
             Paragraph("Board Action", styles["TableHeader"])],
            [Paragraph(f"Strength: {strongest} demonstrates mature controls and should be leveraged as the operational standard for other domains.", styles["TableCell"]),
             Paragraph("Codify practices into enterprise standards. Expand successful patterns to underperforming domains.", styles["TableCell"])],
            [Paragraph(f"Exposure: {weakest} represents the highest concentration of stability risk across the service portfolio.", styles["TableCell"]),
             Paragraph(f"Authorize immediate SIP funding for {top_svc} and peer services in the bottom quartile.", styles["TableCell"])],
            [Paragraph(f"Investment Priority: {top_svc} exhibits top-quartile instability requiring executive intervention.", styles["TableCell"]),
             Paragraph("Assign executive sponsor immediately. Initiate 30-day remediation sprint with weekly board updates.", styles["TableCell"])]
        ]
        
        # Wider table with more space
        imp_table = Table(imp_data, colWidths=[3.5*inch, 3.0*inch])
        imp_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 1), (-1, 1), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEBELOW', (0, 2), (-1, 2), 0.5, colors.HexColor("#E2E8F0")),
            ('LINEBELOW', (0, 3), (-1, 3), 2, colors.HexColor("#1E293B")),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#F8FAFC")),
        ]))
        
        story.append(imp_table)
        story.append(Spacer(1, 16))
        story.append(Paragraph(f"Data Sources: {detected_dataset} | Classification: Confidential", styles["Caption"]))

        story.append(PageBreak())

        # ==================== PAGE 2: STABILITY PROFILE ====================
        story.append(Paragraph("Operational Stability Profile", styles["PageHeader"]))
        
        radar_img = _build_radar_chart(domain_scores)
        if radar_img:
            story.append(Image(radar_img, width=5.5*inch, height=4.5*inch))
            story.append(Spacer(1, 8))
        
        story.append(Paragraph("Domain Scorecard", styles["SectionHeader"]))
        
        domain_rows = [[Paragraph("Domain", styles["TableHeader"]), 
                       Paragraph("Score", styles["TableHeader"]), 
                       Paragraph("Assessment", styles["TableHeader"])]]
        
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
            
            score_color = colors.HexColor("#059669") if score >= 70 else colors.HexColor("#D97706") if score >= 55 else colors.HexColor("#DC2626")
            
            domain_rows.append([
                Paragraph(display, styles["TableCell"]),
                Paragraph(f"{score:.1f}", ParagraphStyle(name='score', parent=styles["TableCell"], textColor=score_color, fontName='Helvetica-Bold')),
                Paragraph(assessment, styles["TableCell"])
            ])
        
        domain_table = Table(domain_rows, colWidths=[2.8*inch, 1.0*inch, 1.7*inch])
        domain_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(domain_rows))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(domain_rows), 2)
        ]))
        
        story.append(domain_table)
        story.append(Spacer(1, 10))
        story.append(Paragraph("Scoring: 80-100 Strong | 60-79 Controlled | 40-59 Weakness | <40 Fragility", styles["Caption"]))

        story.append(PageBreak())

        # ==================== PAGE 3: SERVICE IMPROVEMENT ====================
        story.append(Paragraph("Service Improvement Priorities", styles["PageHeader"]))
        story.append(Paragraph("High-impact SIPs for executive sponsorship. These services represent the highest concentration of operational risk.", styles["ExecutiveBody"]))
        story.append(Spacer(1, 12))
        
        if not sip_candidates.empty:
            story.append(Paragraph("Immediate Executive Actions", styles["SectionHeader"]))
            
            top3_data = [[Paragraph("Service", styles["TableHeader"]), 
                         Paragraph("Tier", styles["TableHeader"]),
                         Paragraph("Risk Driver", styles["TableHeader"]), 
                         Paragraph("Decision", styles["TableHeader"])]]
            
            for idx, (_, row) in enumerate(sip_candidates.head(3).iterrows(), 1):
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "-"))
                why = str(row.get("Why_Flagged", "Risk"))
                why_short = (why[:35] + "...") if len(why) > 35 else why
                action = "Activate SIP" if idx == 1 else "Monitor" if idx == 2 else "Plan"
                
                top3_data.append([
                    Paragraph(svc, styles["TableCell"]),
                    Paragraph(tier, ParagraphStyle(name='center', parent=styles["TableCell"], alignment=1)),
                    Paragraph(why_short, styles["TableCell"]),
                    Paragraph(action, ParagraphStyle(name='center', parent=styles["TableCell"], alignment=1))
                ])
            
            top3_table = Table(top3_data, colWidths=[2.0*inch, 0.8*inch, 2.8*inch, 1.2*inch])
            top3_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(top3_data))
            ] + [
                ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#FEF3C7")) for i in [1]
            ]))
            story.append(top3_table)
            story.append(Spacer(1, 20))
            
            story.append(Paragraph("SIP Portfolio", styles["SectionHeader"]))
            
            sip_data = [[Paragraph("Service", styles["TableHeader"]), 
                        Paragraph("Tier", styles["TableHeader"]),
                        Paragraph("Theme", styles["TableHeader"]),
                        Paragraph("Priority", styles["TableHeader"]),
                        Paragraph("Score", styles["TableHeader"])]]
            
            for _, row in sip_candidates.iterrows():
                svc = str(row.get("Service", "-"))
                tier = str(row.get("Service_Tier", "-"))
                theme = str(row.get("Suggested_Theme", "Stability"))[:12]
                priority = str(row.get("Priority_Label", "-"))
                score = _safe_float(row.get("SIP_Priority_Score", row.get("sip_priority_score", 0)), 0)
                
                sip_data.append([
                    Paragraph(svc, styles["TableCell"]),
                    Paragraph(tier, ParagraphStyle(name='c1', parent=styles["TableCell"], alignment=1)),
                    Paragraph(theme, ParagraphStyle(name='c2', parent=styles["TableCell"], alignment=1)),
                    Paragraph(priority, ParagraphStyle(name='c3', parent=styles["TableCell"], alignment=1)),
                    Paragraph(f"{score:.1f}", ParagraphStyle(name='c4', parent=styles["TableCell"], alignment=1))
                ])
            
            sip_table = Table(sip_data, colWidths=[2.0*inch, 0.7*inch, 1.3*inch, 1.0*inch, 0.9*inch])
            sip_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 8),
                ('TOPPADDING', (0, 0), (-1, -1), 6),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
                ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(sip_data))
            ] + [
                ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#FEF3C7")) 
                for i in range(1, len(sip_data)) if sip_data[i][3].text == "Next SIP"
            ]))
            story.append(sip_table)

        story.append(PageBreak())

        # ==================== PAGE 4: RISK HEATMAP ====================
        story.append(Paragraph("Service Risk Concentration Matrix", styles["PageHeader"]))
        story.append(Paragraph("Cross-dimensional risk analysis. Dark red indicates critical risk concentration requiring immediate executive intervention.", styles["ExecutiveBody"]))
        story.append(Spacer(1, 10))
        
        hm_img = _build_heatmap(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=6.5*inch, height=4.8*inch))
            story.append(Spacer(1, 16))
        
        story.append(Paragraph("Risk Dimension Definitions", styles["SectionHeader"]))
        story.append(Paragraph(
            "Recurrence: Frequency of repeated incidents indicating unresolved root causes. "
            "MTTR Drag: Recovery time exceeds SLAs. "
            "Reopen: Incidents closed prematurely. "
            "Change: Instability correlated with change windows.", 
            styles["ExecutiveBody"]
        ))

        story.append(PageBreak())

        # ==================== PAGE 5: APPENDIX ====================
        story.append(Paragraph("Appendix: Risk Scoring Methodology", styles["PageHeader"]))
        story.append(Spacer(1, 12))
        
        story.append(Paragraph("Risk Score Calibration", styles["SectionHeader"]))
        
        risk_data = [
            ["Score", "Level", "Response", "Investment"],
            ["0 – 30", "Green / Stable", "Standard operations, monitoring", "Maintenance"],
            ["31 – 60", "Yellow / Elevated", "Operational review in 30 days", "Preventive"],
            ["61 – 80", "Orange / High", "Executive sponsor, SIP activation", "Priority"],
            ["81 – 100", "Red / Critical", "Crisis protocol, change freeze", "Emergency"],
        ]
        
        risk_table = Table(risk_data, colWidths=[1.0*inch, 1.4*inch, 2.4*inch, 1.7*inch])
        risk_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEABOVE', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(risk_data))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(risk_data), 2)
        ]))
        story.append(risk_table)
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("About This Report", styles["SectionHeader"]))
        story.append(Paragraph(
            f"The Operational Stability Intelligence Layer (OSIL™) synthesizes incident, change, and problem data. "
            f"The BVSI™ measures organizational capability to maintain service stability. "
            f"This analysis is based on {detected_dataset} data using {service_anchor} as the operational anchor.", 
            styles["ExecutiveBody"]
        ))

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
