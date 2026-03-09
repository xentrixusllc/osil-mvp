"""OSIL Executive PDF Report Generator - Consulting Style Fixed"""
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
)

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

def _get_posture_color(bvsi: float) -> str:
    """Return hex color based on BVSI score"""
    if bvsi >= 85:
        return "#059669"  # Green
    elif bvsi >= 70:
        return "#2563EB"  # Blue
    elif bvsi >= 55:
        return "#D97706"  # Amber/Orange
    elif bvsi >= 40:
        return "#DC2626"  # Red
    else:
        return "#7C2D12"  # Dark Red

def _styles():
    styles = getSampleStyleSheet()
    
    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=28,
            leading=32,
            textColor=colors.HexColor("#0F172A"),
            spaceAfter=4,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",  # FIXED: removed -Light
            fontSize=12,
            leading=14,
            textColor=colors.HexColor("#64748B"),
            spaceAfter=20,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_ReportMeta",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#94A3B8"),
            spaceAfter=24,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=18,
            textColor=colors.HexColor("#0F172A"),
            spaceAfter=12,
            spaceBefore=20,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_SubSection",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=14,
            textColor=colors.HexColor("#334155"),
            spaceAfter=8,
            spaceBefore=12,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#334155"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#64748B"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Label",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#64748B"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_ScoreBig",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=48,
            leading=52,
            textColor=colors.HexColor("#0F172A"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Status",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=11,
            leading=13,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#FFFFFF"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableCell",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#334155"),
        )
    )
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawString(0.7 * inch, 0.4 * inch, "OSIL™ by Xentrixus | Operational Stability Intelligence | Confidential & Proprietary")
    canvas.drawRightString(7.8 * inch, 0.4 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _horizontal_rule(width, color="#E2E8F0"):
    """Create a horizontal line"""
    t = Table([[""]], colWidths=[width], rowHeights=[1])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(color)),
        ('LINEBELOW', (0, 0), (-1, -1), 1, colors.HexColor(color)),
    ]))
    return t

def _build_radar_image(domain_scores: Dict[str, float]) -> Optional[io.BytesIO]:
    try:
        labels = list(domain_scores.keys())
        values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]
        
        if not labels:
            labels = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
            values = [0, 0, 0, 0]
        
        angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
        angles_loop = angles + [angles[0]]
        values_loop = values + [values[0]]
        
        fig = plt.figure(figsize=(5, 4), dpi=120)
        ax = fig.add_subplot(111, polar=True)
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        ax.plot(angles_loop, values_loop, linewidth=2.5, color='#0F172A')
        ax.fill(angles_loop, values_loop, alpha=0.15, color='#3B82F6')
        ax.set_xticks(angles)
        ax.set_xticklabels(labels, fontsize=8)
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
        ax.set_title("Stability Profile", pad=20, fontsize=11, fontweight='bold', color='#0F172A')
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white', edgecolor='none', pad_inches=0.1)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Radar chart error: {e}")
        return None

def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> Optional[io.BytesIO]:
    df = _safe_df(service_risk_top10)
    if df.empty or len(df) == 0:
        return None
    
    required_cols = ["Service", "Service_Tier"]
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    
    if not all(c in df.columns for c in required_cols):
        return None
    
    available_risks = [c for c in risk_cols if c in df.columns]
    if not available_risks:
        return None
    
    try:
        hm = df.head(10).copy()
        hm["Display_Name"] = hm["Service"].astype(str).str[:16] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        
        display_names = {
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen",
            "Change_Collision_Risk": "Change"
        }
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        height = min(5.5, max(3.5, len(hm) * 0.35))
        fig, ax = plt.subplots(figsize=(6.5, height), dpi=120)
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='RdYlGn_r')
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(list(hm.columns), fontsize=9)
        ax.set_yticklabels(list(hm.index), fontsize=8)
        ax.set_title("Service Risk Heatmap", fontsize=11, fontweight='bold', pad=12, color='#0F172A')
        
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(float(hm.iloc[i, j]), 0))
                text_color = "white" if val > 60 else "black"
                ax.text(j, i, str(val), ha="center", va="center", 
                       fontsize=8, color=text_color, fontweight='bold')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label("Risk Score", fontsize=9)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white', edgecolor='none', pad_inches=0.15)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap error: {e}")
        return None

def _make_table(data, col_widths, styles):
    """Create professional table"""
    if not data or len(data) == 0:
        t = Table([["No data"]], colWidths=[6*inch])
        t.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))
        return t
    
    max_cols = max(len(row) for row in data)
    norm_data = []
    for i, row in enumerate(data):
        new_row = []
        for cell in row:
            if i == 0:
                new_row.append(Paragraph(str(cell), styles["OSIL_TableHeader"]))
            else:
                new_row.append(Paragraph(str(cell), styles["OSIL_TableCell"]))
        while len(new_row) < max_cols:
            new_row.append("" if i > 0 else Paragraph("", styles["OSIL_TableHeader"]))
        norm_data.append(new_row)
    
    if not col_widths or len(col_widths) != max_cols:
        col_widths = [6.0/max_cols*inch] * max_cols
    
    t = Table(norm_data, colWidths=col_widths, repeatRows=1)
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LINEABOVE', (0, 0), (-1, 0), 1, colors.HexColor("#0F172A")),
        ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor("#0F172A")),
        ('LINEBELOW', (0, -1), (-1, -1), 1, colors.HexColor("#CBD5E1")),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    
    if len(norm_data) > 1:
        for i in range(1, len(norm_data)):
            style_commands.append(('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")))
    
    t.setStyle(TableStyle(style_commands))
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
            topMargin=0.6 * inch,
            bottomMargin=0.75 * inch,
        )

        tenant_name = str(payload.get("tenant_name", "Default"))
        as_of = str(payload.get("as_of", ""))
        bvsi = _safe_float(payload.get("bvsi", 0.0))
        posture = str(payload.get("posture", "Unknown"))
        exec_summary = _clean_text(payload.get("executive_interpretation", ""))
        domain_scores = payload.get("domain_scores", {}) or {}
        service_risk_df = _safe_df(payload.get("service_risk_top10"))
        sip_candidates = _safe_df(payload.get("sip_candidates"))
        detected_dataset = str(payload.get("detected_dataset", "unknown")).upper()
        service_anchor = str(payload.get("service_anchor_used", "None"))
        readiness = _safe_float(payload.get("data_readiness_score", 0.0))

        story = []

        # EXECUTIVE HEADER
        story.append(Paragraph("OSIL™", styles["OSIL_Title"]))
        story.append(Paragraph("Operational Stability Intelligence Report", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"Prepared for: <b>{tenant_name}</b> | Report Date: {as_of} | Classification: Confidential", 
                              styles["OSIL_ReportMeta"]))
        
        story.append(_horizontal_rule(6.5*inch))
        story.append(Spacer(1, 20))
        
        # Key Metrics
        posture_color = _get_posture_color(bvsi)
        
        status_badge = Table([[Paragraph(posture.upper(), ParagraphStyle(
            name='badge', fontName='Helvetica-Bold', fontSize=9, textColor=colors.white, alignment=1
        ))]], colWidths=[1.8*inch], rowHeights=[24])
        status_badge.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor(posture_color)),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))
        
        metric_row = Table([
            [
                Paragraph("BVSI™ SCORE", styles["OSIL_Label"]),
                Paragraph("OPERATING STATUS", styles["OSIL_Label"]),
                Paragraph("DATA QUALITY", styles["OSIL_Label"]),
            ],
            [
                Paragraph(f"{bvsi:.1f}", styles["OSIL_ScoreBig"]),
                status_badge,
                Paragraph(f"{readiness:.0f}%", ParagraphStyle(
                    name='pct', fontName='Helvetica-Bold', fontSize=32, leading=36, textColor=colors.HexColor("#0F172A")
                )),
            ]
        ], colWidths=[2.0*inch, 2.5*inch, 2.0*inch])
        
        metric_row.setStyle(TableStyle([
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'RIGHT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
        ]))
        
        story.append(metric_row)
        story.append(Spacer(1, 8))
        story.append(_horizontal_rule(6.5*inch, "#CBD5E1"))
        story.append(Spacer(1, 12))
        
        # Executive Summary
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(exec_summary, styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        # Key Insights
        story.append(Paragraph("Key Insights", styles["OSIL_SubSection"]))
        strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        top_svc = "N/A"
        if not sip_candidates.empty:
            top_svc = str(sip_candidates.iloc[0].get("Service", "N/A"))
        
        insights_data = [
            ["Primary Strength", "Critical Gap", "Immediate Action"],
            [f"<b>{strongest}</b><br/>Performing above target", 
             f"<b>{weakest}</b><br/>Requires executive focus", 
             f"<b>{top_svc}</b><br/>Initiate SIP program"],
        ]
        insights_table = Table(insights_data, colWidths=[2.1*inch, 2.1*inch, 2.3*inch])
        insights_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F1F5F9")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#64748B")),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 8),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 6),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
            ('TOPPADDING', (0, 1), (-1, 1), 8),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor("#CBD5E1")),
        ]))
        story.append(insights_table)
        story.append(Spacer(1, 12))
        
        story.append(Paragraph(f"<i>Data Sources: {detected_dataset} | Service Classification: {service_anchor}</i>", 
                              ParagraphStyle(name='meta', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#94A3B8"))))

        # PAGE 2
        story.append(PageBreak())
        story.append(Paragraph("Stability Assessment", styles["OSIL_Section"]))
        story.append(Paragraph("Four-domain analysis of operational maturity.", styles["OSIL_Body"]))
        story.append(Spacer(1, 12))
        
        radar_img = _build_radar_image(domain_scores)
        if radar_img:
            story.append(Image(radar_img, width=4.0*inch, height=3.2*inch))
        
        story.append(Spacer(1, 16))
        story.append(Paragraph("Domain Definitions", styles["OSIL_SubSection"]))
        
        domain_data = [
            ["Domain", "Score", "Business Impact"],
            ["Service Resilience", f"{domain_scores.get('Service Resilience', 0):.0f}", 
             "Recovery speed and incident restoration capability"],
            ["Change Governance", f"{domain_scores.get('Change Governance', 0):.0f}", 
             "Change control effectiveness and collision prevention"],
            ["Structural Risk Debt™", f"{domain_scores.get('Structural Risk Debt™', 0):.0f}", 
             "Accumulated unresolved operational weaknesses"],
            ["Reliability Momentum", f"{domain_scores.get('Reliability Momentum', 0):.0f}", 
             "Direction of operational reliability trends"],
        ]
        story.append(_make_table(domain_data, [1.8*inch, 0.8*inch, 3.9*inch], styles))

        # PAGE 3
        story.append(PageBreak())
        story.append(Paragraph("Strategic Improvement Priorities", styles["OSIL_Section"]))
        story.append(Paragraph("30-60 day high-impact initiatives for executive action.", styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        if not sip_candidates.empty:
            story.append(Paragraph("Priority Initiatives", styles["OSIL_SubSection"]))
            
            sip_data = [["Rank", "Service", "Tier", "Focus", "Status", "Business Case"]]
            for idx, (_, row) in enumerate(sip_candidates.head(5).iterrows(), 1):
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "Unknown"))
                theme = str(row.get("Suggested_Theme", "Stability"))
                priority = str(row.get("Priority_Label", "Monitor"))
                why = str(row.get("Why_Flagged", "Risk"))
                sip_data.append([f"#{idx}", svc, tier, theme, priority, why])
            
            story.append(_make_table(sip_data, [0.5*inch, 1.6*inch, 0.7*inch, 1.2*inch, 0.9*inch, 1.6*inch], styles))

        # PAGE 4
        story.append(PageBreak())
        story.append(Paragraph("Risk Heatmap", styles["OSIL_Section"]))
        story.append(Paragraph("Service-level risk intensity. Red zones require immediate intervention.", styles["OSIL_Body"]))
        story.append(Spacer(1, 12))
        
        hm_img = _build_heatmap_image(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=5.5*inch, height=4.0*inch))

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
