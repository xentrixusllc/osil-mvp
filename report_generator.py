"""OSIL Executive PDF Report Generator - Executive Grade"""
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
    # Fix common spacing issues
    s = s.replace("™", "&trade;")
    return s

def _styles():
    styles = getSampleStyleSheet()
    
    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=24,
            leading=28,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=8,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=11,
            leading=13,
            textColor=colors.HexColor("#555555"),
            spaceAfter=16,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=16,
            leading=20,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=12,
            spaceBefore=16,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_SubSection",
            parent=styles["Heading3"],
            fontName="Helvetica-Bold",
            fontSize=12,
            leading=15,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=8,
            spaceBefore=10,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=14,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#666666"),
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
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Metric",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=20,
            leading=24,
            textColor=colors.HexColor("#0A192F"),
            alignment=1,  # Center
        )
    )
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(0.6 * inch, 0.4 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence • Confidential")
    canvas.drawRightString(7.9 * inch, 0.4 * inch, f"Page {doc.page}")
    canvas.restoreState()

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
        ax.plot(angles_loop, values_loop, linewidth=2.5, color='#0A192F')
        ax.fill(angles_loop, values_loop, alpha=0.15, color='#64FFDA')
        ax.set_xticks(angles)
        ax.set_xticklabels(labels, fontsize=8)
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
        ax.set_title("Operational Stability Profile", pad=20, fontsize=11, fontweight='bold')
        
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
        fig, ax = plt.subplots(figsize=(6, height), dpi=120)
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='RdYlGn_r')
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(list(hm.columns), fontsize=9)
        ax.set_yticklabels(list(hm.index), fontsize=8)
        ax.set_title("Service Stability Risk Heatmap", fontsize=11, fontweight='bold', pad=12)
        
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

def _make_table(data, col_widths, styles, use_paragraphs=True):
    """Create professional table with proper text wrapping"""
    if not data or len(data) == 0:
        t = Table([["No data"]], colWidths=[6*inch])
        t.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))
        return t
    
    # Ensure uniform row length
    max_cols = max(len(row) for row in data)
    norm_data = []
    for i, row in enumerate(data):
        new_row = []
        for j, cell in enumerate(row):
            if use_paragraphs and i > 0:  # Data rows use paragraphs for wrapping
                new_row.append(Paragraph(str(cell), styles["OSIL_TableCell"]))
            elif i == 0:  # Header row
                new_row.append(Paragraph(str(cell), styles["OSIL_TableHeader"]))
            else:
                new_row.append(str(cell))
        # Pad if needed
        while len(new_row) < max_cols:
            new_row.append("" if i > 0 else Paragraph("", styles["OSIL_TableHeader"]) if i == 0 else "")
        norm_data.append(new_row)
    
    if not col_widths or len(col_widths) != max_cols:
        col_widths = [6.0/max_cols*inch] * max_cols
    
    t = Table(norm_data, colWidths=col_widths, repeatRows=1)
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0A192F")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#CCCCCC")),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]
    
    # Alternating row colors for data
    if len(norm_data) > 1:
        for i in range(1, len(norm_data)):
            if i % 2 == 0:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")))
    
    t.setStyle(TableStyle(style_commands))
    return t

def _metric_box(label, value, width):
    """Create metric display box"""
    data = [[Paragraph(label, ParagraphStyle(name='lbl', fontName='Helvetica', fontSize=10, textColor=colors.HexColor("#666666"), alignment=1))],
            [Paragraph(value, ParagraphStyle(name='val', fontName='Helvetica-Bold', fontSize=22, textColor=colors.HexColor("#0A192F"), alignment=1))]]
    t = Table(data, colWidths=[width], rowHeights=[25, 45])
    t.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F5F7FA")),
        ('BOX', (0, 0), (-1, -1), 1.5, colors.HexColor("#0A192F")),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('TOPPADDING', (0, 0), (-1, 0), 8),
        ('BOTTOMPADDING', (0, 1), (-1, 1), 8),
    ]))
    return t

def _posture_text(bvsi, posture):
    if bvsi >= 85:
        return f"<b>Assessment:</b> Operational stability is currently <b>{posture}</b> (BVSI™ {bvsi:.1f}). Technology operations are broadly supporting business confidence."
    elif bvsi >= 70:
        return f"<b>Assessment:</b> Operational stability is currently <b>{posture}</b> (BVSI™ {bvsi:.1f}). Control exists across most dimensions, though localized weaknesses remain."
    elif bvsi >= 55:
        return f"<b>Assessment:</b> Operational stability is currently <b>{posture}</b> (BVSI™ {bvsi:.1f}). Governance mechanisms function, but recurring instability patterns persist."
    elif bvsi >= 40:
        return f"<b>Assessment:</b> Operational stability is currently <b>{posture}</b> (BVSI™ {bvsi:.1f}). Operational exposure is visible and affecting service confidence."
    else:
        return f"<b>Assessment:</b> Operational stability is currently <b>{posture}</b> (BVSI™ {bvsi:.1f}). Immediate executive attention is warranted."

def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    try:
        styles = _styles()
        out = io.BytesIO()
        
        doc = SimpleDocTemplate(
            out,
            pagesize=LETTER,
            rightMargin=0.7 * inch,
            leftMargin=0.7 * inch,
            topMargin=0.6 * inch,
            bottomMargin=0.7 * inch,
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

        # PAGE 1 - Executive Summary
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report — {tenant_name}", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"<i>Report Date: {as_of}</i>", styles["OSIL_Small"]))
        story.append(Spacer(1, 24))
        
        # Metrics Row
        metric_row = Table(
            [[_metric_box("BVSI™ Score", f"{bvsi:.1f}", 2.0*inch),
              _metric_box("Operating Posture", posture, 2.0*inch),
              _metric_box("Data Readiness", f"{readiness:.1f}%", 2.0*inch)]],
            colWidths=[2.2*inch, 2.2*inch, 2.2*inch]
        )
        metric_row.setStyle(TableStyle([('ALIGN', (0, 0), (-1, -1), 'CENTER'), ('VALIGN', (0, 0), (-1, -1), 'TOP')]))
        story.append(metric_row)
        story.append(Spacer(1, 24))
        
        # BVSI Scale
        story.append(Paragraph("BVSI™ Scale Reference", styles["OSIL_Section"]))
        scale_data = [
            ["BVSI™ Range", "Operating Condition", "Executive Interpretation"],
            ["85 — 100", "High Confidence Operations", "Technology stability fully supports business growth and executive confidence"],
            ["70 — 84", "Controlled and Improving", "Core controls exist; targeted improvements can increase resilience"],
            ["55 — 69", "Controlled but Exposed", "Governance functions, but recurring instability creates material exposure"],
            ["40 — 54", "Reactive and Exposed", "Instability is visible and impacting reliability, cost, or customer experience"],
            ["Below 40", "Fragile Operations", "Operational fragility is high; stabilization is an executive priority"],
        ]
        story.append(_make_table(scale_data, [1.2*inch, 2.2*inch, 3.8*inch], styles))
        story.append(Spacer(1, 20))
        
        # Executive Assessment
        story.append(Paragraph("Executive Assessment", styles["OSIL_Section"]))
        story.append(Paragraph(_posture_text(bvsi, posture), styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        # Executive Summary
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(exec_summary, styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        # Key Takeaways
        story.append(Paragraph("Strategic Takeaways", styles["OSIL_Section"]))
        strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        top_svc = "N/A"
        if not sip_candidates.empty:
            top_svc = str(sip_candidates.iloc[0].get("Service", "N/A"))
        
        takeaway_data = [
            ["Area", "Insight", "Recommended Action"],
            ["Strength", f"{strongest} domain", "Leverage as model for other service areas"],
            ["Priority Gap", f"{weakest} domain", "Immediate executive focus and resource allocation"],
            ["Top Initiative", f"{top_svc}", "Initiate Stability Improvement Program (SIP)"],
        ]
        story.append(_make_table(takeaway_data, [1.3*inch, 2.4*inch, 3.5*inch], styles))
        story.append(Spacer(1, 12))
        
        story.append(Paragraph(f"<i>Data Sources: {detected_dataset} | Service Classification: {service_anchor} | Organization: {tenant_name}</i>", 
                              ParagraphStyle(name='footer', fontName='Helvetica', fontSize=8, textColor=colors.HexColor("#999999"))))

        # PAGE 2 - Stability Profile
        story.append(PageBreak())
        story.append(Paragraph("Operational Stability Profile", styles["OSIL_Section"]))
        story.append(Paragraph("The radar chart below displays current performance across four critical stability domains. A balanced shape indicates aligned governance; collapsed areas highlight maturity gaps requiring attention.", 
                              styles["OSIL_Body"]))
        story.append(Spacer(1, 12))
        
        radar_img = _build_radar_image(domain_scores)
        if radar_img:
            story.append(Image(radar_img, width=4.2*inch, height=3.4*inch))
        else:
            story.append(Paragraph("Chart unavailable", styles["OSIL_Small"]))
        
        story.append(Spacer(1, 20))
        story.append(Paragraph("Domain Score Definitions", styles["OSIL_SubSection"]))
        
        domain_data = [
            ["Stability Domain", "Score", "Business Definition"],
            ["Service Resilience", f"{domain_scores.get('Service Resilience', 0):.1f}", 
             "Ability to recover quickly and consistently from incidents. Measures MTTR effectiveness and restoration capability."],
            ["Change Governance", f"{domain_scores.get('Change Governance', 0):.1f}", 
             "Effectiveness of controls preventing operational instability during change activity. Measures change collision and failure rates."],
            ["Structural Risk Debt™", f"{domain_scores.get('Structural Risk Debt™', 0):.1f}", 
             "Accumulated instability from unresolved recurring weaknesses. Measures problem management maturity and recurrence patterns."],
            ["Reliability Momentum", f"{domain_scores.get('Reliability Momentum', 0):.1f}", 
             "Direction of operational reliability based on trending stability patterns. Forward-looking indicator of operational trajectory."],
        ]
        story.append(_make_table(domain_data, [1.9*inch, 0.8*inch, 4.5*inch], styles))
        story.append(Spacer(1, 8))
        story.append(Paragraph("<i>Score Guide: 80-100 = Strong | 60-79 = Controlled | 40-59 = At Risk | Below 40 = Critical</i>", styles["OSIL_Small"]))

        # PAGE 3 - SIP Priorities
        story.append(PageBreak())
        story.append(Paragraph("Service Improvement Priorities", styles["OSIL_Section"]))
        story.append(Paragraph("The following initiatives represent the highest-impact stability improvements for the next 30-60 days. These should guide leadership briefing, ownership alignment, and targeted operational remediation.", 
                              styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        if not sip_candidates.empty:
            story.append(Paragraph("Top 3 Initiatives for Executive Attention", styles["OSIL_SubSection"]))
            
            top3_data = [["Service", "Tier", "Focus Area", "Priority", "Business Rationale"]]
            for _, row in sip_candidates.head(3).iterrows():
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "Unknown"))
                theme = str(row.get("Suggested_Theme", "Stability"))
                priority = str(row.get("Priority_Label", "Monitor"))
                why = str(row.get("Why_Flagged", "Risk identified"))
                top3_data.append([svc, tier, theme, priority, why])
            
            story.append(_make_table(top3_data, [1.6*inch, 0.7*inch, 1.3*inch, 0.9*inch, 2.7*inch], styles))
            story.append(Spacer(1, 24))
            
            story.append(Paragraph("Complete SIP Candidate Registry", styles["OSIL_SubSection"]))
            
            display_cols = ["Service", "Service_Tier", "SIP_Priority_Score", "Priority_Label"]
            available_cols = [c for c in display_cols if c in sip_candidates.columns]
            if available_cols:
                all_data = [["Service", "Tier", "Priority Score", "Classification"]]
                for _, row in sip_candidates[available_cols].head(10).iterrows():
                    all_data.append([
                        str(row.get("Service", "")),
                        str(row.get("Service_Tier", "")),
                        f"{float(row.get('SIP_Priority_Score', 0)):.1f}",
                        str(row.get("Priority_Label", ""))
                    ])
                
                col_widths = [2.5*inch, 1.0*inch, 1.5*inch, 1.5*inch]
                story.append(_make_table(all_data, col_widths, styles))
        else:
            story.append(Paragraph("No SIP candidates identified in current dataset.", styles["OSIL_Body"]))

        # PAGE 4 - Heatmap
        story.append(PageBreak())
        story.append(Paragraph("Service Stability Risk Heatmap", styles["OSIL_Section"]))
        story.append(Paragraph("This heatmap visualizes stability risk intensity across your top services. <b>Red indicates higher risk</b> requiring immediate attention; <b>green indicates stable operations</b>. Services with multiple red columns are priority candidates for executive intervention.", 
                              styles["OSIL_Body"]))
        story.append(Spacer(1, 16))
        
        hm_img = _build_heatmap_image(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=5.5*inch, height=4.0*inch))
        else:
            story.append(Paragraph("Insufficient data for heatmap generation.", styles["OSIL_Body"]))

        # Build PDF
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
