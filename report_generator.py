"""OSIL Executive PDF Report Generator - Working Version"""
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
    return s

def _styles():
    styles = getSampleStyleSheet()
    
    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=6,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=12,
            textColor=colors.HexColor("#555555"),
            spaceAfter=12,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=14,
            leading=17,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=10,
            spaceBefore=12,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=10,
            leading=13,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Small",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#666666"),
        )
    )
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(0.6 * inch, 0.4 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
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
        ax.set_xticklabels(labels, fontsize=9)
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)
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
        hm["Display_Name"] = hm["Service"].astype(str).str[:18] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        
        display_names = {
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision"
        }
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        # Smaller figure size for PDF
        height = min(6, max(4, len(hm) * 0.4))
        fig, ax = plt.subplots(figsize=(6, height), dpi=120)
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='RdYlGn_r')
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(list(hm.columns), fontsize=9)
        ax.set_yticklabels(list(hm.index), fontsize=9)
        ax.set_title("Service Stability Risk Heatmap", fontsize=11, fontweight='bold', pad=12)
        
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(float(hm.iloc[i, j]), 0))
                text_color = "white" if val > 60 else "black"
                ax.text(j, i, str(val), ha="center", va="center", 
                       fontsize=8, color=text_color, fontweight='bold')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label("Risk Score", fontsize=9)
        cbar.ax.tick_params(labelsize=8)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white', edgecolor='none', pad_inches=0.15)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap error: {e}")
        return None

def _create_table(data: List[List[str]], col_widths: List[float], header: bool = True) -> Table:
    """Create table with validation"""
    if not data or len(data) == 0:
        t = Table([["No data available"]], colWidths=[sum(col_widths) if col_widths else 6*inch])
        t.setStyle(TableStyle([
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 0), (-1, -1), 10),
        ]))
        return t
    
    max_cols = max(len(row) for row in data)
    normalized_data = []
    for row in data:
        if len(row) < max_cols:
            row = row + [""] * (max_cols - len(row))
        normalized_data.append([str(cell) for cell in row])
    
    if not col_widths or len(col_widths) != max_cols:
        col_widths = [6.0/max_cols*inch] * max_cols
    
    t = Table(normalized_data, colWidths=col_widths, repeatRows=1 if header else 0)
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0A192F")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
    ]
    
    if header and len(normalized_data) > 1:
        for i in range(1, len(normalized_data)):
            if i % 2 == 0:
                style_commands.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F5F7FA")))
    
    t.setStyle(TableStyle(style_commands))
    return t

def _posture_text(bvsi: float, posture: str) -> str:
    if bvsi >= 85:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Technology operations are broadly supporting business confidence."
    if bvsi >= 70:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Control exists across most operational dimensions, though localized weaknesses remain."
    if bvsi >= 55:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Governance mechanisms are functioning, but recurring instability patterns remain."
    if bvsi >= 40:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Operational exposure is visible and likely affecting service confidence."
    return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Immediate executive attention is warranted."

def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    try:
        styles = _styles()
        out = io.BytesIO()
        
        doc = SimpleDocTemplate(
            out,
            pagesize=LETTER,
            rightMargin=0.6 * inch,
            leftMargin=0.6 * inch,
            topMargin=0.5 * inch,
            bottomMargin=0.6 * inch,
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

        # PAGE 1
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report - {tenant_name}", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"As of: {as_of}", styles["OSIL_Small"]))
        story.append(Spacer(1, 20))
        
        metrics_data = [
            ["BVSI™ Score", "Operating Posture", "Data Readiness"],
            [f"{bvsi:.1f}", posture, f"{readiness:.1f}%"]
        ]
        story.append(_create_table(metrics_data, [2.3*inch, 2.3*inch, 2.3*inch]))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("BVSI™ Scale Reference", styles["OSIL_Section"]))
        scale_data = [
            ["BVSI™ Range", "Operating Condition", "Executive Meaning"],
            ["85-100", "High Confidence Operations", "Technology stability supports business growth and executive confidence."],
            ["70-84", "Controlled and Improving", "Core controls exist; targeted improvement can increase resilience."],
            ["55-69", "Controlled but Exposed", "Operational control exists, but recurring instability creates material exposure."],
            ["40-54", "Reactive and Exposed", "Instability is visible and may be impacting reliability, cost, or customer experience."],
            ["<40", "Fragile Operations", "Operational fragility is high; stabilization is an executive priority."],
        ]
        story.append(_create_table(scale_data, [1.1*inch, 2.0*inch, 4.0*inch]))
        story.append(Spacer(1, 15))
        
        story.append(Paragraph("Executive Assessment", styles["OSIL_Section"]))
        story.append(Paragraph(_posture_text(bvsi, posture), styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(exec_summary, styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        story.append(Paragraph("Key Takeaways", styles["OSIL_Section"]))
        strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        top_svc = "N/A"
        if not sip_candidates.empty:
            top_svc = str(sip_candidates.iloc[0].get("Service", "N/A"))
        
        takeaway_data = [
            ["Key Insight", "Details"],
            ["Strength", f"{strongest} shows the strongest stability performance"],
            ["Priority", f"Address {weakest} for immediate improvement"],
            ["Action Item", f"Focus on {top_svc} as the leading service candidate"],
        ]
        story.append(_create_table(takeaway_data, [1.5*inch, 5.5*inch]))
        story.append(Spacer(1, 10))
        
        story.append(Paragraph(f"<i>Data: {detected_dataset} | Anchor: {service_anchor} | Org: {tenant_name}</i>", styles["OSIL_Small"]))

        # PAGE 2
        story.append(PageBreak())
        story.append(Paragraph("Operational Stability Profile", styles["OSIL_Section"]))
        story.append(Paragraph("Performance across four critical stability domains. Balanced shape indicates aligned governance.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))
        
        radar_img = _build_radar_image(domain_scores)
        if radar_img:
            story.append(Image(radar_img, width=4.5*inch, height=3.6*inch))
        else:
            story.append(Paragraph("Unable to generate radar chart.", styles["OSIL_Small"]))
        
        story.append(Spacer(1, 15))
        story.append(Paragraph("Domain Score Definitions", styles["OSIL_Section"]))
        
        domain_data = [
            ["Domain", "Score", "Definition"],
            ["Service Resilience", f"{domain_scores.get('Service Resilience', 0):.1f}", "Ability to recover quickly and consistently from incidents."],
            ["Change Governance", f"{domain_scores.get('Change Governance', 0):.1f}", "Effectiveness of controls preventing operational instability during change activity."],
            ["Structural Risk Debt™", f"{domain_scores.get('Structural Risk Debt™', 0):.1f}", "Accumulated instability caused by unresolved recurring weaknesses and operational debt."],
            ["Reliability Momentum", f"{domain_scores.get('Reliability Momentum', 0):.1f}", "Direction of operational reliability based on recurring instability and recovery behavior."],
        ]
        story.append(_create_table(domain_data, [2.0*inch, 0.8*inch, 4.2*inch]))

        # PAGE 3
        story.append(PageBreak())
        story.append(Paragraph("Service Improvement Priorities", styles["OSIL_Section"]))
        story.append(Paragraph("Highest-impact stability improvements for the next 30-60 days.", styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        if not sip_candidates.empty:
            story.append(Paragraph("Top 3 Initiatives for Leadership", styles["OSIL_Section"]))
            
            top3_data = [["Service", "Tier", "Theme", "Priority", "Rationale"]]
            for _, row in sip_candidates.head(3).iterrows():
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "Unknown"))
                theme = str(row.get("Suggested_Theme", "Stability"))
                priority = str(row.get("Priority_Label", "Monitor"))
                why = str(row.get("Why_Flagged", "Risk identified"))
                top3_data.append([svc, tier, theme, priority, why])
            
            story.append(_create_table(top3_data, [1.8*inch, 0.8*inch, 1.4*inch, 0.9*inch, 2.1*inch]))
            story.append(Spacer(1, 20))
            
            story.append(Paragraph("All SIP Candidates", styles["OSIL_Section"]))
            
            display_cols = ["Service", "Service_Tier", "SIP_Priority_Score", "Priority_Label"]
            available_cols = [c for c in display_cols if c in sip_candidates.columns]
            if available_cols:
                all_data = [available_cols]
                for _, row in sip_candidates[available_cols].head(10).iterrows():
                    all_data.append([str(row.get(c, "")) for c in available_cols])
                
                col_widths = [7.0/len(available_cols)*inch] * len(available_cols)
                story.append(_create_table(all_data, col_widths))
        else:
            story.append(Paragraph("No SIP candidates identified.", styles["OSIL_Body"]))

        # PAGE 4
        story.append(PageBreak())
        story.append(Paragraph("Service Stability Heatmap", styles["OSIL_Section"]))
        story.append(Paragraph("Risk visualization across top services. Red indicates higher risk, green indicates lower risk.", styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        hm_img = _build_heatmap_image(service_risk_df)
        if hm_img:
            # FIXED: Explicit dimensions to fit page
            img = Image(hm_img, width=5.0*inch, height=4.5*inch)
            story.append(img)
        else:
            story.append(Paragraph("Insufficient data for heatmap generation.", styles["OSIL_Body"]))

        # Build PDF
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
