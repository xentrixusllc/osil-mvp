"""OSIL Executive PDF Report Generator - Simplified Robust Version"""
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
    allowed = {"b", "i", "u", "br", "font"}
    
    def repl_tag(match):
        full = match.group(0)
        tag = match.group(1).lower().replace("/", "")
        return full if tag in allowed else ""
    
    s = re.sub(r"</?([a-zA-Z0-9]+).*?>", repl_tag, s)
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
            fontSize=13,
            leading=16,
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

def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    labels = list(domain_scores.keys())
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]
    
    if not labels:
        labels = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
        values = [0, 0, 0, 0]
    
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]
    
    fig = plt.figure(figsize=(5, 4), dpi=150)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.plot(angles_loop, values_loop, linewidth=2, color='#0A192F')
    ax.fill(angles_loop, values_loop, alpha=0.15, color='#64FFDA')
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
    ax.set_title("Operational Stability Profile", pad=15, fontsize=11, fontweight='bold')
    
    plt.tight_layout()
    img = io.BytesIO()
    plt.savefig(img, format="png", bbox_inches="tight", facecolor='white')
    plt.close(fig)
    img.seek(0)
    return img

def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> Optional[io.BytesIO]:
    df = _safe_df(service_risk_top10)
    if df.empty:
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
        hm["Display_Name"] = hm["Service"].astype(str).str[:25] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        
        display_names = {
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision"
        }
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        height = max(4, len(hm) * 0.35)
        fig = plt.figure(figsize=(7, height), dpi=150)
        ax = plt.gca()
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='RdYlGn_r')
        ax.set_xticks(range(len(hm.columns)))
        ax.set_xticklabels(list(hm.columns), fontsize=9)
        ax.set_yticks(range(len(hm.index)))
        ax.set_yticklabels(list(hm.index), fontsize=9)
        ax.set_title("Service Stability Risk Heatmap", fontsize=11, fontweight='bold', pad=15)
        
        for i in range(hm.shape[0]):
            for j in range(hm.shape[1]):
                val = int(round(float(hm.iat[i, j]), 0))
                text_color = "white" if val > 60 else "black"
                ax.text(j, i, str(val), ha="center", va="center", fontsize=8, color=text_color, fontweight='bold')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label("Risk Score", fontsize=9)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap error: {e}")
        return None

def _posture_text(bvsi: float, posture: str) -> str:
    if bvsi >= 85:
        return f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Technology operations are broadly supporting business confidence."
    if bvsi >= 70:
        return f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Control exists across most operational dimensions, though localized weaknesses remain."
    if bvsi >= 55:
        return f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Governance mechanisms are functioning, but recurring instability patterns remain."
    if bvsi >= 40:
        return f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Operational exposure is visible and likely affecting service confidence."
    return f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Immediate executive attention is warranted."

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

        # PAGE 1 - Title and Metrics
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report - {tenant_name}", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"As of: {as_of}", styles["OSIL_Small"]))
        story.append(Spacer(1, 20))
        
        # Simple metrics display
        metrics_data = [
            ["BVSI™ Score", "Operating Posture", "Data Readiness"],
            [f"{bvsi:.1f}", posture, f"{readiness:.1f}%"]
        ]
        metrics_table = Table(metrics_data, colWidths=[2.3*inch, 2.3*inch, 2.3*inch])
        metrics_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0A192F")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 8),
            ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#F5F7FA")),
            ('FONTNAME', (0, 1), (-1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (-1, 1), 18),
            ('TEXTCOLOR', (0, 1), (-1, 1), colors.HexColor("#0A192F")),
            ('TOPPADDING', (0, 1), (-1, 1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 10),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#D1D5DB")),
        ]))
        story.append(metrics_table)
        story.append(Spacer(1, 20))
        
        # BVSI Scale - Simple format
        story.append(Paragraph("BVSI™ Scale Reference", styles["OSIL_Section"]))
        scale_text = """
        <b>85-100:</b> High Confidence Operations - Technology stability supports business growth<br/>
        <b>70-84:</b> Controlled and Improving - Core controls exist; targeted improvement needed<br/>
        <b>55-69:</b> Controlled but Exposed - Governance functions, but recurring instability remains<br/>
        <b>40-54:</b> Reactive and Exposed - Instability visible and impacting operations<br/>
        <b>&lt;40:</b> Fragile Operations - Immediate executive attention required
        """
        story.append(Paragraph(scale_text, styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        # Executive Signal
        story.append(Paragraph("Executive Assessment", styles["OSIL_Section"]))
        story.append(Paragraph(_posture_text(bvsi, posture), styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        # Executive Summary
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(exec_summary, styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        # Key Takeaways
        story.append(Paragraph("Key Takeaways", styles["OSIL_Section"]))
        strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "N/A"
        top_svc = "N/A"
        if not sip_candidates.empty:
            top_svc = str(sip_candidates.iloc[0].get("Service", "N/A"))
        
        takeaways = f"""
        • <b>Strength:</b> {strongest} shows the strongest stability performance<br/>
        • <b>Priority:</b> Address {weakest} for immediate improvement<br/>
        • <b>Action Item:</b> Focus on {top_svc} as the leading service candidate
        """
        story.append(Paragraph(takeaways, styles["OSIL_Body"]))
        story.append(Spacer(1, 10))
        
        # Footer info
        story.append(Paragraph(f"<i>Data: {detected_dataset} | Anchor: {service_anchor} | Org: {tenant_name}</i>", styles["OSIL_Small"]))

        # PAGE 2 - Stability Profile
        story.append(PageBreak())
        story.append(Paragraph("Operational Stability Profile", styles["OSIL_Section"]))
        story.append(Paragraph("Performance across four critical stability domains. Balanced shape indicates aligned governance.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))
        
        try:
            radar_img = _build_radar_image(domain_scores)
            story.append(Image(radar_img, width=4.5*inch, height=3.8*inch))
        except Exception as e:
            story.append(Paragraph(f"Chart error: {e}", styles["OSIL_Small"]))
        
        story.append(Spacer(1, 15))
        story.append(Paragraph("Domain Scores", styles["OSIL_Section"]))
        
        domain_data = [
            ["Domain", "Score", "Definition"],
            ["Service Resilience", f"{domain_scores.get('Service Resilience', 0):.1f}", "Ability to recover quickly from incidents"],
            ["Change Governance", f"{domain_scores.get('Change Governance', 0):.1f}", "Effectiveness of change controls"],
            ["Structural Risk Debt", f"{domain_scores.get('Structural Risk Debt™', 0):.1f}", "Accumulated unresolved weaknesses"],
            ["Reliability Momentum", f"{domain_scores.get('Reliability Momentum', 0):.1f}", "Direction of operational reliability"],
        ]
        
        domain_table = Table(domain_data, colWidths=[2.0*inch, 0.8*inch, 4.2*inch])
        domain_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E9EDF3")),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
            ('TOPPADDING', (0, 0), (-1, -1), 6),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
            ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ]))
        story.append(domain_table)

        # PAGE 3 - SIP Priorities
        story.append(PageBreak())
        story.append(Paragraph("Service Improvement Priorities", styles["OSIL_Section"]))
        story.append(Paragraph("Highest-impact stability improvements for the next 30-60 days.", styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        if not sip_candidates.empty:
            # Top 3 summary
            story.append(Paragraph("Top 3 Initiatives", styles["OSIL_Section"]))
            for idx, row in sip_candidates.head(3).iterrows():
                svc = str(row.get("Service", "Unknown"))
                tier = str(row.get("Service_Tier", "Unknown"))
                theme = str(row.get("Suggested_Theme", "Stability"))
                why = str(row.get("Why_Flagged", "Risk identified"))
                story.append(Paragraph(f"<b>{svc} ({tier}) - {theme}</b><br/>{why}", styles["OSIL_Body"]))
                story.append(Spacer(1, 8))
            
            story.append(Spacer(1, 10))
            story.append(Paragraph("All Candidates", styles["OSIL_Section"]))
            
            # Detailed table - keep it simple
            display_cols = ["Service", "Service_Tier", "SIP_Priority_Score", "Priority_Label"]
            available_cols = [c for c in display_cols if c in sip_candidates.columns]
            if available_cols:
                sip_display = sip_candidates[available_cols].head(10)
                sip_data = [available_cols] + sip_display.values.tolist()
                
                col_widths = [7.0/len(available_cols)*inch] * len(available_cols)
                sip_table = Table(sip_data, colWidths=col_widths)
                sip_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#E9EDF3")),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, -1), 9),
                    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                    ('TOPPADDING', (0, 0), (-1, -1), 5),
                    ('BOTTOMPADDING', (0, 0), (-1, -1), 5),
                    ('LEFTPADDING', (0, 0), (-1, -1), 6),
                ]))
                story.append(sip_table)
        else:
            story.append(Paragraph("No SIP candidates identified.", styles["OSIL_Body"]))

        # PAGE 4 - Heatmap
        story.append(PageBreak())
        story.append(Paragraph("Service Stability Heatmap", styles["OSIL_Section"]))
        story.append(Paragraph("Risk visualization across top services. Red indicates higher risk.", styles["OSIL_Body"]))
        story.append(Spacer(1, 15))
        
        hm_img = _build_heatmap_image(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=6.5*inch))
        else:
            story.append(Paragraph("Insufficient data for heatmap generation.", styles["OSIL_Body"]))

        # Build PDF
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        raise Exception(f"PDF Generation Failed: {str(e)}")
