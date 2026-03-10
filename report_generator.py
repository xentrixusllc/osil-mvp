"""OSIL Executive PDF Report Generator - Layout and Heatmap Fixed"""
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

# Configure matplotlib to avoid LaTeX and rendering issues
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
        name="OSIL_MainTitle",
        parent=styles["Title"],
        fontName="Helvetica-Bold",
        fontSize=22,
        leading=26,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=4,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_ReportSubtitle",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=11,
        leading=13,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=6,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_Date",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=12,
        textColor=colors.HexColor("#64748B"),
        spaceAfter=16,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_SectionHeader",
        parent=styles["Heading2"],
        fontName="Helvetica-Bold",
        fontSize=14,
        leading=18,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=10,
        spaceBefore=12,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_PageHeader",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=16,
        leading=20,
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=12,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_Body",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=10,
        leading=14,
        textColor=colors.HexColor("#334155"),
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_Small",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=12,
        textColor=colors.HexColor("#64748B"),
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_TableHeader",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.white,
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_TableCell",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#334155"),
    ))
    
    styles.add(ParagraphStyle(
        name="OSIL_TableCellBold",
        parent=styles["BodyText"],
        fontName="Helvetica-Bold",
        fontSize=9,
        leading=11,
        textColor=colors.HexColor("#0F172A"),
    ))
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#94A3B8"))
    canvas.drawString(0.75 * inch, 0.5 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
    canvas.drawRightString(7.75 * inch, 0.5 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _build_radar_image(domain_scores: Dict[str, float]) -> Optional[io.BytesIO]:
    try:
        label_map = {
            "Service Resilience": "Service\nResilience",
            "Change Governance": "Change\nGovernance", 
            "Structural Risk Debt": "Structural\nRisk Debt",
            "Structural Risk Debt™": "Structural\nRisk Debt",
            "Reliability Momentum": "Reliability\nMomentum"
        }
        
        default_domains = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
        values = []
        labels = []
        
        for domain in default_domains:
            if domain in domain_scores:
                values.append(_safe_float(domain_scores[domain], 0.0))
            elif domain + "™" in domain_scores:
                values.append(_safe_float(domain_scores[domain + "™"], 0.0))
            else:
                values.append(0.0)
            labels.append(label_map.get(domain, domain))
        
        angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
        angles_loop = angles + [angles[0]]
        values_loop = values + [values[0]]
        
        fig = plt.figure(figsize=(6, 5.5))
        ax = fig.add_subplot(111, polar=True)
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        ax.fill(angles_loop, values_loop, color='#3B82F6', alpha=0.25)
        ax.plot(angles_loop, values_loop, color='#1E40AF', linewidth=2.5)
        
        ax.grid(True, linestyle='--', alpha=0.6, color='gray')
        ax.set_axisbelow(True)
        
        ax.set_xticks(angles)
        ax.set_xticklabels(labels, fontsize=10, fontweight='bold')
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=9, color='gray')
        
        ax.set_title("Operational Stability Radar", fontsize=14, fontweight='bold', 
                    color='#0F172A', pad=20, y=1.08)
        
        plt.tight_layout(pad=0.5)
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.2)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Radar chart error: {e}")
        return None

def _build_heatmap_image(service_risk_df: pd.DataFrame) -> Optional[io.BytesIO]:
    df = _safe_df(service_risk_df)
    if df.empty or len(df) == 0:
        return None
    
    required_cols = ["Service", "Service_Tier"]
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_names = {
        "Recurrence_Risk": "Recurrence",
        "MTTR_Drag_Risk": "MTTR Drag", 
        "Reopen_Churn_Risk": "Reopen",
        "Change_Collision_Risk": "Change"
    }
    
    if not all(c in df.columns for c in required_cols):
        return None
    
    available_risks = [c for c in risk_cols if c in df.columns]
    if not available_risks:
        return None
    
    try:
        hm = df.head(10).copy()
        hm["Display_Name"] = hm["Service"].astype(str).str[:20] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        fig, ax = plt.subplots(figsize=(7, 5))
        
        cmap = plt.cm.RdYlGn_r
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap=cmap)
        
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(list(hm.columns), fontsize=10, fontweight='bold')
        ax.set_yticklabels(list(hm.index), fontsize=9)
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label("Risk Score", fontsize=10, fontweight='bold')
        cbar.ax.tick_params(labelsize=9)
        
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(float(hm.iloc[i, j]), 0))
                text_color = "white" if val > 60 else "black"
                ax.text(j, i, str(val), ha="center", va="center", 
                       fontsize=9, color=text_color, fontweight='bold')
        
        ax.set_title("Service Stability Heatmap", fontsize=13, fontweight='bold', 
                    color='#0F172A', pad=15)
        
        plt.tight_layout(pad=0.5)
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.2)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap error: {e}")
        return None

def _create_table(data, col_widths, header_style, cell_style, header_bg="#0F172A"):
    if not data or len(data) == 0:
        return Table([["No data available"]])
    
    table_data = []
    for i, row in enumerate(data):
        new_row = []
        for cell in row:
            if i == 0:
                new_row.append(Paragraph(str(cell), header_style))
            else:
                new_row.append(Paragraph(str(cell), cell_style))
        table_data.append(new_row)
    
    t = Table(table_data, colWidths=col_widths, repeatRows=1)
    
    style_commands = [
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor(header_bg)),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 9),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('LEFTPADDING', (0, 0), (-1, -1), 8),
        ('RIGHTPADDING', (0, 0), (-1, -1), 8),
        ('TOPPADDING', (0, 0), (-1, -1), 6),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 6),
        ('LINEABOVE', (0, 0), (-1, 0), 1.5, colors.HexColor(header_bg)),
        ('LINEBELOW', (0, 0), (-1, 0), 1.5, colors.HexColor(header_bg)),
    ]
    
    for i in range(1, len(table_data)):
        if i % 2 == 0:
            style_commands.append(('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")))
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
            rightMargin=0.7 * inch,
            leftMargin=0.7 * inch,
            topMargin=0.7 * inch,
            bottomMargin=0.7 * inch,
        )

        tenant_name = str(payload.get("tenant_name", "Default"))
        as_of = str(payload.get("as_of", "2026-03-28"))
        bvsi = _safe_float(payload.get("bvsi", 69.6))
        posture = str(payload.get("posture", _get_posture_from_bvsi(bvsi)))
        exec_summary = _clean_text(payload.get("executive_interpretation", ""))
        domain_scores = payload.get("domain_scores", {}) or {}
        service_risk_df = _safe_df(payload.get("service_risk_top10"))
        sip_candidates = _safe_df(payload.get("sip_candidates"))
        detected_dataset = str(payload.get("detected_dataset", "INCIDENT + CHANGE + PROBLEM")).upper()
        service_anchor = str(payload.get("service_anchor_used", "Service"))
        readiness = _safe_float(payload.get("data_readiness_score", 87.5))
        
        strongest = max(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Change Governance"
        weakest = min(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Structural Risk Debt"
        
        top_svc = "Mobile Banking"
        top_priority = "Next SIP"
        if not sip_candidates.empty:
            top_svc = str(sip_candidates.iloc[0].get("Service", "Mobile Banking"))
            top_priority = str(sip_candidates.iloc[0].get("Priority_Label", "Next SIP"))

        story = []

        # ==================== PAGE 1 CONTENT ====================
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_MainTitle"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report — {tenant_name}", styles["OSIL_ReportSubtitle"]))
        story.append(Paragraph(f"As of {as_of}", styles["OSIL_Date"]))
        
        story.append(Paragraph("Executive Brief", styles["OSIL_SectionHeader"]))
        
        brief_data = [
            ["BVSI™ Score", "Operating Posture", "Data Readiness"],
            [f"<b>{bvsi:.1f}</b>", posture, f"{readiness:.1f}%"]
        ]
        brief_table = _create_table(brief_data, [2.0*inch, 3.0*inch, 1.5*inch], 
                                   styles["OSIL_TableHeader"], styles["OSIL_TableCellBold"])
        story.append(brief_table)
        story.append(Spacer(1, 12))
        
        story.append(Paragraph("BVSI™ Interpretation", styles["OSIL_SectionHeader"]))
        
        interp_data = [
            ["BVSI™ Range", "Operating Condition", "Executive Meaning"],
            ["85–100", "High Confidence Operations", "Technology stability supports business growth and executive confidence."],
            ["70–84", "Controlled and Improving", "Core controls exist; targeted improvement can increase resilience and scale-readiness."],
            ["55–69", "Controlled but Exposed", "Operational control exists, but recurring instability still creates material exposure."],
            ["40–54", "Reactive and Exposed", "Instability is visible and may be impacting reliability, cost, or customer experience."],
            ["<40", "Fragile Operations", "Operational fragility is high and stabilization should be treated as an executive priority."]
        ]
        interp_table = _create_table(interp_data, [1.2*inch, 2.0*inch, 3.3*inch],
                                    styles["OSIL_TableHeader"], styles["OSIL_TableCell"])
        story.append(interp_table)
        story.append(Spacer(1, 10))
        
        story.append(Paragraph("Executive Signal", styles["OSIL_SectionHeader"]))
        signal_text = f"Operational stability is currently {posture}, with a BVSI™ score of {bvsi:.1f}. Governance mechanisms are functioning, but recurring instability patterns remain across higher-impact services."
        story.append(Paragraph(signal_text, styles["OSIL_Body"]))
        story.append(Spacer(1, 8))
        
        story.append(Paragraph("Executive Summary", styles["OSIL_SectionHeader"]))
        if exec_summary:
            story.append(Paragraph(exec_summary, styles["OSIL_Body"]))
        else:
            default_summary = f"Your organization is operating in a {posture} posture (BVSI™ {bvsi:.1f}). Current stability signals suggest the greatest exposure sits in {weakest}. Incident restoration signals show visible drag. Structural learning signals remain inconsistent. Change governance appears steady."
            story.append(Paragraph(default_summary, styles["OSIL_Body"]))
        story.append(Spacer(1, 8))
        
        # Key Takeaways - Keep this section together
        kt_header = Paragraph("Key Takeaways", styles["OSIL_SectionHeader"])
        
        takeaway_data = [
            ["Key Takeaway"],
            [f"Operational strength is most visible in {strongest}, which currently provides the strongest stability support."],
            [f"The primary exposure area is {weakest}, and it should anchor near-term stabilization decisions."],
            [f"{top_svc} is the leading service candidate for action and is currently classified as {top_priority}."]
        ]
        
        takeaway_table = Table(takeaway_data, colWidths=[6.5*inch])
        takeaway_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#F1F5F9")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 9),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('RIGHTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEBELOW', (0, 0), (-1, 0), 1, colors.HexColor("#CBD5E1")),
            ('LINEBELOW', (0, -1), (-1, -1), 1, colors.HexColor("#CBD5E1")),
            ('LINEABOVE', (0, 0), (-1, 0), 1, colors.HexColor("#CBD5E1")),
        ]))
        
        kt_section = [kt_header, Spacer(1, 4), takeaway_table]
        story.append(KeepTogether(kt_section))
        
        story.append(Spacer(1, 12))
        
        meta_text = f"Detected dataset: {detected_dataset} • Service anchor used: {service_anchor} • Organization: {tenant_name}"
        story.append(Paragraph(meta_text, styles["OSIL_Small"]))

        story.append(PageBreak())

        # ==================== PAGE 2 CONTENT ====================
        story.append(Paragraph("Operational Stability Profile", styles["OSIL_PageHeader"]))
        story.append(Spacer(1, 6))
        
        radar_img = _build_radar_image(domain_scores)
        if radar_img:
            story.append(Image(radar_img, width=4.8*inch, height=4.4*inch))
            story.append(Spacer(1, 6))
        
        how_to_read = "How to read this chart: higher scores indicate stronger operational stability. The radar shows current performance across service resilience, change governance, structural risk debt, and reliability momentum."
        story.append(Paragraph(how_to_read, styles["OSIL_Small"]))
        story.append(Spacer(1, 12))
        
        story.append(Paragraph("Domain Scores", styles["OSIL_SectionHeader"]))
        
        domain_data = [["Domain", "Score", "What It Means"]]
        domain_definitions = {
            "Service Resilience": "Ability to recover quickly and consistently from incidents.",
            "Change Governance": "Effectiveness of controls preventing operational instability during change activity.",
            "Structural Risk Debt": "Accumulated instability caused by unresolved recurring weaknesses and operational debt.",
            "Structural Risk Debt™": "Accumulated instability caused by unresolved recurring weaknesses and operational debt.",
            "Reliability Momentum": "Direction of operational reliability based on recurring instability and recovery behavior."
        }
        
        for domain_key in ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]:
            display_name = "Structural Risk Debt™" if domain_key == "Structural Risk Debt" else domain_key
            score = 0.0
            if domain_key in domain_scores:
                score = _safe_float(domain_scores[domain_key], 0.0)
            elif domain_key + "™" in domain_scores:
                score = _safe_float(domain_scores[domain_key + "™"], 0.0)
            definition = domain_definitions.get(domain_key, "")
            domain_data.append([display_name, f"{score:.1f}", definition])
        
        domain_table = _create_table(domain_data, [2.2*inch, 0.7*inch, 3.6*inch],
                                    styles["OSIL_TableHeader"], styles["OSIL_TableCell"])
        story.append(domain_table)
        story.append(Spacer(1, 8))
        
        story.append(Paragraph("Score guide: 80–100 = strong maturity; 60–79 = controlled but improving; 40–59 = operational weakness; below 40 = structural fragility", 
                              styles["OSIL_Small"]))

        story.append(PageBreak())

        # ==================== PAGE 3 CONTENT ====================
        story.append(Paragraph("Service Improvement Priorities", styles["OSIL_PageHeader"]))
        story.append(Paragraph("These represent the highest-impact stability improvements for the next 30–60 days. They should be used to guide leadership briefing, ownership alignment, and targeted operational investment.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))
        
        if not sip_candidates.empty:
            story.append(Paragraph("Top 3 Initiatives to Brief Leadership", styles["OSIL_SectionHeader"]))
            
            initiatives_data = [["Initiative", "Why Leadership Should Care"]]
            for idx, (_, row) in enumerate(sip_candidates.head(3).iterrows(), 1):
                svc = str(row.get("Service", row.get("service", "Unknown")))
                tier = str(row.get("Service_Tier", row.get("service_tier", "Unknown")))
                theme = str(row.get("Suggested_Theme", row.get("suggested_theme", "Stability")))
                why = str(row.get("Why_Flagged", row.get("why_flagged", "Risk exposure")))
                
                initiative_name = f"{svc} ({tier}) — {theme}"
                if idx == 1:
                    initiative_name = f"<b>{svc}</b> ({tier}) — {theme}"
                
                initiatives_data.append([initiative_name, why])
            
            initiatives_table = _create_table(initiatives_data, [3.0*inch, 3.5*inch],
                                             styles["OSIL_TableHeader"], styles["OSIL_TableCell"])
            story.append(initiatives_table)
            story.append(Spacer(1, 12))
            
            story.append(Paragraph("Detailed SIP Candidates", styles["OSIL_SectionHeader"]))
            
            sip_data = [["Service", "Tier", "Theme", "Score", "Priority", "Why Flagged"]]
            for _, row in sip_candidates.iterrows():
                svc = str(row.get("Service", row.get("service", "Unknown")))
                tier = str(row.get("Service_Tier", row.get("service_tier", "Unknown")))
                theme = str(row.get("Suggested_Theme", row.get("suggested_theme", "Stability")))
                priority = str(row.get("Priority_Label", row.get("priority_label", "Monitor")))
                why = str(row.get("Why_Flagged", row.get("why_flagged", "Risk")))
                score = _safe_float(row.get("SIP_Priority_Score", row.get("sip_priority_score", 0.0)), 0.0)
                
                sip_data.append([svc, tier, theme, f"{score:.1f}", priority, why])
            
            sip_table = _create_table(sip_data, [1.5*inch, 0.6*inch, 1.1*inch, 0.7*inch, 0.9*inch, 1.7*inch],
                                     styles["OSIL_TableHeader"], styles["OSIL_TableCell"])
            story.append(sip_table)

        story.append(PageBreak())

        # ==================== PAGE 4 CONTENT (HEATMAP) ====================
        story.append(Paragraph("Service Stability Heatmap", styles["OSIL_PageHeader"]))
        story.append(Paragraph("Executive view: Service × Stability Risk across recurrence, MTTR drag, reopen churn, and change collision. This visual helps leadership identify where concentrated instability is likely to create operational exposure.", styles["OSIL_Body"]))
        story.append(Spacer(1, 8))
        
        hm_img = _build_heatmap_image(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=6.0*inch, height=4.5*inch))
            story.append(Spacer(1, 10))
        
        explanation_header = Paragraph("Understanding the Heatmap", styles["OSIL_SectionHeader"])
        explanation_text = Paragraph(
            "The heatmap above displays stability risk concentration across your top services. "
            "Each cell represents a risk score from 0 (green, stable) to 100 (red, critical). "
            "High scores in the <b>Recurrence</b> column indicate services with repeated incidents. "
            "<b>MTTR Drag</b> shows slow recovery times. <b>Reopen</b> highlights incidents that were "
            "closed prematurely and reopened. <b>Change</b> indicates instability related to change activity. "
            "Services with multiple red/orange cells require immediate executive attention and SIP intervention.",
            styles["OSIL_Body"]
        )
        
        story.append(KeepTogether([explanation_header, Spacer(1, 4), explanation_text]))
        story.append(Spacer(1, 8))
        
        story.append(Paragraph("Risk Score Guide", styles["OSIL_SectionHeader"]))
        risk_guide_data = [
            ["Risk Level", "Color Indicator", "Interpretation", "Recommended Action"],
            ["0-30", "Green", "Stable operations within acceptable parameters", "Maintain current controls, monitor trends"],
            ["31-60", "Yellow", "Elevated risk requiring attention", "Review operational procedures, consider preventive measures"],
            ["61-80", "Orange", "High risk, potential service impact", "Initiate SIP, allocate resources for improvement"],
            ["81-100", "Red", "Critical risk, immediate intervention required", "Executive escalation, emergency SIP activation"]
        ]
        
        risk_table = _create_table(risk_guide_data, [1.0*inch, 1.3*inch, 2.2*inch, 2.0*inch],
                                  styles["OSIL_TableHeader"], styles["OSIL_TableCell"])
        story.append(risk_table)

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
