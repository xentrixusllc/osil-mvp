"""OSIL Executive PDF Report Generator - Executive Grade Version"""
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
            fontSize=24,
            leading=28,
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
            leading=18,
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
            leading=14,
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

    styles.add(
        ParagraphStyle(
            name="OSIL_Table",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableSmall",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=8,
            leading=10,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#0A192F"),
        )
    )
    
    return styles

def _footer(canvas, doc):
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(0.55 * inch, 0.35 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
    canvas.drawRightString(7.95 * inch, 0.35 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _header_band(text: str, width: float = 7.35 * inch) -> Table:
    t = Table([[text]], colWidths=[width], rowHeights=[28])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0A192F")),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 12),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ]
        )
    )
    return t

def _accent_rule(width: float = 7.35 * inch, color_hex: str = "#64FFDA", height: int = 3) -> Table:
    t = Table([[""]], colWidths=[width], rowHeights=[height])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor(color_hex)),
                ("LEFTPADDING", (0, 0), (-1, -1), 0),
                ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                ("TOPPADDING", (0, 0), (-1, -1), 0),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
            ]
        )
    )
    return t

def _kpi_box(label: str, value: str, width: float = 2.35 * inch) -> Table:
    t = Table([[label], [value]], colWidths=[width], rowHeights=[20, 40])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F7FA")),
                ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#D1D5DB")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("FONTSIZE", (0, 1), (-1, 1), 20),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#666666")),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#0A192F")),
                ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 4),
                ("TOPPADDING", (0, 1), (-1, 1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
            ]
        )
    )
    return t

def _signal_box(title: str, body: str, width: float = 7.35 * inch, styles=None) -> Table:
    title_para = Paragraph(f"<b>{_clean_text(title)}</b>", styles["OSIL_TableHeader"])
    body_para = Paragraph(_clean_text(body), styles["OSIL_Body"])
    
    t = Table([[title_para], [body_para]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF4FF")),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F8FBFF")),
                ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#B8D6F2")),
                ("LEFTPADDING", (0, 0), (-1, -1), 12),
                ("RIGHTPADDING", (0, 0), (-1, -1), 12),
                ("TOPPADDING", (0, 0), (-1, 0), 10),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("TOPPADDING", (0, 1), (-1, 1), 6),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 12),
            ]
        )
    )
    return t

def _simple_table(df: pd.DataFrame, col_widths: List[float], styles, small_font: bool = False) -> Table:
    """Create a simple table with proper wrapping"""
    df = _safe_df(df)
    if df.empty:
        data = [["No data available"]]
        tbl = Table(data, colWidths=[sum(col_widths)])
        tbl.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                ]
            )
        )
        return tbl

    style_name = "OSIL_TableSmall" if small_font else "OSIL_Table"
    header_style = "OSIL_TableHeader"
    
    # Convert all cells to Paragraphs for text wrapping
    header = [Paragraph(_clean_text(str(c)), styles[header_style]) for c in df.columns]
    rows = []
    for row in df.values.tolist():
        new_row = []
        for cell in row:
            text = _clean_text(str(cell))
            new_row.append(Paragraph(text, styles[style_name]))
        rows.append(new_row)
    
    tbl = Table([header] + rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#E9EDF3")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#0A192F")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 9),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 1), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
            ]
        )
    )
    return tbl

def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    """Build radar chart with fixed sizing"""
    labels = list(domain_scores.keys())
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]
    
    if not labels:
        labels = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
        values = [0, 0, 0, 0]
    
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]
    
    fig = plt.figure(figsize=(5, 4.5), dpi=150)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    
    ax.plot(angles_loop, values_loop, linewidth=2.5, color='#0A192F')
    ax.fill(angles_loop, values_loop, alpha=0.15, color='#64FFDA')
    
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=9)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=8)
    ax.set_title("Operational Stability Profile", pad=20, fontsize=12, fontweight='bold')
    
    plt.tight_layout()
    img = io.BytesIO()
    plt.savefig(img, format="png", bbox_inches="tight", facecolor='white', edgecolor='none')
    plt.close(fig)
    img.seek(0)
    return img

def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> Optional[io.BytesIO]:
    """Build heatmap with error handling and validation"""
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
        # Truncate long service names
        hm["Display_Name"] = hm["Service"].astype(str).str[:20] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        
        display_names = {
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision"
        }
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        # Determine figure size based on number of services
        height = max(4, len(hm) * 0.4)
        fig = plt.figure(figsize=(7, height), dpi=150)
        ax = plt.gca()
        
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='RdYlGn_r')
        ax.set_xticks(range(len(hm.columns)))
        ax.set_xticklabels(list(hm.columns), fontsize=9)
        ax.set_yticks(range(len(hm.index)))
        ax.set_yticklabels(list(hm.index), fontsize=9)
        ax.set_title("Service Stability Risk Heatmap", fontsize=12, fontweight='bold', pad=15)
        
        # Add text annotations
        for i in range(hm.shape[0]):
            for j in range(hm.shape[1]):
                val = int(round(float(hm.iat[i, j]), 0))
                text_color = "white" if val > 60 else "black"
                ax.text(j, i, str(val), ha="center", va="center", 
                       fontsize=9, color=text_color, fontweight='bold')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label("Risk Score (0-100)", fontsize=9)
        cbar.ax.tick_params(labelsize=8)
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap generation error: {e}")
        return None

def _posture_signal_text(bvsi: float, posture: str) -> str:
    if bvsi >= 85:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Technology operations are broadly supporting business confidence, though targeted prevention can strengthen resilience further."
    if bvsi >= 70:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Control exists across most operational dimensions, though localized weaknesses remain that could limit reliability at scale."
    if bvsi >= 55:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Governance mechanisms are functioning, but recurring instability patterns remain across higher-impact services."
    if bvsi >= 40:
        return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Operational exposure is visible and likely affecting service confidence, efficiency, or customer experience."
    return f"Operational stability is currently <b>{posture}</b>, with a BVSI™ score of {bvsi:.1f}. Immediate executive attention is warranted to contain structural instability and protect business performance."

def _bvsi_scale_table(styles) -> Table:
    df = pd.DataFrame(
        [
            ["85-100", "High Confidence Operations", "Technology stability supports business growth and executive confidence."],
            ["70-84", "Controlled and Improving", "Core controls exist; targeted improvement can increase resilience."],
            ["55-69", "Controlled but Exposed", "Operational control exists, but recurring instability creates material exposure."],
            ["40-54", "Reactive and Exposed", "Instability is visible and may be impacting reliability, cost, or customer experience."],
            ["<40", "Fragile Operations", "Operational fragility is high; stabilization is an executive priority."],
        ],
        columns=["BVSI™ Range", "Operating Condition", "Executive Meaning"],
    )
    return _simple_table(df, [1.0 * inch, 2.0 * inch, 4.35 * inch], styles)

def _domain_definitions_table(domain_scores: Dict[str, float], styles) -> Table:
    rows = [
        ["Service Resilience", f'{round(_safe_float(domain_scores.get("Service Resilience", 0.0)), 1)}', "Ability to recover quickly and consistently from incidents."],
        ["Change Governance", f'{round(_safe_float(domain_scores.get("Change Governance", 0.0)), 1)}', "Effectiveness of controls preventing operational instability during change activity."],
        ["Structural Risk Debt™", f'{round(_safe_float(domain_scores.get("Structural Risk Debt™", 0.0)), 1)}', "Accumulated instability caused by unresolved recurring weaknesses and operational debt."],
        ["Reliability Momentum", f'{round(_safe_float(domain_scores.get("Reliability Momentum", 0.0)), 1)}', "Direction of operational reliability based on recurring instability and recovery behavior."],
    ]
    df = pd.DataFrame(rows, columns=["Domain", "Score", "Definition"])
    return _simple_table(df, [2.0 * inch, 0.8 * inch, 4.55 * inch], styles)

def _top_three_briefs(sip_candidates: pd.DataFrame) -> pd.DataFrame:
    df = _safe_df(sip_candidates)
    if df.empty:
        return pd.DataFrame(columns=["Initiative", "Why Leadership Should Care"])
    
    brief_rows = []
    for _, row in df.head(3).iterrows():
        svc = str(row.get("Service", "Unknown Service"))
        tier = str(row.get("Service_Tier", "Unknown Tier"))
        theme = str(row.get("Suggested_Theme", "Stability Improvement"))
        why = str(row.get("Why_Flagged", "Elevated operational instability"))
        brief_rows.append({
            "Initiative": f"{svc} ({tier}) - {theme}",
            "Why Leadership Should Care": why
        })
    
    return pd.DataFrame(brief_rows)

def _key_takeaways(domain_scores: Dict[str, float], sip_candidates: pd.DataFrame) -> pd.DataFrame:
    strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Service Resilience"
    weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Structural Risk Debt™"
    
    top_service = "No service identified"
    priority_label = "Monitor"
    if isinstance(sip_candidates, pd.DataFrame) and not sip_candidates.empty:
        top_service = str(sip_candidates.iloc[0].get("Service", top_service))
        priority_label = str(sip_candidates.iloc[0].get("Priority_Label", priority_label))
    
    rows = [
        {"Key Insight": f"Strength: {strongest} provides the strongest stability support."},
        {"Key Insight": f"Priority: Focus on {weakest} for near-term stabilization."},
        {"Key Insight": f"Action: {top_service} is the leading candidate ({priority_label})."},
    ]
    return pd.DataFrame(rows)

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
        executive_interpretation = _clean_text(payload.get("executive_interpretation", ""))
        domain_scores = payload.get("domain_scores", {}) or {}
        service_risk_top10 = _safe_df(payload.get("service_risk_top10"))
        sip_candidates = _safe_df(payload.get("sip_candidates"))
        detected_dataset = str(payload.get("detected_dataset", "unknown")).upper()
        service_anchor_used = str(payload.get("service_anchor_used", "None"))
        data_readiness_score = _safe_float(payload.get("data_readiness_score", 0.0))

        story = []

        # Page 1 - Executive Summary
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report - {tenant_name}", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"As of {as_of}", styles["OSIL_Small"]))
        story.append(Spacer(1, 12))
        story.append(_header_band("Executive Brief"))
        story.append(_accent_rule())
        story.append(Spacer(1, 16))

        # KPI Row
        metric_row = Table(
            [[_kpi_box("BVSI™ Score", f"{bvsi:.1f}"), 
              _kpi_box("Operating Posture", posture), 
              _kpi_box("Data Readiness", f"{data_readiness_score:.1f}%")]],
            colWidths=[2.35 * inch, 2.35 * inch, 2.35 * inch],
        )
        metric_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(metric_row)
        story.append(Spacer(1, 20))

        # BVSI Scale
        story.append(Paragraph("BVSI™ Scale Reference", styles["OSIL_Section"]))
        story.append(_bvsi_scale_table(styles))
        story.append(Spacer(1, 16))

        # Executive Signal Box
        story.append(_signal_box("Executive Assessment", _posture_signal_text(bvsi, posture), styles=styles))
        story.append(Spacer(1, 16))

        # Executive Summary
        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(executive_interpretation, styles["OSIL_Body"]))
        story.append(Spacer(1, 16))

        # Key Takeaways
        story.append(Paragraph("Key Takeaways", styles["OSIL_Section"]))
        story.append(_simple_table(_key_takeaways(domain_scores, sip_candidates), [7.0 * inch], styles))
        story.append(Spacer(1, 12))
        
        # Footer info
        story.append(Paragraph(
            f"<i>Dataset: {detected_dataset} | Service Anchor: {service_anchor_used} | Organization: {tenant_name}</i>", 
            styles["OSIL_Small"]
        ))

        # Page 2 - Stability Profile
        story.append(PageBreak())
        story.append(_header_band("Operational Stability Profile"))
        story.append(_accent_rule())
        story.append(Spacer(1, 16))
        
        story.append(Paragraph(
            "This radar chart displays current performance across four critical stability domains. "
            "A balanced shape indicates aligned governance; collapsed areas highlight maturity gaps requiring attention.",
            styles["OSIL_Body"]
        ))
        story.append(Spacer(1, 12))

        # Radar Chart
        try:
            radar_img = _build_radar_image(domain_scores)
            story.append(Image(radar_img, width=4.5 * inch, height=4.0 * inch))
        except Exception as e:
            story.append(Paragraph(f"Chart generation error: {e}", styles["OSIL_Small"]))
        
        story.append(Spacer(1, 12))
        story.append(Paragraph("Domain Score Definitions", styles["OSIL_Section"]))
        story.append(_domain_definitions_table(domain_scores, styles))
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            "<i>Score Guide: 80-100 = Strong Maturity | 60-79 = Controlled | 40-59 = Weakness | <40 = Structural Instability</i>",
            styles["OSIL_Small"]
        ))

        # Page 3 - SIP Priorities
        story.append(PageBreak())
        story.append(_header_band("Service Improvement Priorities"))
        story.append(_accent_rule())
        story.append(Spacer(1, 16))
        
        story.append(Paragraph(
            "The following initiatives represent the highest-impact stability improvements for the next 30-60 days. "
            "These should guide leadership briefing and resource allocation.",
            styles["OSIL_Body"]
        ))
        story.append(Spacer(1, 16))

        story.append(Paragraph("Top 3 Initiatives for Leadership", styles["OSIL_Section"]))
        top3_df = _top_three_briefs(sip_candidates)
        if not top3_df.empty:
            story.append(_simple_table(top3_df, [3.0 * inch, 4.0 * inch], styles))
        else:
            story.append(Paragraph("No SIP candidates identified.", styles["OSIL_Body"]))
        
        story.append(Spacer(1, 20))
        story.append(Paragraph("Detailed SIP Candidates", styles["OSIL_Section"]))
        
        # Detailed table - limit columns if too many
        detailed = sip_candidates.head(10).copy()
        if not detailed.empty:
            # Select and rename columns for better display
            display_cols = ["Service", "Service_Tier", "Suggested_Theme", "SIP_Priority_Score", "Priority_Label", "Why_Flagged"]
            available_cols = [c for c in display_cols if c in detailed.columns]
            detailed_display = detailed[available_cols]
            
            # Calculate column widths based on content
            col_count = len(available_cols)
            if col_count == 6:
                widths = [1.5*inch, 0.8*inch, 1.2*inch, 0.8*inch, 0.8*inch, 1.9*inch]
            else:
                widths = [7.0 / col_count * inch] * col_count
            
            story.append(_simple_table(detailed_display, widths, styles, small_font=True))
        else:
            story.append(Paragraph("No detailed data available.", styles["OSIL_Body"]))

        # Page 4 - Heatmap
        story.append(PageBreak())
        story.append(_header_band("Service Stability Heatmap"))
        story.append(_accent_rule())
        story.append(Spacer(1, 16))
        
        story.append(Paragraph(
            "This heatmap visualizes stability risk across services. Darker colors indicate higher risk. "
            "Services with high values across multiple columns are priority candidates for intervention.",
            styles["OSIL_Body"]
        ))
        story.append(Spacer(1, 16))

        hm_img = _build_heatmap_image(service_risk_top10)
        if hm_img is not None:
            story.append(Image(hm_img, width=6.5 * inch))
        else:
            story.append(Paragraph(
                "Insufficient data to generate heatmap. Ensure incident data includes Service, Service_Tier, and risk metrics.",
                styles["OSIL_Body"]
            ))

        # Build PDF
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        raise Exception(f"PDF Generation Failed: {str(e)}")
