"""OSIL Executive PDF Report Generator - Fixed Version"""
import io
import re
from typing import Any, Dict, List, Optional
import matplotlib
matplotlib.use('Agg')  # Fix: Use non-interactive backend before importing pyplot
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
    """Safely convert value to float"""
    try:
        if val is None or (isinstance(val, float) and np.isnan(val)):
            return default
        return float(val)
    except Exception:
        return default

def _safe_df(val: Any) -> pd.DataFrame:
    """Safely convert to DataFrame"""
    if isinstance(val, pd.DataFrame):
        return val.copy()
    return pd.DataFrame()

def _clean_text(text: Any) -> str:
    """Clean text for PDF rendering"""
    if text is None:
        return ""
    s = str(text)
    # Bold formatting
    s = re.sub(r"\*\*(.*?)\*\*", r"<b>\1</b>", s)
    s = s.replace("<br>", "<br/>").replace("<br />", "<br/>")
    # Only allow safe HTML tags
    allowed = {"b", "i", "u", "br"}
    
    def repl_tag(match):
        full = match.group(0)
        tag = match.group(1).lower().replace("/", "")
        return full if tag in allowed else ""
    
    s = re.sub(r"</?([a-zA-Z0-9]+).*?>", repl_tag, s)
    return s

def _styles():
    """Create and return styles"""
    styles = getSampleStyleSheet()
    
    styles.add(
        ParagraphStyle(
            name="OSIL_Title",
            parent=styles["Title"],
            fontName="Helvetica-Bold",
            fontSize=22,
            leading=26,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=4,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Subtitle",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9,
            leading=11,
            textColor=colors.HexColor("#555555"),
            spaceAfter=8,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Section",
            parent=styles["Heading2"],
            fontName="Helvetica-Bold",
            fontSize=13.5,
            leading=16,
            textColor=colors.HexColor("#0A192F"),
            spaceAfter=6,
            spaceBefore=2,
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_Body",
            parent=styles["BodyText"],
            fontName="Helvetica",
            fontSize=9.2,
            leading=12.6,
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
            fontSize=7.8,
            leading=9.3,
            textColor=colors.HexColor("#222222"),
        )
    )

    styles.add(
        ParagraphStyle(
            name="OSIL_TableHeader",
            parent=styles["BodyText"],
            fontName="Helvetica-Bold",
            fontSize=8,
            leading=10,
            textColor=colors.black,
        )
    )
    
    return styles

def _footer(canvas, doc):
    """Add footer to each page"""
    canvas.saveState()
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#666666"))
    canvas.drawString(0.55 * inch, 0.35 * inch, "OSIL™ by Xentrixus • Operational Stability Intelligence")
    canvas.drawRightString(7.95 * inch, 0.35 * inch, f"Page {doc.page}")
    canvas.restoreState()

def _header_band(text: str, width: float = 7.35 * inch) -> Table:
    """Create header band"""
    t = Table([[text]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#0A192F")),
                ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, -1), 12),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    return t

def _accent_rule(width: float = 7.35 * inch, color_hex: str = "#64FFDA", height: int = 3) -> Table:
    """Create accent line"""
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

def _paragraph_table(df: pd.DataFrame, col_widths: List[float], styles, header_bg: str = "#E9EDF3") -> Table:
    """Create paragraph-based table with text wrapping"""
    df = _safe_df(df)
    if df.empty:
        data = [[Paragraph("No data available", styles["OSIL_Table"])]]
        tbl = Table(data, colWidths=[sum(col_widths)])
        tbl.setStyle(
            TableStyle(
                [
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.grey),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("TOPPADDING", (0, 0), (-1, -1), 4),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ]
            )
        )
        return tbl

    header = [Paragraph(_clean_text(c), styles["OSIL_TableHeader"]) for c in df.columns]
    rows = [[Paragraph(_clean_text(str(x)), styles["OSIL_Table"]) for x in row] for row in df.values.tolist()]
    
    tbl = Table([header] + rows, colWidths=col_widths, repeatRows=1)
    tbl.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(header_bg)),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.grey),
                ("VALIGN", (0, 0), (-1, -1), "TOP"),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, colors.HexColor("#F8FAFC")]),
                ("LEFTPADDING", (0, 0), (-1, -1), 5),
                ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                ("TOPPADDING", (0, 0), (-1, -1), 4),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ]
        )
    )
    return tbl

def _kpi_box(label: str, value: str, width: float = 2.35 * inch) -> Table:
    """Create KPI display box"""
    t = Table([[label], [value]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#F5F7FA")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#D1D5DB")),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"),
                ("FONTSIZE", (0, 0), (-1, 0), 8),
                ("FONTSIZE", (0, 1), (-1, 1), 17),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#666666")),
                ("TEXTCOLOR", (0, 1), (-1, 1), colors.HexColor("#111111")),
                ("ALIGN", (0, 0), (-1, -1), "LEFT"),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
                ("TOPPADDING", (0, 1), (-1, 1), 0),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
            ]
        )
    )
    return t

def _signal_box(title: str, body: str, width: float = 7.35 * inch, styles=None) -> Table:
    """Create signal/info box"""
    body_para = Paragraph(_clean_text(body), styles["OSIL_Body"])
    title_para = Paragraph(_clean_text(title), styles["OSIL_TableHeader"])
    
    t = Table([[title_para], [body_para]], colWidths=[width])
    t.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#EAF4FF")),
                ("BACKGROUND", (0, 1), (-1, 1), colors.HexColor("#F8FBFF")),
                ("BOX", (0, 0), (-1, -1), 0.6, colors.HexColor("#B8D6F2")),
                ("LEFTPADDING", (0, 0), (-1, -1), 10),
                ("RIGHTPADDING", (0, 0), (-1, -1), 10),
                ("TOPPADDING", (0, 0), (-1, 0), 8),
                ("BOTTOMPADDING", (0, 0), (-1, 0), 6),
                ("TOPPADDING", (0, 1), (-1, 1), 4),
                ("BOTTOMPADDING", (0, 1), (-1, 1), 10),
            ]
        )
    )
    return t

def _build_radar_image(domain_scores: Dict[str, float]) -> io.BytesIO:
    """Build radar chart image with proper cleanup"""
    labels = list(domain_scores.keys())
    values = [_safe_float(domain_scores.get(k, 0.0)) for k in labels]
    
    if not labels:
        labels = ["Service Resilience", "Change Governance", "Structural Risk Debt", "Reliability Momentum"]
        values = [0, 0, 0, 0]
    
    angles = np.linspace(0, 2 * np.pi, len(labels), endpoint=False).tolist()
    angles_loop = angles + [angles[0]]
    values_loop = values + [values[0]]
    
    # Create new figure
    fig = plt.figure(figsize=(4.4, 4.1), dpi=180)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2)
    ax.set_theta_direction(-1)
    ax.plot(angles_loop, values_loop, linewidth=2, color='#0A192F')
    ax.fill(angles_loop, values_loop, alpha=0.12, color='#64FFDA')
    ax.set_xticks(angles)
    ax.set_xticklabels(labels, fontsize=8)
    ax.set_ylim(0, 100)
    ax.set_yticks([20, 40, 60, 80, 100])
    ax.set_yticklabels(["20", "40", "60", "80", "100"], fontsize=7)
    ax.set_title("Operational Stability Profile", pad=16, fontsize=11)
    
    img = io.BytesIO()
    plt.tight_layout()
    plt.savefig(img, format="png", bbox_inches="tight", facecolor='white')
    plt.close(fig)  # Critical: Close figure to free memory
    img.seek(0)
    return img

def _build_heatmap_image(service_risk_top10: pd.DataFrame) -> Optional[io.BytesIO]:
    """Build heatmap image with validation"""
    df = _safe_df(service_risk_top10)
    if df.empty:
        return None
    
    # Check for required columns with flexible naming
    required_cols = ["Service", "Service_Tier"]
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    
    # Check if we have the data needed
    missing_required = [c for c in required_cols if c not in df.columns]
    if missing_required:
        return None
    
    # Get available risk columns
    available_risks = [c for c in risk_cols if c in df.columns]
    if not available_risks:
        return None
    
    try:
        hm = df.head(10).copy()
        hm["Display_Name"] = hm["Service"].astype(str) + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display_Name")[available_risks]
        
        # Rename columns for display
        display_names = {
            "Recurrence_Risk": "Recurrence",
            "MTTR_Drag_Risk": "MTTR Drag",
            "Reopen_Churn_Risk": "Reopen Churn",
            "Change_Collision_Risk": "Change Collision"
        }
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0.0)
        
        fig = plt.figure(figsize=(6.8, 3.9), dpi=180)
        ax = plt.gca()
        im = ax.imshow(hm.values, aspect="auto", vmin=0, vmax=100, cmap='YlOrRd')
        ax.set_xticks(range(len(hm.columns)))
        ax.set_xticklabels(list(hm.columns), fontsize=8)
        ax.set_yticks(range(len(hm.index)))
        ax.set_yticklabels(list(hm.index), fontsize=8)
        ax.set_title("Service Stability Heatmap", fontsize=11)
        
        for i in range(hm.shape[0]):
            for j in range(hm.shape[1]):
                val = int(round(float(hm.iat[i, j]), 0))
                ax.text(j, i, str(val), ha="center", va="center", 
                       fontsize=8, color="white" if val > 50 else "black")
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.03, pad=0.02)
        cbar.set_label("Risk Score", fontsize=8)
        cbar.ax.tick_params(labelsize=7)
        
        img = io.BytesIO()
        plt.tight_layout()
        plt.savefig(img, format="png", bbox_inches="tight", facecolor='white')
        plt.close(fig)  # Critical: Close figure
        img.seek(0)
        return img
    except Exception as e:
        print(f"Heatmap generation error: {e}")
        return None

def _posture_signal_text(bvsi: float, posture: str) -> str:
    """Generate posture description"""
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
    """Create BVSI scale reference table"""
    df = pd.DataFrame(
        [
            ["85-100", "High Confidence Operations", "Technology stability supports business growth and executive confidence."],
            ["70-84", "Controlled and Improving", "Core controls exist; targeted improvement can increase resilience and scale-readiness."],
            ["55-69", "Controlled but Exposed", "Operational control exists, but recurring instability still creates material exposure."],
            ["40-54", "Reactive and Exposed", "Instability is visible and may be impacting reliability, cost, or customer experience."],
            ["<40", "Fragile Operations", "Operational fragility is high and stabilization should be treated as an executive priority."],
        ],
        columns=["BVSI™ Range", "Operating Condition", "Executive Meaning"],
    )
    return _paragraph_table(df, [1.0 * inch, 2.0 * inch, 4.35 * inch], styles)

def _domain_definitions_table(domain_scores: Dict[str, float], styles) -> Table:
    """Create domain definitions table"""
    rows = [
        ["Service Resilience", f'{round(_safe_float(domain_scores.get("Service Resilience", 0.0)), 1)}', "Ability to recover quickly and consistently from incidents."],
        ["Change Governance", f'{round(_safe_float(domain_scores.get("Change Governance", 0.0)), 1)}', "Effectiveness of controls preventing operational instability during change activity."],
        ["Structural Risk Debt™", f'{round(_safe_float(domain_scores.get("Structural Risk Debt™", 0.0)), 1)}', "Accumulated instability caused by unresolved recurring weaknesses and operational debt."],
        ["Reliability Momentum", f'{round(_safe_float(domain_scores.get("Reliability Momentum", 0.0)), 1)}', "Direction of operational reliability based on recurring instability and recovery behavior."],
    ]
    df = pd.DataFrame(rows, columns=["Domain", "Score", "What It Means"])
    return _paragraph_table(df, [2.0 * inch, 0.7 * inch, 4.65 * inch], styles)

def _top_three_briefs(sip_candidates: pd.DataFrame) -> pd.DataFrame:
    """Extract top 3 SIP briefs"""
    df = _safe_df(sip_candidates)
    if df.empty:
        return pd.DataFrame(columns=["Initiative", "Why Leadership Should Care"])
    
    brief_rows = []
    for _, row in df.head(3).iterrows():
        svc = str(row.get("Service", "Unknown Service"))
        tier = str(row.get("Service_Tier", "Unknown Tier"))
        theme = str(row.get("Suggested_Theme", "Stability Improvement"))
        why = str(row.get("Why_Flagged", "Elevated operational instability"))
        brief_rows.append({"Initiative": f"{svc} ({tier}) - {theme}", "Why Leadership Should Care": why})
    
    return pd.DataFrame(brief_rows)

def _key_takeaways(domain_scores: Dict[str, float], sip_candidates: pd.DataFrame) -> pd.DataFrame:
    """Generate key takeaways"""
    strongest = max(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Service Resilience"
    weakest = min(domain_scores.items(), key=lambda x: x[1])[0] if domain_scores else "Structural Risk Debt™"
    
    top_service = "No service identified"
    priority_label = "Monitor"
    if isinstance(sip_candidates, pd.DataFrame) and not sip_candidates.empty:
        top_service = str(sip_candidates.iloc[0].get("Service", top_service))
        priority_label = str(sip_candidates.iloc[0].get("Priority_Label", priority_label))
    
    rows = [
        {"Key Takeaway": f"Operational strength is most visible in {strongest}, which currently provides the strongest stability support."},
        {"Key Takeaway": f"The primary exposure area is {weakest}, and it should anchor near-term stabilization decisions."},
        {"Key Takeaway": f"{top_service} is the leading service candidate for action and is currently classified as {priority_label}."},
    ]
    return pd.DataFrame(rows)

def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    """Main function to build OSIL PDF report with error handling"""
    try:
        styles = _styles()
        out = io.BytesIO()
        
        doc = SimpleDocTemplate(
            out,
            pagesize=LETTER,
            rightMargin=0.55 * inch,
            leftMargin=0.55 * inch,
            topMargin=0.45 * inch,
            bottomMargin=0.55 * inch,
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

        # Page 1
        story.append(Paragraph("OSIL™ by Xentrixus", styles["OSIL_Title"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report - {tenant_name}", styles["OSIL_Subtitle"]))
        story.append(Paragraph(f"As of {as_of}", styles["OSIL_Small"]))
        story.append(Spacer(1, 6))
        story.append(_header_band("Executive Brief"))
        story.append(_accent_rule())
        story.append(Spacer(1, 10))

        metric_row = Table(
            [[_kpi_box("BVSI™ Score", f"{bvsi:.1f}"), _kpi_box("Operating Posture", posture), _kpi_box("Data Readiness", f"{data_readiness_score:.1f}%")]],
            colWidths=[2.45 * inch, 2.45 * inch, 2.45 * inch],
        )
        metric_row.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        story.append(metric_row)
        story.append(Spacer(1, 10))

        story.append(Paragraph("BVSI™ Interpretation", styles["OSIL_Section"]))
        story.append(_bvsi_scale_table(styles))
        story.append(Spacer(1, 10))

        story.append(_signal_box("Executive Signal", _posture_signal_text(bvsi, posture), styles=styles))
        story.append(Spacer(1, 10))

        story.append(Paragraph("Executive Summary", styles["OSIL_Section"]))
        story.append(Paragraph(executive_interpretation, styles["OSIL_Body"]))
        story.append(Spacer(1, 10))

        story.append(Paragraph("Key Takeaways", styles["OSIL_Section"]))
        story.append(_paragraph_table(_key_takeaways(domain_scores, sip_candidates), [7.2 * inch], styles))
        story.append(Spacer(1, 8))

        story.append(Paragraph(f"Detected dataset: {detected_dataset} - Service anchor used: {service_anchor_used} - Organization: {tenant_name}", styles["OSIL_Small"]))

        # Page 2
        story.append(PageBreak())
        story.append(_header_band("Operational Stability Profile"))
        story.append(_accent_rule())
        story.append(Spacer(1, 10))
        story.append(Paragraph("How to read this chart: higher scores indicate stronger operational stability. The radar shows current performance across service resilience, change governance, structural risk debt, and reliability momentum.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))

        try:
            radar_img = _build_radar_image(domain_scores)
            story.append(Image(radar_img, width=4.75 * inch, height=4.4 * inch))
        except Exception as e:
            story.append(Paragraph(f"Unable to generate radar chart: {e}", styles["OSIL_Small"]))

        story.append(Spacer(1, 8))
        story.append(Paragraph("Domain Scores", styles["OSIL_Section"]))
        story.append(_domain_definitions_table(domain_scores, styles))
        story.append(Paragraph("Score guide: 80-100 = strong maturity; 60-79 = controlled but improving; 40-59 = operational weakness; below 40 = structural instability.", styles["OSIL_Small"]))

        # Page 3
        story.append(PageBreak())
        story.append(_header_band("Service Improvement Priorities"))
        story.append(_accent_rule())
        story.append(Spacer(1, 10))
        story.append(Paragraph("These represent the highest-impact stability improvements for the next 30-60 days. They should be used to guide leadership briefing, ownership alignment, and targeted operational remediation.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))

        story.append(Paragraph("Top 3 Initiatives to Brief Leadership", styles["OSIL_Section"]))
        story.append(_paragraph_table(_top_three_briefs(sip_candidates), [3.2 * inch, 4.0 * inch], styles))
        story.append(Spacer(1, 12))

        story.append(Paragraph("Detailed SIP Candidates", styles["OSIL_Section"]))
        detailed = sip_candidates.head(10).copy()
        if not detailed.empty:
            widths = [7.0 / max(len(detailed.columns), 1) * inch] * len(detailed.columns)
            story.append(_paragraph_table(detailed, widths, styles))
        else:
            story.append(_paragraph_table(detailed, [7.0 * inch], styles))

        # Page 4
        story.append(PageBreak())
        story.append(_header_band("Service Stability Heatmap (Top 10 Services by Risk)"))
        story.append(_accent_rule())
        story.append(Spacer(1, 10))
        story.append(Paragraph("Executive view: Service × Stability Risk across recurrence, MTTR drag, reopen churn, and change collision. This visual helps leadership identify where concentrated instability is likely to create operational exposure.", styles["OSIL_Body"]))
        story.append(Spacer(1, 10))

        hm_img = _build_heatmap_image(service_risk_top10)
        if hm_img is not None:
            story.append(Image(hm_img, width=6.8 * inch, height=4.2 * inch))
        else:
            story.append(Paragraph("No service risk heatmap available - insufficient data.", styles["OSIL_Body"]))

        # Build PDF
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        # Return error as PDF text if build fails completely
        raise Exception(f"PDF Generation Failed: {str(e)}")
