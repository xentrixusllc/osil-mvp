"""OSIL Executive Report Production Design"""
import io
import re
import textwrap
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

def _get_posture_from_bvsi(bvsi: float) -> str:
    if bvsi >= 85: return "High Confidence Operations"
    elif bvsi >= 70: return "Controlled and Improving"
    elif bvsi >= 55: return "Controlled but Exposed"
    elif bvsi >= 40: return "Reactive and Exposed"
    else: return "Fragile Operations"

def _styles():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="ExecutiveTitle", parent=styles["Title"], fontName="Helvetica-Bold", fontSize=24, leading=28, textColor=colors.HexColor("#0F172A"), spaceAfter=6, alignment=0))
    styles.add(ParagraphStyle(name="ExecutiveSubtitle", parent=styles["BodyText"], fontName="Helvetica", fontSize=11, leading=14, textColor=colors.HexColor("#64748B"), spaceAfter=24))
    styles.add(ParagraphStyle(name="SectionHeader", parent=styles["Heading2"], fontName="Helvetica-Bold", fontSize=16, leading=20, textColor=colors.HexColor("#0F172A"), spaceAfter=12, spaceBefore=18))
    styles.add(ParagraphStyle(name="PageHeader", parent=styles["Heading1"], fontName="Helvetica-Bold", fontSize=18, leading=22, textColor=colors.HexColor("#0F172A"), spaceAfter=16))
    styles.add(ParagraphStyle(name="ExecutiveBody", parent=styles["BodyText"], fontName="Helvetica", fontSize=11, leading=16, textColor=colors.HexColor("#334155"), spaceAfter=12))
    styles.add(ParagraphStyle(name="Caption", parent=styles["BodyText"], fontName="Helvetica", fontSize=9, leading=12, textColor=colors.HexColor("#64748B"), spaceAfter=8))
    styles.add(ParagraphStyle(name="TableHeader", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=10, leading=12, textColor=colors.white))
    styles.add(ParagraphStyle(name="TableCell", parent=styles["BodyText"], fontName="Helvetica", fontSize=10, leading=14, textColor=colors.HexColor("#334155")))
    styles.add(ParagraphStyle(name="TableCellBold", parent=styles["BodyText"], fontName="Helvetica-Bold", fontSize=10, leading=14, textColor=colors.HexColor("#0F172A")))
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
        values = [_safe_float(domain_scores.get(d, domain_scores.get(d + "™", 0)), 0) for d in domains]
        angles = np.linspace(0, 2 * np.pi, len(domains), endpoint=False).tolist()
        values += values[:1]
        angles += angles[:1]
        fig, ax = plt.subplots(figsize=(4.0, 3.5), subplot_kw=dict(projection='polar'))
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        ax.fill(angles, values, color='#3B82F6', alpha=0.2)
        ax.plot(angles, values, color='#2563EB', linewidth=2.5, marker='o', markersize=5)
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(labels, fontsize=9, fontweight='bold', color='#0F172A')
        ax.set_ylim(0, 100)
        ax.grid(True, linestyle='solid', alpha=0.2, color='#94A3B8')
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none', pad_inches=0.1)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_heatmap(service_risk_df: pd.DataFrame) -> Optional[io.BytesIO]:
    df = _safe_df(service_risk_df)
    if df.empty or not all(c in df.columns for c in ["Service", "Service_Tier"]): return None
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Execution_Churn_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_names = {"Recurrence_Risk": "Recurrence", "MTTR_Drag_Risk": "MTTR Drag", "Execution_Churn_Risk": "Exec Churn", "Reopen_Churn_Risk": "Reopen", "Change_Collision_Risk": "Change"}
    try:
        hm = df.head(8).copy()
        hm["Display"] = hm["Service"].astype(str).str[:14] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display")[[c for c in risk_cols if c in df.columns]].rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0)
        fig, ax = plt.subplots(figsize=(8.0, 5.5))
        im = ax.imshow(hm.values, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(hm.columns, fontsize=10, fontweight='bold', color='#0F172A')
        ax.set_yticklabels(hm.index, fontsize=10, color='#334155')
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label('Risk Score', fontsize=10, fontweight='bold', color='#0F172A')
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(hm.iloc[i, j], 0))
                ax.text(j, i, str(val), ha='center', va='center', fontsize=10, color='white' if val > 60 else '#0F172A', fontweight='bold')
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none', pad_inches=0.2)
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_pareto_image(df: pd.DataFrame) -> Optional[io.BytesIO]:
    if df.empty: return None
    try:
        fig, ax1 = plt.subplots(figsize=(7.0, 4.5), dpi=120)
        ax1.bar(np.arange(len(df)), df["Frequency"], color="#3B82F6", width=0.55)
        ax2 = ax1.twinx()
        ax2.plot(np.arange(len(df)), df["Cumulative_Pct"], color="#DC2626", marker="o", linewidth=2.5)
        ax2.set_ylim(0, 110)
        plt.title("Eighty Twenty Rule: Top Structural Risk Themes", fontweight="bold", color='#0F172A', pad=15)
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_impact_matrix_image(service_risk_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if service_risk_df.empty or "Active_Disruption_P1_P2" not in service_risk_df.columns: return None
    try:
        merged = service_risk_df.sort_values(by=["Active_Disruption_P1_P2", "Total_Service_Risk"], ascending=[False, False]).head(5).copy()
        fig, ax1 = plt.subplots(figsize=(7.5, 4.5), dpi=120)
        ax1.bar(np.arange(len(merged)), merged["Active_Disruption_P1_P2"], color="#DC2626", width=0.45, alpha=0.9)
        ax2 = ax1.twinx()
        ax2.plot(np.arange(len(merged)), merged["Recurrence_Risk"], color="#0F172A", marker="o", linewidth=2.5)
        ax2.set_ylim(0, 115)
        plt.title("Executive Strike Zone: Top 5 Services", fontweight="bold", color="#0F172A", pad=15)
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_macro_trend_image(history_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if history_df.empty or len(history_df) <= 1: return None
    try:
        df = history_df.copy()
        df["display_date"] = pd.to_datetime(df["run_date"]).dt.strftime('%m/%Y')
        fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=120)
        ax.plot(df["display_date"], df["bvsi_score"], marker='o', linewidth=3.0, color='#0F172A', label='Global Stability (BVSI™)')
        ax.plot(df["display_date"], df["debt_score"], marker='s', linewidth=2.0, color='#DC2626', linestyle='--', label='Structural Risk Debt™')
        ax.set_ylim(0, 110)
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_micro_trend_image(history_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if history_df.empty or len(history_df) <= 1: return None
    try:
        df = history_df.copy()
        df["display_date"] = pd.to_datetime(df["run_date"]).dt.strftime('%m/%Y')
        fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=120)
        ax.plot(df["display_date"], df["resilience_score"], marker='^', linewidth=2.0, color='#2563EB', label='Resilience')
        ax.plot(df["display_date"], df["governance_score"], marker='d', linewidth=2.0, color='#059669', label='Governance')
        ax.plot(df["display_date"], df["momentum_score"], marker='*', linewidth=2.0, color='#D97706', label='Momentum')
        ax.plot(df["display_date"], df["debt_score"], marker='s', linewidth=2.0, color='#DC2626', linestyle=':', label='Risk Debt')
        ax.set_ylim(0, 110)
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception: return None

def _build_domain_insight_box(domain_scores: dict, styles: Any) -> Table:
    weakest = min(domain_scores.items(), key=lambda x: float(x[1]))[0] if domain_scores else "Resilience"
    insight = "<b>Cost of Inaction:</b> Operational friction is decaying stability. <b>Mandate:</b> Executive review."
    if "Resilience" in weakest: insight = "<b>Cost of Inaction:</b> Prolonged recovery times degrade trust. <b>Mandate:</b> Automate runbooks."
    elif "Governance" in weakest: insight = "<b>Cost of Inaction:</b> Release collisions threaten revenue. <b>Mandate:</b> Institute mandatory peer reviews."
    elif "Debt" in weakest: insight = "<b>Cost of Inaction:</b> Recurring issues tax IT capacity. <b>Mandate:</b> Reallocate capacity to debt paydown."
    
    box = Table([[Paragraph(f"Primary Exposure: {weakest}", styles["TableHeader"])], [Paragraph(insight, styles["TableCell"])]], colWidths=[3.9*inch])
    box.setStyle(TableStyle([('BACKGROUND', (0,0), (-1,0), colors.HexColor("#D97706")), ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#D97706")), ('LEFTPADDING', (0,0), (-1,-1), 10)]))
    return box

def build_osil_pdf_report(payload: Dict[str, Any]) -> bytes:
    try:
        styles = _styles()
        out = io.BytesIO()
        doc = SimpleDocTemplate(out, pagesize=LETTER, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.75*inch, bottomMargin=0.75*inch)
        story = []
        
        # Header
        story.append(Paragraph("OSIL™ Executive Briefing", styles["ExecutiveTitle"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report for {payload.get('tenant_name')} | As of {payload.get('as_of')}", styles["ExecutiveSubtitle"]))
        
        # Assessment & Remediation
        svs = payload.get("svs_scores", {})
        svc = payload.get("svc_scores", {})
        weakest_svs = min(svs, key=svs.get) if svs else "Practices"
        weakest_svc = min(svc, key=svc.get) if svc else "Deliver & Support"
        svs_adv = {"Governance": "Invest in policy enforcement.", "Continual Improvement": "Reallocate capacity to debt paydown.", "Practices": "Standardize assignment routing.", "Guiding Principles": "Realign IT with business outcomes."}
        svc_adv = {"Plan": "Immediate data audit.", "Improve": "Implement formal Problem Management.", "Engage": "Bridge trust gaps.", "Design & Transition": "Enforce Tier 1 validation.", "Obtain/Build": "Formalize quality gates.", "Deliver & Support": "Deploy automated playbooks."}
        
        rem_data = [[Paragraph("Focus Area", styles["TableHeader"]), Paragraph("Diagnostic Analysis & Mandate", styles["TableHeader"])],
                    [Paragraph(f"SVS: {weakest_svs}", styles["TableCellBold"]), Paragraph(f"Bottleneck: {weakest_svs}. <b>Mandate:</b> {svs_adv.get(weakest_svs)}", styles["TableCell"])],
                    [Paragraph(f"SVC: {weakest_svc}", styles["TableCellBold"]), Paragraph(f"Execution Failure: {weakest_svc}. <b>Mandate:</b> {svc_adv.get(weakest_svc)}", styles["TableCell"])]]
        
        story.append(Paragraph("Operating Model Assessment", styles["PageHeader"]))
        story.append(Table(rem_data, colWidths=[2.0*inch, 4.5*inch], style=TableStyle([('GRID',(0,0),(-1,-1),0.5,colors.grey), ('BACKGROUND',(0,0),(-1,0),colors.HexColor("#0F172A")), ('TEXTCOLOR',(0,0),(-1,0),colors.white)])))
        
        # SVC Diagram
        def _get_svc_pdf_color(s): return colors.HexColor("#059669") if s >= 70 else colors.HexColor("#D97706") if s >= 55 else colors.HexColor("#DC2626")
        c_p, c_imp, c_eng, c_dt, c_ob, c_ds = [_get_svc_pdf_color(_safe_float(svc.get(k, 0), 0)) for k in ['Plan', 'Improve', 'Engage', 'Design & Transition', 'Obtain/Build', 'Deliver & Support']]
        box_s = ParagraphStyle(name='svc_box', parent=styles['BodyText'], textColor=colors.white, fontName='Helvetica-Bold', alignment=1, fontSize=10)
        
        svc_diag = Table([
            [Table([[Paragraph(f"Plan ({svc.get('Plan', 0):.1f})", box_s)]], style=[('BACKGROUND',(0,0),(-1,-1),c_p)])],
            [Table([[Table([[Paragraph(f"Engage<br/>({svc.get('Engage', 0):.1f})", box_s)]], style=[('BACKGROUND',(0,0),(-1,-1),c_eng)]),
                     Table([[Paragraph(f"Design ({svc.get('Design & Transition', 0):.1f})", box_s)], [Paragraph(f"Build ({svc.get('Obtain/Build', 0):.1f})", box_s)], [Paragraph(f"Deliver ({svc.get('Deliver & Support', 0):.1f})", box_s)]], 
                           style=[('BACKGROUND',(0,0),(-1,0),c_dt), ('BACKGROUND',(0,1),(-1,1),c_ob), ('BACKGROUND',(0,2),(-1,2),c_ds)])]])],
            [Table([[Paragraph(f"Improve ({svc.get('Improve', 0):.1f})", box_s)]], style=[('BACKGROUND',(0,0),(-1,-1),c_imp)])]
        ], colWidths=[6.5*inch])
        story.append(Spacer(1, 24))
        story.append(svc_diag)
        
        # Trend Intelligence
        if not payload.get("history_df", pd.DataFrame()).empty:
            story.append(PageBreak())
            story.append(Paragraph("Executive Trend Intelligence", styles["PageHeader"]))
            if macro := _build_macro_trend_image(payload["history_df"]): story.append(Image(macro, width=6.2*inch, height=3.2*inch))
            if micro := _build_micro_trend_image(payload["history_df"]): story.append(Image(micro, width=6.2*inch, height=3.2*inch))
        
        # Trust Gap Matrix
        if not payload.get("trust_gap_df", pd.DataFrame()).empty:
            story.append(PageBreak())
            trust_data = [[Paragraph("Service", styles["TableHeader"]), Paragraph("Critical Disruption", styles["TableHeader"]), Paragraph("Silent Friction", styles["TableHeader"]), Paragraph("Friction Ratio", styles["TableHeader"])]]
            for _, row in payload["trust_gap_df"].iterrows():
                trust_data.append([Paragraph(str(row["Service"]), styles["TableCellBold"]), Paragraph(str(row["Active_Disruption_P1_P2"]), styles["TableCell"]), Paragraph(str(row["Silent_Friction_P3_P5"]), styles["TableCell"]), Paragraph(f"{_safe_float(row.get('Friction_Ratio',0)):.1f}x", styles["TableCell"])])
            story.append(Paragraph("Xentrixus Trust Gap Matrix (P1 to P5)", styles["SectionHeader"]))
            story.append(Table(trust_data, colWidths=[2.5*inch, 1.5*inch, 1.5*inch, 1.0*inch], style=TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor("#0F172A")), ('TEXTCOLOR',(0,0),(-1,0),colors.white)])))

        # Pareto
        if not payload.get("rca_pareto_df", pd.DataFrame()).empty:
            story.append(PageBreak())
            story.append(Paragraph("Structural Risk Debt™: Thematic Extraction", styles["PageHeader"]))
            if pareto := _build_pareto_image(payload["rca_pareto_df"]): story.append(Image(pareto, width=6.5*inch, height=4.2*inch))
            
        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        return out.getvalue()
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
