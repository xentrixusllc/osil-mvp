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
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=6,
        alignment=0
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
        textColor=colors.HexColor("#0F172A"),
        spaceAfter=12,
        spaceBefore=18,
    ))
    
    styles.add(ParagraphStyle(
        name="PageHeader",
        parent=styles["Heading1"],
        fontName="Helvetica-Bold",
        fontSize=18,
        leading=22,
        textColor=colors.HexColor("#0F172A"),
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
        textColor=colors.HexColor("#0F172A"),
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
        
        fig, ax = plt.subplots(figsize=(4.0, 3.5), subplot_kw=dict(projection='polar'))
        ax.set_theta_offset(np.pi / 2)
        ax.set_theta_direction(-1)
        
        ax.fill(angles, values, color='#3B82F6', alpha=0.2)
        ax.plot(angles, values, color='#2563EB', linewidth=2.5, marker='o', markersize=5)
        
        ax.set_xticks(angles[:-1])
        ax.set_xticklabels(labels, fontsize=9, fontweight='bold', color='#0F172A')
        ax.set_ylim(0, 100)
        ax.set_yticks([20, 40, 60, 80, 100])
        ax.set_yticklabels(['20', '40', '60', '80', '100'], fontsize=7, color='#64748B')
        ax.grid(True, linestyle='solid', alpha=0.2, color='#94A3B8')
        
        plt.tight_layout()
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', 
                   edgecolor='none', pad_inches=0.1)
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
    
    risk_cols = ["Recurrence_Risk", "MTTR_Drag_Risk", "Execution_Churn_Risk", "Reopen_Churn_Risk", "Change_Collision_Risk"]
    display_names = {
        "Recurrence_Risk": "Recurrence",
        "MTTR_Drag_Risk": "MTTR Drag", 
        "Execution_Churn_Risk": "Exec Churn",
        "Reopen_Churn_Risk": "Reopen",
        "Change_Collision_Risk": "Change"
    }
    
    available = [c for c in risk_cols if c in df.columns]
    if not available:
        return None
    
    try:
        hm = df.head(8).copy()
        hm["Display"] = hm["Service"].astype(str).str[:14] + " (" + hm["Service_Tier"].astype(str) + ")"
        hm = hm.set_index("Display")[available]
        hm = hm.rename(columns=display_names)
        hm = hm.apply(pd.to_numeric, errors="coerce").fillna(0)
        
        fig, ax = plt.subplots(figsize=(8.0, 5.5))
        
        im = ax.imshow(hm.values, aspect='auto', cmap='RdYlGn_r', vmin=0, vmax=100)
        
        ax.set_xticks(np.arange(len(hm.columns)))
        ax.set_yticks(np.arange(len(hm.index)))
        ax.set_xticklabels(hm.columns, fontsize=10, fontweight='bold', color='#0F172A')
        ax.set_yticklabels(hm.index, fontsize=10, color='#334155')
        
        cbar = plt.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
        cbar.set_label('Risk Score', fontsize=10, fontweight='bold', color='#0F172A')
        cbar.ax.tick_params(labelsize=9, colors='#64748B')
        
        for i in range(len(hm.index)):
            for j in range(len(hm.columns)):
                val = int(round(hm.iloc[i, j], 0))
                color = 'white' if val > 60 else '#0F172A'
                ax.text(j, i, str(val), ha='center', va='center', 
                       fontsize=10, color=color, fontweight='bold')
        
        ax.set_title('Service Risk Concentration', fontsize=13, fontweight='bold', 
                    color='#0F172A', pad=15)
        
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

def _build_pareto_image(df: pd.DataFrame) -> Optional[io.BytesIO]:
    if df.empty:
        return None
    try:
        fig, ax1 = plt.subplots(figsize=(7.0, 4.5), dpi=120)
        
        labels = [str(x)[:22] + "..." if len(str(x)) > 22 else str(x) for x in df["Theme"]]
        x_pos = np.arange(len(df))
        
        ax1.bar(x_pos, df["Frequency"], color="#3B82F6", width=0.55)
        ax1.set_ylabel("Frequency of Root Cause", color="#0F172A", fontweight="bold", fontsize=9)
        ax1.tick_params(axis="y", labelcolor="#0F172A")
        
        ax1.set_xticks(x_pos)
        ax1.set_xticklabels(labels, rotation=35, ha="right", fontsize=8)

        ax2 = ax1.twinx()
        ax2.plot(x_pos, df["Cumulative_Pct"], color="#DC2626", marker="o", linewidth=2.5)
        ax2.set_ylabel("Cumulative Impact Percentage", color="#DC2626", fontweight="bold", fontsize=9)
        ax2.set_ylim(0, 110)
        
        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        plt.title("Eighty Twenty Rule: Top Structural Risk Themes", fontweight="bold", color="#0F172A", pad=15)
        
        plt.gcf().subplots_adjust(bottom=0.35)
        plt.tight_layout()
        
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Pareto error: {e}")
        return None

def _build_impact_matrix_image(service_risk_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if service_risk_df.empty or "Active_Disruption_P1_P2" not in service_risk_df.columns:
        return None
    try:
        merged = service_risk_df.sort_values(by=["Active_Disruption_P1_P2", "Total_Service_Risk"], ascending=[False, False]).head(5).copy()
        
        merged["Active_Disruption_P1_P2"] = pd.to_numeric(merged["Active_Disruption_P1_P2"], errors="coerce").fillna(0)
        merged["Recurrence_Risk"] = pd.to_numeric(merged["Recurrence_Risk"], errors="coerce").fillna(0)
            
        fig, ax1 = plt.subplots(figsize=(7.5, 4.5), dpi=120)
        
        x_pos = np.arange(len(merged))
        labels = [textwrap.fill(str(x)[:20], width=12) for x in merged["Service"]]
        
        bars = ax1.bar(x_pos, merged["Active_Disruption_P1_P2"], color="#DC2626", width=0.45, alpha=0.9, label="Active Disruption (P1/P2 Count)")
        ax1.set_ylabel("Active Disruption Volume (P1 and P2)", color="#DC2626", fontweight="bold", fontsize=9)
        ax1.tick_params(axis="y", labelcolor="#DC2626")
        ax1.set_xticks(x_pos)
        ax1.set_xticklabels(labels, rotation=0, ha="center", fontsize=8, fontweight="bold")
        
        for bar in bars:
            yval = bar.get_height()
            if yval > 0:
                ax1.text(bar.get_x() + bar.get_width()/2.0, yval + 0.05, f"{int(yval)}", ha='center', va='bottom', color='#DC2626', fontweight='bold', fontsize=9)
        
        max_disruption = float(merged["Active_Disruption_P1_P2"].max())
        if max_disruption < 5:
            ax1.set_ylim(0, 5)
        else:
            ax1.set_ylim(0, max_disruption * 1.3)

        ax2 = ax1.twinx()
        ax2.plot(x_pos, merged["Recurrence_Risk"], color="#0F172A", marker="o", linewidth=2.5, markersize=8, label="Recurrence Risk Score")
        ax2.set_ylabel("Recurrence Risk Score (Zero to 100)", color="#0F172A", fontweight="bold", fontsize=9)
        ax2.tick_params(axis="y", labelcolor="#0F172A")
        ax2.set_ylim(0, 115)
        
        for i, val in enumerate(merged["Recurrence_Risk"]):
            ax2.text(x_pos[i], val + 3, f"{int(val)}", ha='center', va='bottom', color='#0F172A', fontweight='bold', fontsize=8)

        ax1.spines['top'].set_visible(False)
        ax2.spines['top'].set_visible(False)
        
        plt.title("Executive Strike Zone: Top 5 Services (Disruption vs Recurrence)", fontweight="bold", color="#0F172A", pad=15)
        
        lines_1, labels_1 = ax1.get_legend_handles_labels()
        lines_2, labels_2 = ax2.get_legend_handles_labels()
        ax1.legend(lines_1 + lines_2, labels_1 + labels_2, loc='upper center', bbox_to_anchor=(0.5, -0.15), frameon=False, ncol=2, fontsize=8)
        
        plt.gcf().subplots_adjust(bottom=0.25)
        plt.tight_layout()
        
        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Impact matrix error: {e}")
        return None

def _build_macro_trend_image(history_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if history_df.empty or len(history_df) <= 1:
        return None
    try:
        df = history_df.copy()
        df["display_date"] = pd.to_datetime(df["run_date"]).dt.strftime('%m/%Y')

        fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=120)

        ax.plot(df["display_date"], df["bvsi_score"], marker='o', linewidth=3.0, color='#0F172A', label='Global Stability (BVSI™)')
        ax.plot(df["display_date"], df["debt_score"], marker='s', linewidth=2.0, color='#DC2626', linestyle='--', label='Structural Risk Debt™')

        ax.set_ylim(0, 110)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.set_ylabel("Index Score", fontweight='bold', color='#0F172A', fontsize=9)

        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha='center', fontsize=8, fontweight='bold', color='#334155')
        plt.yticks(fontsize=9, color='#334155')

        ax.set_title("Macro Trajectory: Stability vs. Debt", fontweight='bold', color='#0F172A', pad=15, fontsize=12)

        ax.legend(frameon=False, loc='upper center', bbox_to_anchor=(0.5, -0.28), ncol=2, fontsize=9)
        plt.tight_layout(rect=[0, 0.12, 1, 1])

        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Macro trend error: {e}")
        return None

def _build_micro_trend_image(history_df: pd.DataFrame) -> Optional[io.BytesIO]:
    if history_df.empty or len(history_df) <= 1:
        return None
    try:
        df = history_df.copy()
        df["display_date"] = pd.to_datetime(df["run_date"]).dt.strftime('%m/%Y')

        fig, ax = plt.subplots(figsize=(6.2, 3.2), dpi=120)

        ax.plot(df["display_date"], df["resilience_score"], marker='^', linewidth=2.0, color='#2563EB', alpha=0.8, label='Resilience')
        ax.plot(df["display_date"], df["governance_score"], marker='d', linewidth=2.0, color='#059669', alpha=0.8, label='Governance')
        ax.plot(df["display_date"], df["momentum_score"], marker='*', linewidth=2.0, color='#D97706', alpha=0.8, label='Momentum')
        ax.plot(df["display_date"], df["debt_score"], marker='s', linewidth=2.0, color='#DC2626', linestyle=':', alpha=0.6, label='Risk Debt')

        ax.set_ylim(0, 110)
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)

        plt.setp(ax.xaxis.get_majorticklabels(), rotation=0, ha='center', fontsize=8, fontweight='bold', color='#334155')
        plt.yticks(fontsize=9, color='#334155')

        ax.set_title("Diagnostic Trajectory: Domain Breakdown", fontweight='bold', color='#0F172A', pad=15, fontsize=12)

        ax.legend(frameon=False, loc='upper center', bbox_to_anchor=(0.5, -0.28), ncol=4, fontsize=9)
        plt.tight_layout(rect=[0, 0.12, 1, 1])

        img = io.BytesIO()
        plt.savefig(img, format='png', bbox_inches='tight', facecolor='white', edgecolor='none')
        plt.close(fig)
        img.seek(0)
        return img
    except Exception as e:
        print(f"Micro trend error: {e}")
        return None

def _build_domain_insight_box(domain_scores: dict, styles: Any) -> Table:
    weakest_domain = "Service Resilience"
    if domain_scores:
        lowest_val = 101.0
        for key, val in domain_scores.items():
            if float(val) < lowest_val:
                lowest_val = float(val)
                weakest_domain = key

    if "Resilience" in weakest_domain:
        insight = "<b>Cost of Inaction:</b> Prolonged recovery times, execution churn, and high reopen rates actively degrade business productivity and erode end user trust. Engineering teams are trapped in reactive firefighting.<br/><br/><b>Mandate:</b> Automate runbook execution and enforce strict incident closure criteria."
    elif "Governance" in weakest_domain:
        insight = "<b>Cost of Inaction:</b> High collision rates expose the enterprise to self inflicted outages, directly threatening revenue generating hours. Current release controls are failing to predict impact.<br/><br/><b>Mandate:</b> Institute mandatory peer reviews and freeze non emergency deployments for high risk services."
    elif "Debt" in weakest_domain:
        insight = "<b>Cost of Inaction:</b> Recurring issues are closed without root cause remediation, creating a compounding tax on IT capacity. This hidden debt prevents scale and innovation.<br/><br/><b>Mandate:</b> Shift engineering capacity from feature development directly to technical debt paydown."
    else:
        insight = "<b>Cost of Inaction:</b> The overall trajectory of operational stability is decaying. Despite existing controls, systemic friction is increasing across the portfolio.<br/><br/><b>Mandate:</b> Execute a holistic, top down review of the service management lifecycle to reverse this trend."

    title = Paragraph(f"Primary Exposure: {weakest_domain}", styles["TableHeader"])
    body = Paragraph(insight, styles["TableCell"])

    box = Table([[title], [body]], colWidths=[3.9*inch])
    box.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#D97706")), 
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#FFFBEB")), 
        ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#D97706")),
        ('LEFTPADDING', (0, 0), (-1, -1), 10),
        ('RIGHTPADDING', (0, 0), (-1, -1), 10),
        ('TOPPADDING', (0, 0), (-1, -1), 8),
        ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
    ]))
    return box

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
        automation_df = _safe_df(payload.get("automation_df"))
        trust_gap_df = _safe_df(payload.get("trust_gap_df"))
        rca_themes_df = _safe_df(payload.get("rca_themes_df"))
        rca_pareto_df = _safe_df(payload.get("rca_pareto_df"))
        history_df = _safe_df(payload.get("history_df"))
        detected_dataset = str(payload.get("detected_dataset", "INCIDENT")).upper()
        service_anchor = str(payload.get("service_anchor_used", "Service"))
        readiness = _safe_float(payload.get("data_readiness_score", 0))

        story = []

        story.append(Paragraph("OSIL™ Executive Briefing", styles["ExecutiveTitle"]))
        story.append(Paragraph(f"Operational Stability Intelligence Report for {tenant_name} | As of {as_of}", 
                              styles["ExecutiveSubtitle"]))
        
        metric_data = [
            ["BVSI™ Score", "Operating Posture", "Data Quality"],
            [f"{bvsi:.1f}", posture, f"{readiness:.0f}%"]
        ]
        
        metric_table = Table(metric_data, colWidths=[2.2*inch, 3.3*inch, 1.5*inch])
        metric_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, 1), 16),
            ('BOTTOMPADDING', (0, 1), (-1, 1), 16),
            ('FONTNAME', (0, 1), (0, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 1), (0, 1), 22),
            ('TEXTCOLOR', (0, 1), (0, 1), colors.HexColor("#2563EB") if bvsi >= 70 else colors.HexColor("#D97706") if bvsi >= 55 else colors.HexColor("#DC2626")),
            ('FONTNAME', (1, 1), (1, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (1, 1), (1, 1), 14),
            ('FONTNAME', (2, 1), (2, 1), 'Helvetica-Bold'),
            ('FONTSIZE', (2, 1), (2, 1), 16),
            ('GRID', (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
        ]))
        story.append(metric_table)
        story.append(Spacer(1, 24))

        story.append(Paragraph("Executive Signal", styles["SectionHeader"]))
        signal = f"Current operational posture is {posture} (BVSI™ {bvsi:.1f}). "
        signal += "Governance mechanisms are active, but concentrated instability in specific service domains requires targeted executive attention and execution."
        story.append(Paragraph(signal, styles["ExecutiveBody"]))
        story.append(Spacer(1, 20))
        
        story.append(Paragraph("Strategic Imperatives", styles["SectionHeader"]))
        
        strongest = max(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Change Governance"
        weakest = min(domain_scores.items(), key=lambda x: float(x[1]))[0] if domain_scores else "Structural Risk Debt"
        
        imp_data = [
            [Paragraph("Executive Insight", styles["TableHeader"]), 
             Paragraph("Leadership Action", styles["TableHeader"])],
            [Paragraph(f"<b>Strength:</b> {strongest}", styles["TableCell"]),
             Paragraph("Leverage mature controls as the enterprise standard and expand successful patterns.", styles["TableCell"])],
            [Paragraph(f"<b>Primary Exposure:</b> {weakest}", styles["TableCell"]),
             Paragraph("Authorize immediate stabilization resourcing for services in the bottom quartile.", styles["TableCell"])],
        ]
        
        imp_table = Table(imp_data, colWidths=[2.0*inch, 5.0*inch])
        imp_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'TOP'),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 14),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 14),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ('LINEBELOW', (0, -1), (-1, -1), 2, colors.HexColor("#0F172A")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(imp_data))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(1, len(imp_data), 2)
        ]))
        
        story.append(imp_table)
        story.append(Spacer(1, 20))
        story.append(Paragraph(f"Data Sources: {detected_dataset} | Classification: Confidential", styles["Caption"]))

        story.append(PageBreak())

        # EXECUTIVE TREND INTELLIGENCE SECTION
        if not history_df.empty and len(history_df) > 1:
            story.append(Paragraph("Executive Trend Intelligence", styles["PageHeader"]))
            
            narrative = (
                "Mathematical proof of execution. The Macro Trajectory tracks the absolute success of the operation, "
                "proving that business value stability is rising while operational friction collapses. The Diagnostic "
                "Trajectory isolates the specific underlying domains driving that overarching success."
            )
            story.append(Paragraph(narrative, styles["ExecutiveBody"]))
            story.append(Spacer(1, 12))
            
            macro_img = _build_macro_trend_image(history_df)
            if macro_img:
                story.append(Image(macro_img, width=6.2*inch, height=3.2*inch))
                story.append(Spacer(1, 16))
                
            micro_img = _build_micro_trend_image(history_df)
            if micro_img:
                story.append(Image(micro_img, width=6.2*inch, height=3.2*inch))
                story.append(Spacer(1, 16))

            story.append(PageBreak())

        # DUAL COLUMN LAYOUT FOR STABILITY PROFILE
        story.append(Paragraph("Operational Stability Profile", styles["PageHeader"]))
        story.append(Spacer(1, 12))
        
        radar_img = _build_radar_chart(domain_scores)
        
        domain_rows = [[Paragraph("Domain", styles["TableHeader"]), 
                        Paragraph("Score", styles["TableHeader"]), 
                        Paragraph("Assessment", styles["TableHeader"])]]
        
        domain_keys = ["Service Resilience", "Change Governance", "Structural Risk Debt™", "Reliability Momentum"]
        for key in domain_keys:
            score = _safe_float(domain_scores.get(key, domain_scores.get(key.replace("™", ""), 0)), 0)
            assessment = "Strong" if score >= 80 else "Improving" if score >= 60 else "At Risk" if score >= 40 else "Critical"
            
            score_color = colors.HexColor("#059669") if score >= 70 else colors.HexColor("#D97706") if score >= 55 else colors.HexColor("#DC2626")
            
            domain_rows.append([
                Paragraph(key, styles["TableCell"]),
                Paragraph(f"{score:.1f}", ParagraphStyle(name='score', parent=styles["TableCell"], textColor=score_color, fontName='Helvetica-Bold')),
                Paragraph(assessment, styles["TableCellBold"])
            ])
        
        domain_table = Table(domain_rows, colWidths=[2.0*inch, 0.6*inch, 1.3*inch])
        domain_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('ALIGN', (0, 0), (0, -1), 'LEFT'),
            ('ALIGN', (1, 0), (1, -1), 'CENTER'),
            ('ALIGN', (2, 0), (2, -1), 'CENTER'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 6),
            ('RIGHTPADDING', (0, 0), (-1, -1), 6),
            ('TOPPADDING', (0, 0), (-1, -1), 8),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(domain_rows))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(domain_rows), 2)
        ]))
        
        right_col_elements = [
            Paragraph("Domain Scorecard", styles["SectionHeader"]),
            domain_table,
            Spacer(1, 8),
            Paragraph("Scoring: 80 to 100 Strong | 60 to 79 Controlled | <60 Weakness", styles["Caption"]),
            Spacer(1, 16),
            _build_domain_insight_box(domain_scores, styles)
        ]

        if radar_img:
            radar_flowable = Image(radar_img, width=3.1*inch, height=2.7*inch)
            layout_table = Table([[radar_flowable, right_col_elements]], colWidths=[3.1*inch, 3.9*inch])
            layout_table.setStyle(TableStyle([
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 0),
                ('RIGHTPADDING', (0, 0), (-1, -1), 0),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 0),
                ('TOPPADDING', (0, 0), (-1, -1), 0),
            ]))
            story.append(KeepTogether([layout_table]))
        else:
            story.append(KeepTogether(right_col_elements))

        story.append(PageBreak())

        story.append(Paragraph("Service Improvement Priorities & Confidence Lens", styles["PageHeader"]))
        
        lens_text = (
            "<b>The Executive Confidence Lens:</b> A high instability score is actionable only if underlying data is accurate. "
            "The <b>Data Confidence</b> metric grades telemetry hygiene. <br/>"
            "<b>High:</b> Trust the risk score entirely. <br/>"
            "<b>Medium/Low:</b> The instability signal is clouded by poor ticket logging. The immediate mandate is enforcing data hygiene, not technical rewrites."
        )
        
        box_data = [[Paragraph(lens_text, styles["ExecutiveBody"])]]
        box_table = Table(box_data, colWidths=[7.0*inch])
        box_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
            ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#0F172A")),
            ('LEFTPADDING', (0, 0), (-1, -1), 12),
            ('RIGHTPADDING', (0, 0), (-1, -1), 12),
            ('TOPPADDING', (0, 0), (-1, -1), 12),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
        ]))
        story.append(box_table)
        story.append(Spacer(1, 16))
        
        if not sip_candidates.empty:
            top3_data = [[Paragraph("Service", styles["TableHeader"]), 
                         Paragraph("Focus Area", styles["TableHeader"]),
                         Paragraph("Data Conf.", styles["TableHeader"]), 
                         Paragraph("Mandate", styles["TableHeader"])]]
            
            for idx, (_, row) in enumerate(sip_candidates.head(3).iterrows(), 1):
                svc = str(row.get("Service", "Unknown"))
                focus = str(row.get("Primary_Focus", str(row.get("Why_Flagged", "Risk"))))
                confidence = str(row.get("Data_Hygiene_Check", str(row.get("Data_Confidence", "Unknown"))))
                action = str(row.get("Mandate", "Activate SIP" if idx == 1 else "Monitor"))
                
                top3_data.append([
                    Paragraph(svc, styles["TableCellBold"]),
                    Paragraph(focus, styles["TableCell"]),
                    Paragraph(confidence, ParagraphStyle(name='center', parent=styles["TableCell"], alignment=1)),
                    Paragraph(action, ParagraphStyle(name='center', parent=styles["TableCellBold"], alignment=1))
                ])
            
            top3_table = Table(top3_data, colWidths=[2.2*inch, 2.6*inch, 1.0*inch, 1.2*inch])
            top3_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 10),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(top3_data))
            ] + [
                ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#FEF3C7"))
            ]))
            
            story.append(KeepTogether([
                Paragraph("Immediate Executive Actions", styles["SectionHeader"]),
                top3_table
            ]))
            
            story.append(Spacer(1, 24))
            
            sip_data = [[Paragraph("Service", styles["TableHeader"]), 
                        Paragraph("Tier", styles["TableHeader"]),
                        Paragraph("Focus / Theme", styles["TableHeader"]),
                        Paragraph("Data Conf.", styles["TableHeader"]),
                        Paragraph("Action", styles["TableHeader"])]]
            
            for _, row in sip_candidates.iterrows():
                svc = str(row.get("Service", ""))
                tier = str(row.get("Service_Tier", "N/A"))
                theme = str(row.get("Primary_Focus", row.get("Suggested_Theme", "Stability")))
                confidence = str(row.get("Data_Hygiene_Check", "Unknown"))
                action = str(row.get("Mandate", row.get("Priority_Label", "Monitor")))
                
                sip_data.append([
                    Paragraph(svc, styles["TableCellBold"]),
                    Paragraph(tier, ParagraphStyle(name='c1', parent=styles["TableCell"], alignment=1)),
                    Paragraph(theme, ParagraphStyle(name='c2', parent=styles["TableCell"], alignment=1)),
                    Paragraph(confidence, ParagraphStyle(name='c3', parent=styles["TableCellBold"], alignment=1)),
                    Paragraph(action, ParagraphStyle(name='c4', parent=styles["TableCellBold"], alignment=1))
                ])
            
            sip_table = Table(sip_data, colWidths=[1.8*inch, 0.6*inch, 2.3*inch, 1.1*inch, 1.2*inch])
            sip_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(sip_data))
            ]))
            
            story.append(KeepTogether([
                Paragraph("Full SIP Portfolio", styles["SectionHeader"]),
                sip_table
            ]))

        story.append(PageBreak())
        
        story.append(Paragraph("Automation Strike Zone (AIOps Readiness)", styles["PageHeader"]))
        
        if not automation_df.empty:
            story.append(Paragraph("The OSIL engine has identified specific workflows causing excessive manual toil. Deploy automated remediation scripts and orchestration workflows to eliminate this operational friction.", styles["ExecutiveBody"]))
            
            auto_data = [[Paragraph("Target Workflow", styles["TableHeader"]), Paragraph("Type", styles["TableHeader"]), Paragraph("Wasted Hours", styles["TableHeader"])]]
            for _, row in automation_df.iterrows():
                auto_data.append([
                    Paragraph(str(row.get("Target_Service", "")), styles["TableCellBold"]),
                    Paragraph(str(row.get("Automation_Type", "")), styles["TableCell"]),
                    Paragraph(str(row.get("Wasted_Hours", "")), styles["TableCellBold"])
                ])
            auto_table = Table(auto_data, colWidths=[3.0*inch, 2.5*inch, 1.5*inch])
            auto_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#DC2626")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(auto_data))
            ]))
            story.append(auto_table)
        else:
            story.append(Paragraph("Insufficient Channel or Service Request data to calculate automation deficits. Map these intake sources to reveal manual effort waste.", styles["ExecutiveBody"]))
            
        story.append(PageBreak()) 
        
        story.append(Paragraph("Business Trust & Root Cause Analytics", styles["PageHeader"]))
        story.append(Paragraph("The Xentrixus OSIL™ framework identifies structural operational gaps by measuring the ratio of silent friction to active disruption, and by extracting true thematic root causes rather than relying on administrative closure flags.", styles["ExecutiveBody"]))
        story.append(Spacer(1, 16))

        impact_img = _build_impact_matrix_image(service_risk_df)
        if impact_img:
            impact_elements = []
            impact_elements.append(Image(impact_img, width=6.5*inch, height=4.2*inch))
            impact_elements.append(Spacer(1, 12))
            
            impact_narrative = (
                "<b>Executive Insight:</b> Services showing high disruption volumes (red bars) combined with elevated recurrence risk (black lines) represent immediate executive danger zones. "
                "These services require immediate executive sponsorship and capital allocation to stop the bleeding."
            )
            
            box_data = [[Paragraph(impact_narrative, styles["ExecutiveBody"])]]
            box_table = Table(box_data, colWidths=[6.5*inch])
            box_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FEF2F2")),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#FCA5A5")),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ('RIGHTPADDING', (0, 0), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ]))
            
            impact_elements.append(box_table)
            story.append(KeepTogether(impact_elements))
            story.append(Spacer(1, 24))
        else:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Insufficient Data for Executive Strike Zone", styles["SectionHeader"]))
            story.append(Paragraph("The engine requires distinct Priority mapping separating high urgency events from low urgency friction. Without this, the system cannot calculate the business trust disruption ratio.", styles["ExecutiveBody"]))
            story.append(Spacer(1, 24))

        if not trust_gap_df.empty:
            trust_data = [[Paragraph("Service", styles["TableHeader"]), 
                           Paragraph("Critical Disruption<br/>(P1 and P2)", styles["TableHeader"]),
                           Paragraph("Silent Friction<br/>(P3 to P5)", styles["TableHeader"]), 
                           Paragraph("Friction Ratio", styles["TableHeader"])]]
            
            for _, row in trust_gap_df.iterrows():
                svc = str(row.get("Service", ""))
                crit = str(row.get("Active_Disruption_P1_P2", "0"))
                fric = str(row.get("Silent_Friction_P3_P5", "0"))
                ratio = f"{_safe_float(row.get('Friction_Ratio', 0)):.1f}x"
                
                trust_data.append([
                    Paragraph(svc, styles["TableCellBold"]),
                    Paragraph(crit, ParagraphStyle(name='t1', parent=styles["TableCellBold"], alignment=1)),
                    Paragraph(fric, ParagraphStyle(name='t2', parent=styles["TableCell"], alignment=1)),
                    Paragraph(ratio, ParagraphStyle(name='t3', parent=styles["TableCell"], alignment=1))
                ])
                
            trust_table = Table(trust_data, colWidths=[3.0*inch, 1.5*inch, 1.5*inch, 1.0*inch])
            trust_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (0, -1), 'LEFT'),
                ('ALIGN', (1, 0), (-1, -1), 'CENTER'),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(trust_data))
            ] + [
                ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(trust_data), 2)
            ]))
            
            story.append(KeepTogether([
                Paragraph("Xentrixus Trust Gap Matrix (P1 to P5)", styles["SectionHeader"]),
                trust_table
            ]))
            
        story.append(PageBreak())
        
        story.append(Paragraph("Structural Risk Debt™: Thematic Extraction", styles["PageHeader"]))
        
        pareto_img = _build_pareto_image(rca_pareto_df)
        if pareto_img:
            pareto_elements = []
            pareto_elements.append(Image(pareto_img, width=6.5*inch, height=4.2*inch))
            pareto_elements.append(Spacer(1, 12))
            
            pareto_narrative = (
                "<b>Executive Insight:</b> The Pareto principle dictates that roughly eighty percent of operational instability stems from twenty percent of the underlying causes. "
                "The themes clustered on the left side of this axis represent your highest leverage remediation targets. "
                "Funding service improvement programs focused exclusively on these top drivers will mathematically eliminate the vast majority of your structural risk debt while optimizing resource allocation."
            )
            
            box_data = [[Paragraph(pareto_narrative, styles["ExecutiveBody"])]]
            box_table = Table(box_data, colWidths=[6.5*inch])
            box_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#F8FAFC")),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E1")),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ('RIGHTPADDING', (0, 0), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ]))
            
            pareto_elements.append(box_table)
            story.append(KeepTogether(pareto_elements))
            story.append(Spacer(1, 24))
        else:
            story.append(Spacer(1, 12))
            story.append(Paragraph("Insufficient Data for Pareto Analysis", styles["SectionHeader"]))
            
            error_narrative = (
                "<b>Structural Data Void Detected.</b> The engine could not detect structured root cause themes. "
                "This indicates a severe maturity gap in problem management. Operational teams are likely closing tickets without documenting thematic resolution. "
                "Relying purely on boolean checkboxes hides the true drivers of structural risk debt."
            )
            
            box_data = [[Paragraph(error_narrative, styles["ExecutiveBody"])]]
            box_table = Table(box_data, colWidths=[6.5*inch])
            box_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, -1), colors.HexColor("#FEF2F2")),
                ('BOX', (0, 0), (-1, -1), 1, colors.HexColor("#FCA5A5")),
                ('LEFTPADDING', (0, 0), (-1, -1), 12),
                ('RIGHTPADDING', (0, 0), (-1, -1), 12),
                ('TOPPADDING', (0, 0), (-1, -1), 12),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 12),
            ]))
            story.append(box_table)
            story.append(Spacer(1, 24))

        if not rca_themes_df.empty:
            rca_data = [[Paragraph("Service", styles["TableHeader"]), 
                         Paragraph("Documented Root Cause Themes", styles["TableHeader"])]]
            
            for _, row in rca_themes_df.iterrows():
                svc = str(row.get("Service", ""))
                themes = str(row.get("Documented_Themes", ""))
                
                rca_data.append([
                    Paragraph(svc, styles["TableCellBold"]),
                    Paragraph(themes, styles["TableCell"])
                ])
                
            rca_table = Table(rca_data, colWidths=[2.5*inch, 4.0*inch])
            rca_table.setStyle(TableStyle([
                ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
                ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
                ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
                ('VALIGN', (0, 0), (-1, -1), 'TOP'),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                ('TOPPADDING', (0, 0), (-1, -1), 8),
                ('BOTTOMPADDING', (0, 0), (-1, -1), 8),
                ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
            ] + [
                ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(rca_data))
            ] + [
                ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(rca_data), 2)
            ]))
            
            story.append(KeepTogether([
                Paragraph("Structural Risk Debt™: Root Cause Ledger", styles["SectionHeader"]),
                rca_table
            ]))
        else:
            story.append(Paragraph("Structural Risk Debt™: Root Cause Ledger", styles["SectionHeader"]))
            story.append(Paragraph("No thematic root cause data was detected in the active problem dataset. Ensure the Root Cause Description field is mapped and actively utilized by operational teams.", styles["ExecutiveBody"]))


        story.append(PageBreak())

        story.append(Paragraph("Service Risk Concentration Matrix", styles["PageHeader"]))
        story.append(Paragraph("Multidimensional risk analysis. Dark red indicates critical risk concentration requiring immediate executive intervention.", styles["ExecutiveBody"]))
        story.append(Spacer(1, 12))
        
        hm_img = _build_heatmap(service_risk_df)
        if hm_img:
            story.append(Image(hm_img, width=7.0*inch, height=4.8*inch))
            story.append(Spacer(1, 20))
        else:
            story.append(Paragraph("Insufficient data to generate heatmap. Please ensure Service and Service Tier columns are accurately mapped.", styles["ExecutiveBody"]))
            story.append(Spacer(1, 20))
        
        story.append(KeepTogether([
            Paragraph("Risk Dimension Definitions", styles["SectionHeader"]),
            Paragraph(
                "<b>Recurrence:</b> Frequency of repeated incidents indicating unresolved root causes.<br/>"
                "<b>MTTR Drag:</b> Recovery time exceeds target levels.<br/>"
                "<b>Execution Churn:</b> Frequency of ticket reassignment (ping-pong effect) indicating broken routing.<br/>"
                "<b>Reopen:</b> Incidents closed prematurely.<br/>"
                "<b>Change:</b> Instability correlated with deployment windows.", 
                styles["ExecutiveBody"]
            )
        ]))

        story.append(PageBreak())

        story.append(Paragraph("Appendix: Reference Methodology", styles["PageHeader"]))
        story.append(Spacer(1, 16))
        
        scale_data = [
            ["Range", "Posture", "Executive Action"],
            ["85 to 100", "High Confidence", "Invest in scaling operations; stability enables growth"],
            ["70 to 84", "Controlled and Improving", "Maintain course; selective SIPs in weak domains"],
            ["55 to 69", "Controlled but Exposed", "Immediate SIP activation in Structural Risk Debt"],
            ["40 to 54", "Reactive and Exposed", "Executive intervention required; halt perimeter changes"],
            ["Below 40", "Fragile Operations", "Emergency stabilization; executive crisis protocol"],
        ]
        
        scale_table = Table(scale_data, colWidths=[1.2*inch, 2.0*inch, 3.8*inch])
        scale_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(scale_data))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(scale_data), 2)
        ]))
        
        story.append(KeepTogether([
            Paragraph("BVSI™ Interpretation Framework", styles["SectionHeader"]),
            scale_table
        ]))
        
        story.append(Spacer(1, 24))
        
        risk_data = [
            ["Score", "Level", "Response", "Investment"],
            ["0 to 30", "Green or Stable", "Standard operations, monitoring", "Maintenance"],
            ["31 to 60", "Yellow or Elevated", "Operational review in thirty days", "Preventive"],
            ["61 to 80", "Orange or High", "Executive sponsor, SIP activation", "Priority"],
            ["81 to 100", "Red or Critical", "Crisis protocol, freeze deployments", "Emergency"],
        ]
        
        risk_table = Table(risk_data, colWidths=[1.2*inch, 1.6*inch, 2.5*inch, 1.7*inch])
        risk_table.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#0F172A")),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ('LEFTPADDING', (0, 0), (-1, -1), 10),
            ('TOPPADDING', (0, 0), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 0), (-1, -1), 10),
            ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#0F172A")),
        ] + [
            ('LINEBELOW', (0, i), (-1, i), 0.5, colors.HexColor("#E2E8F0")) for i in range(1, len(risk_data))
        ] + [
            ('BACKGROUND', (0, i), (-1, i), colors.HexColor("#F8FAFC")) for i in range(2, len(risk_data), 2)
        ]))
        
        story.append(KeepTogether([
            Paragraph("Risk Score Calibration", styles["SectionHeader"]),
            risk_table
        ]))
        
        story.append(PageBreak()) 
        
        about_text = (
            f"The Operational Stability Intelligence Layer (OSIL™) is an enterprise advisory platform developed exclusively by Xentrixus. "
            f"OSIL synthesizes complex IT service data to reveal structural vulnerabilities and operational friction. "
            f"The framework is anchored by the Business Value Stability Index (BVSI™). "
            f"BVSI translates technical metrics into a unified executive score that measures the absolute capability of an organization to sustain operational resilience and protect core business value. "
            f"This intelligence brief is derived from {detected_dataset} telemetry using {service_anchor} as the foundational perspective."
        )
        
        story.append(KeepTogether([
            Paragraph("About This Report", styles["SectionHeader"]),
            Paragraph(about_text, styles["ExecutiveBody"])
        ]))

        doc.build(story, onFirstPage=_footer, onLaterPages=_footer)
        out.seek(0)
        return out.getvalue()
        
    except Exception as e:
        import traceback
        raise Exception(f"PDF Generation Failed: {str(e)}\n{traceback.format_exc()}")
