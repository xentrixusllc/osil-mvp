# Replace the "Strategic Imperatives" section in the previous code with this:

# Strategic Imperatives - Fixed Layout
story.append(Paragraph("Strategic Imperatives", styles["SectionHeader"]))

strongest = max(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Change Governance"
weakest = min(domain_scores.items(), key=lambda x: _safe_float(x[1], 0))[0] if domain_scores else "Structural Risk Debt"
top_svc = str(sip_candidates.iloc[0].get("Service", "Critical Service")) if not sip_candidates.empty else "Priority Service"

# Simplified, executive-grade text (shened for clarity)
imp_data = [
    ["Strategic Insight", "Board Action"],
    [
        f"<b>Operational Strength:</b><br/>{strongest} demonstrates mature controls and should be leveraged as the operational standard for other domains.", 
        f"Codify {strongest} practices into enterprise standards. Expand successful patterns to underperforming domains."
    ],
    [
        f"<b>Critical Exposure:</b><br/>{weakest} represents the highest concentration of stability risk across the service portfolio.", 
        f"Authorize immediate SIP funding for {top_svc} and peer services in the bottom quartile."
    ],
    [
        f"<b>Investment Priority:</b><br/>{top_svc} exhibits top-quartile instability requiring executive intervention.", 
        "Assign executive sponsor immediately. Initiate 30-day remediation sprint with weekly board updates."
    ]
]

# Wider columns with more breathing room
imp_table = Table(imp_data, colWidths=[3.4*inch, 3.1*inch])
imp_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1E293B")),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 10),
    ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
    ('VALIGN', (0, 0), (-1, -1), 'TOP'),  # Top align for multi-line text
    ('LEFTPADDING', (0, 0), (-1, -1), 12),
    ('RIGHTPADDING', (0, 0), (-1, -1), 12),
    ('TOPPADDING', (0, 0), (-1, 0), 10),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
    ('TOPPADDING', (0, 1), (-1, -1), 10),  # More vertical space for content
    ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
    ('LINEBELOW', (0, 0), (-1, 0), 2, colors.HexColor("#1E293B")),
    ('LINEBELOW', (0, 1), (-1, 1), 0.5, colors.HexColor("#E2E8F0")),
    ('LINEBELOW', (0, 2), (-1, 2), 0.5, colors.HexColor("#E2E8F0")),
    ('LINEBELOW', (0, 3), (-1, 3), 2, colors.HexColor("#1E293B")),
    # Alternating row backgrounds for readability
    ('BACKGROUND', (0, 1), (-1, 1), colors.HexColor("#F8FAFC")),
]))

# Keep header with table to prevent splits
story.append(KeepTogether([
    Paragraph("Strategic Imperatives", styles["SectionHeader"]),
    Spacer(1, 8),
    imp_table
]))

story.append(Spacer(1, 20))
story.append(Paragraph(f"<i>Data Sources: {detected_dataset} | Classification: Confidential</i>", styles["Caption"]))
