"""
KKC Audit Sampling Tool – Report Generator
Produces a KKC-branded PDF (ReportLab) documenting the sampling process,
and calls the Anthropic API to generate an AI-assisted work paper narrative.
"""

from __future__ import annotations
import io
from datetime import datetime
from typing import Dict, Optional

from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.colors import HexColor, white, black, Color
from reportlab.lib.units import cm
from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import (
    SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
    HRFlowable, KeepTogether,
)

# ─── KKC BRAND COLOURS ────────────────────────────────────────────────────────
KKC_GREEN  = HexColor("#7CB542")
KKC_GREY   = HexColor("#808285")
KKC_LGREY  = HexColor("#F5F5F5")
KKC_LGREEN = HexColor("#EEF7E3")
KKC_RED    = HexColor("#dc3545")
KKC_GREEN2 = HexColor("#D4EDDA")
KKC_RED2   = HexColor("#F8D7DA")


def _styles():
    base = getSampleStyleSheet()
    styles = {}

    styles["title"] = ParagraphStyle(
        "KKCTitle", parent=base["Normal"],
        fontSize=20, textColor=white, fontName="Helvetica-Bold", leading=26,
    )
    styles["subtitle"] = ParagraphStyle(
        "KKCSubtitle", parent=base["Normal"],
        fontSize=11, textColor=white, fontName="Helvetica", leading=16,
    )
    styles["section"] = ParagraphStyle(
        "KKCSection", parent=base["Normal"],
        fontSize=12, textColor=KKC_GREEN, fontName="Helvetica-Bold", leading=18,
        spaceBefore=12, spaceAfter=4,
    )
    styles["body"] = ParagraphStyle(
        "KKCBody", parent=base["Normal"],
        fontSize=9.5, textColor=black, fontName="Helvetica", leading=14,
    )
    styles["bold"] = ParagraphStyle(
        "KKCBold", parent=base["Normal"],
        fontSize=9.5, textColor=black, fontName="Helvetica-Bold", leading=14,
    )
    styles["small"] = ParagraphStyle(
        "KKCSmall", parent=base["Normal"],
        fontSize=8, textColor=KKC_GREY, fontName="Helvetica", leading=12,
        alignment=TA_CENTER,
    )
    styles["conclusion_ok"] = ParagraphStyle(
        "KKCConclOK", parent=base["Normal"],
        fontSize=10, textColor=HexColor("#155724"), fontName="Helvetica-Bold",
        leading=14,
    )
    styles["conclusion_fail"] = ParagraphStyle(
        "KKCConclFail", parent=base["Normal"],
        fontSize=10, textColor=HexColor("#721c24"), fontName="Helvetica-Bold",
        leading=14,
    )
    return styles


def _table_style_base():
    return TableStyle([
        ("FONTNAME",    (0, 0), (-1, 0),  "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("BACKGROUND",  (0, 0), (-1,  0), KKC_GREEN),
        ("TEXTCOLOR",   (0, 0), (-1,  0), white),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [KKC_LGREEN, white]),
        ("GRID",        (0, 0), (-1, -1), 0.25, KKC_GREY),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0,0), (-1, -1), 4),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("VALIGN",      (0, 0), (-1, -1), "MIDDLE"),
    ])


# ─── PUBLIC: PDF GENERATOR ────────────────────────────────────────────────────

def generate_pdf_report(report_data: Dict) -> bytes:
    """
    Build and return a KKC-branded PDF documenting the sampling process.

    report_data keys:
      engagement  : dict  (client_name, audit_area, financial_year,
                            prepared_by, reviewed_by, date)
      sampling    : dict  (sampling_type, population, selection_method)
      tod         : dict  (optional) all TOD parameters
      toc         : dict  (optional) all TOC parameters
      projection  : dict  (optional) projected misstatement results
      ai_narrative: str   (optional) AI-generated work paper text
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=2 * cm, rightMargin=2 * cm,
        topMargin=1.5 * cm, bottomMargin=2 * cm,
        title="KKC Audit Sampling Documentation",
        author="KKC & Associates LLP",
    )

    S      = _styles()
    W      = 17 * cm   # usable page width
    elems  = []

    # ── HEADER ────────────────────────────────────────────────────────────────
    header = Table(
        [[Paragraph("KKC &amp; Associates LLP", S["title"])],
         [Paragraph("Audit Sampling Documentation — Internal Work Paper", S["subtitle"])]],
        colWidths=[W],
    )
    header.setStyle(TableStyle([
        ("BACKGROUND",   (0, 0), (-1, 0), KKC_GREEN),
        ("BACKGROUND",   (0, 1), (-1, 1), KKC_GREY),
        ("TOPPADDING",   (0, 0), (-1, -1), 9),
        ("BOTTOMPADDING",(0, 0), (-1, -1), 9),
        ("LEFTPADDING",  (0, 0), (-1, -1), 14),
    ]))
    elems.append(header)
    elems.append(Spacer(1, 0.4 * cm))

    # ── ENGAGEMENT DETAILS ────────────────────────────────────────────────────
    eng = report_data.get("engagement", {})
    eng_data = [
        ["Client Name",    eng.get("client_name", "—"),
         "Prepared By",    eng.get("prepared_by", "—")],
        ["Audit Area",     eng.get("audit_area", "—"),
         "Date",           eng.get("date", "—")],
        ["Financial Year", eng.get("financial_year", "—"),
         "Reviewed By",    eng.get("reviewed_by", "—")],
    ]
    eng_table = Table(eng_data, colWidths=[4 * cm, 4.5 * cm, 4 * cm, 4.5 * cm])
    eng_table.setStyle(TableStyle([
        ("FONTNAME",    (0, 0), (0, -1), "Helvetica-Bold"),
        ("FONTNAME",    (2, 0), (2, -1), "Helvetica-Bold"),
        ("FONTSIZE",    (0, 0), (-1, -1), 9),
        ("TOPPADDING",  (0, 0), (-1, -1), 4),
        ("BOTTOMPADDING",(0,0), (-1, -1), 4),
        ("LINEBELOW",   (0, -1), (-1, -1), 0.5, KKC_GREY),
    ]))
    elems.append(eng_table)
    elems.append(Spacer(1, 0.3 * cm))

    # ── SECTION 1: SAMPLING OVERVIEW ──────────────────────────────────────────
    elems.append(Paragraph("1.  Sampling Overview", S["section"]))
    sampling = report_data.get("sampling", {})
    ov_data = [
        ["Parameter", "Value"],
        ["Sampling Type",        sampling.get("sampling_type", "—")],
        ["Population Amount",    f"\u20b9 {sampling.get('population', 0):,.2f}"],
        ["Sample Selection Method", sampling.get("selection_method", "—")],
        ["Applicable Standard",  "SA 530 – Audit Sampling (ICAI)"],
        ["Firm Reference",       "KKC Sampling Guide – Section 26"],
    ]
    ov_table = Table(ov_data, colWidths=[7 * cm, 10 * cm])
    ov_table.setStyle(_table_style_base())
    elems.append(ov_table)
    elems.append(Spacer(1, 0.3 * cm))

    # ── SECTION 2: TOD ────────────────────────────────────────────────────────
    tod = report_data.get("tod")
    if tod:
        elems.append(Paragraph("2.  Test of Details – Sample Size Determination (Poisson MUS)", S["section"]))

        tod_rows = [
            ["Parameter", "Value Selected", "Impact on Sample Size"],
            ["Performance Materiality",
             f"\u20b9 {tod.get('pm', 0):,.2f}", "Sets Tolerable Misstatement Rate"],
            ["Tolerable Misstatement Rate",
             f"{tod.get('tm_rate', 0):.4%}", "PM ÷ Population"],
            ["Expected Error Rate",
             tod.get("expected_error", "—"), "Uplift applied to base n"],
            ["Entity Level Risk",
             tod.get("entity_risk", "—"), "Determines base confidence level"],
            ["Assertion Risk",
             tod.get("assertion_risk", "—"), "Primary driver of base confidence"],
            ["Controls Response",
             tod.get("controls_response", "—"), "Reduces n if controls effective"],
            ["SADA Response",
             tod.get("sada_response", "—"), "Significant reduction if High"],
            ["SAP Response",
             tod.get("sap_response", "—"), "Moderate reduction if High"],
            ["Reliability Factor (R)",
             f"{tod.get('reliability_factor', 0):.4f}", "= –ln(1 – Confidence)"],
            ["Required Confidence Level",
             f"{tod.get('confidence', 0):.0%}", "Derived from all parameters above"],
            ["Calculated Sample Size",
             str(tod.get("sample_size", "—")), "MINIMUM — cannot be reduced"],
        ]
        tod_table = Table(tod_rows, colWidths=[6 * cm, 4.5 * cm, 6.5 * cm])
        ts = _table_style_base()
        ts.add("FONTNAME",   (0, -1), (-1, -1), "Helvetica-Bold")
        ts.add("BACKGROUND", (0, -1), (-1, -1), KKC_GREEN2)
        tod_table.setStyle(ts)
        elems.append(tod_table)
        elems.append(Spacer(1, 0.3 * cm))

    # ── SECTION 3: TOC ────────────────────────────────────────────────────────
    toc = report_data.get("toc")
    if toc:
        sec_num = 3 if tod else 2
        elems.append(Paragraph(f"{sec_num}.  Test of Controls – Minimum Sample Size (Appendix V)", S["section"]))
        toc_rows = [
            ["Parameter", "Value"],
            ["Control Frequency",    toc.get("frequency", "—")],
            ["Assertion Risk Level", toc.get("risk_level", "—")],
            ["Minimum Sample Size",  str(toc.get("min_sample", "—"))],
            ["Firm Reference",       "KKC Sampling Guide – Appendix V"],
        ]
        toc_table = Table(toc_rows, colWidths=[7 * cm, 10 * cm])
        ts2 = _table_style_base()
        ts2.add("FONTNAME",   (0, -2), (-1, -2), "Helvetica-Bold")
        ts2.add("BACKGROUND", (0, -2), (-1, -2), KKC_GREEN2)
        toc_table.setStyle(ts2)
        elems.append(toc_table)
        elems.append(Spacer(1, 0.3 * cm))

    # ── SECTION: AI NARRATIVE ─────────────────────────────────────────────────
    narrative = report_data.get("ai_narrative")
    if narrative:
        next_sec = (4 if (tod and toc) else 3 if (tod or toc) else 2)
        elems.append(Paragraph(f"{next_sec}.  Process Explanation (AI-Assisted Work Paper Narrative)", S["section"]))
        elems.append(Paragraph(narrative.replace("\n", "<br/>"), S["body"]))
        elems.append(Spacer(1, 0.3 * cm))
        next_sec += 1
    else:
        next_sec = (4 if (tod and toc) else 3 if (tod or toc) else 2)

    # ── SECTION: PROJECTED MISSTATEMENT ───────────────────────────────────────
    proj = report_data.get("projection")
    if proj:
        elems.append(Paragraph(f"{next_sec}.  Projected Misstatement Summary (Section 26.12–26.13)", S["section"]))
        is_ok = proj.get("is_acceptable", True)
        proj_rows = [
            ["Component", "Amount (₹)"],
            ["Projected Misstatement (non-anomalous)",
             f"\u20b9 {proj.get('projected', 0):,.2f}"],
            ["Anomalous Misstatement (actual, per Section 26.11)",
             f"\u20b9 {proj.get('anomalous', 0):,.2f}"],
            ["Total Misstatement (Projected + Anomalous)",
             f"\u20b9 {proj.get('total', 0):,.2f}"],
            ["Tolerable Misstatement",
             f"\u20b9 {proj.get('tolerable', 0):,.2f}"],
        ]
        proj_table = Table(proj_rows, colWidths=[11 * cm, 6 * cm])
        ts3 = _table_style_base()
        ts3.add("FONTNAME",   (0, -2), (-1, -1), "Helvetica-Bold")
        ts3.add("BACKGROUND", (0, -2), (-1, -2), HexColor("#FFF3CD"))
        ts3.add("BACKGROUND", (0, -1), (-1, -1), HexColor("#FFF3CD"))
        proj_table.setStyle(ts3)
        elems.append(proj_table)
        elems.append(Spacer(1, 0.2 * cm))

        concl_style = S["conclusion_ok"] if is_ok else S["conclusion_fail"]
        concl_bg    = KKC_GREEN2 if is_ok else KKC_RED2
        concl_text  = proj.get("conclusion", "")
        concl_box   = Table([[Paragraph(concl_text, concl_style)]],
                            colWidths=[W])
        concl_box.setStyle(TableStyle([
            ("BACKGROUND",   (0, 0), (-1, -1), concl_bg),
            ("TOPPADDING",   (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING",(0, 0), (-1, -1), 8),
            ("LEFTPADDING",  (0, 0), (-1, -1), 12),
            ("BOX",          (0, 0), (-1, -1), 0.5,
             KKC_GREEN if is_ok else KKC_RED),
        ]))
        elems.append(concl_box)

    # ── FOOTER ────────────────────────────────────────────────────────────────
    elems.append(Spacer(1, 0.8 * cm))
    elems.append(HRFlowable(width="100%", thickness=0.5, color=KKC_GREY))
    elems.append(Spacer(1, 0.15 * cm))
    elems.append(Paragraph(
        f"KKC &amp; Associates LLP | Chartered Accountants | Mumbai · Pune · Bengaluru · Ahmedabad &nbsp;|&nbsp; "
        f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')} &nbsp;|&nbsp; "
        "For internal audit documentation purposes only",
        S["small"],
    ))

    doc.build(elems)
    return buf.getvalue()


# ─── PUBLIC: AI NARRATIVE ─────────────────────────────────────────────────────

def generate_ai_narrative(report_data: Dict, api_key: str) -> str:
    """
    Call the Anthropic API to produce a professional work paper narrative
    explaining the sampling process carried out.
    Returns plain text (no markdown) suitable for inclusion in the PDF.
    """
    try:
        import anthropic  # local import to keep reportlab dependency separate

        client = anthropic.Anthropic(api_key=api_key)

        eng      = report_data.get("engagement", {})
        sampling = report_data.get("sampling", {})
        tod      = report_data.get("tod", {})
        toc      = report_data.get("toc", {})
        proj     = report_data.get("projection", {})

        prompt = f"""You are a Senior Chartered Accountant at KKC & Associates LLP, Mumbai, documenting an audit sampling procedure for a statutory audit work paper. Write in a formal, professional tone consistent with ICAI standards (SA 530 – Audit Sampling).

ENGAGEMENT CONTEXT
Client          : {eng.get('client_name', 'Not specified')}
Audit Area      : {eng.get('audit_area', 'Not specified')}
Financial Year  : {eng.get('financial_year', 'Not specified')}
Sampling Type   : {sampling.get('sampling_type', 'Not specified')}
Population (₹)  : {sampling.get('population', 0):,.2f}
Selection Method: {sampling.get('selection_method', 'Not specified')}

{"TOD PARAMETERS" if tod else ""}
{("Performance Materiality: ₹" + f"{tod.get('pm',0):,.2f}") if tod else ""}
{("Entity Risk: " + tod.get('entity_risk','')) if tod else ""}
{("Assertion Risk: " + tod.get('assertion_risk','')) if tod else ""}
{("Controls Response: " + tod.get('controls_response','')) if tod else ""}
{("SADA Response: " + tod.get('sada_response','')) if tod else ""}
{("SAP Response: " + tod.get('sap_response','')) if tod else ""}
{("Confidence Level: " + f"{tod.get('confidence',0):.0%}") if tod else ""}
{("Calculated Sample Size: " + str(tod.get('sample_size',''))) if tod else ""}

{"TOC PARAMETERS" if toc else ""}
{("Control Frequency: " + toc.get('frequency','')) if toc else ""}
{("Assertion Risk: " + toc.get('risk_level','')) if toc else ""}
{("Minimum Sample Size: " + str(toc.get('min_sample',''))) if toc else ""}

{"PROJECTION RESULTS" if proj else ""}
{("Projected Misstatement: ₹" + f"{proj.get('projected',0):,.2f}") if proj else ""}
{("Total Misstatement: ₹" + f"{proj.get('total',0):,.2f}") if proj else ""}
{("Tolerable Misstatement: ₹" + f"{proj.get('tolerable',0):,.2f}") if proj else ""}
{("Conclusion: " + proj.get('conclusion','')) if proj else ""}

Write EXACTLY four paragraphs (no headings, no bullet points, no markdown):
1. Purpose of sampling in this context and why this methodology was appropriate.
2. How the sample size was determined — the specific risk parameters, confidence level derived, and Poisson reliability factor logic.
3. How sample items were selected, the selection method rationale, and any stratification applied.
4. Conclusion on audit sufficiency based on projection results (or note that projection is pending if no results provided).

Write in plain paragraphs. Do NOT use any formatting symbols."""

        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[{"role": "user", "content": prompt}],
        )
        return response.content[0].text.strip()

    except Exception as exc:
        return (
            f"AI-assisted narrative could not be generated at this time. "
            f"Reason: {exc}. "
            "Please document the sampling rationale manually in accordance with "
            "Section 26.4 and 26.6 of the KKC Sampling Guide."
        )
