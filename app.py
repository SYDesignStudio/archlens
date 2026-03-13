
import os
import time
import uuid
import tempfile
from io import BytesIO
from typing import Dict, List, Tuple

import fitz
import streamlit as st
import pdf_summary
from docx import Document
from docx.enum.table import WD_ALIGN_VERTICAL, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt

DEFAULT_STATE = {
    "report": None,
    "sections": None,
    "word_file": None,
    "pdf_file": None,
    "last_filename": None,
    "last_error": None,
    "report_id": None,
    "active_module": "Building Regulations Review",
    "planning_statement_type": "Planning Statement",
}
for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value

MAX_FILE_SIZE_MB = 50
MAX_PAGE_COUNT = 120

BUILDING_REQUIRED_HEADINGS = [
    "PROJECT CLASSIFICATION",
    "PROJECT DETAILS",
    "TOP SUMMARY",
    "DRAWING-PACK INCONSISTENCIES",
    "EXECUTIVE SUMMARY",
    "DRAWING PACK SUMMARY",
    "COMPLIANCE STATUS BY APPROVED DOCUMENT",
    "KEY RISKS",
    "MISSING INFORMATION",
    "RECOMMENDED ACTIONS",
    "BUILDING CONTROL SUBMISSION READINESS",
]

PLANNING_REQUIRED_HEADINGS = [
    "PROJECT CLASSIFICATION",
    "SITE AND PROPOSAL OVERVIEW",
    "TOP SUMMARY",
    "LOCAL AUTHORITY CONTEXT",
    "PD / PRIOR APPROVAL / PLANNING ROUTE",
    "PLANNING OFFICER STYLE REASONING",
    "DRAWING-PACK INCONSISTENCIES",
    "KEY RISKS",
    "MISSING INFORMATION",
    "RECOMMENDED ACTIONS",
    "SUBMISSION READINESS",
]

BUILDING_SECTION_ORDER = [
    ("PROJECT CLASSIFICATION", "Project Classification"),
    ("PROJECT DETAILS", "Project Details"),
    ("TOP SUMMARY", "Top Summary"),
    ("DRAWING-PACK INCONSISTENCIES", "Drawing-Pack Inconsistencies"),
    ("EXECUTIVE SUMMARY", "Executive Summary"),
    ("DRAWING PACK SUMMARY", "Drawing Pack Summary"),
    ("COMPLIANCE STATUS BY APPROVED DOCUMENT", "Compliance Status by Approved Document"),
    ("KEY RISKS", "Key Risks"),
    ("MISSING INFORMATION", "Missing Information"),
    ("RECOMMENDED ACTIONS", "Recommended Actions"),
    ("BUILDING CONTROL SUBMISSION READINESS", "Building Control Submission Readiness"),
]

PLANNING_SECTION_ORDER = [
    ("PROJECT CLASSIFICATION", "Project Classification"),
    ("SITE AND PROPOSAL OVERVIEW", "Site and Proposal Overview"),
    ("TOP SUMMARY", "Top Summary"),
    ("LOCAL AUTHORITY CONTEXT", "Local Authority Context"),
    ("PD / PRIOR APPROVAL / PLANNING ROUTE", "PD / Prior Approval / Planning Route"),
    ("PLANNING OFFICER STYLE REASONING", "Planning Officer Style Reasoning"),
    ("DRAWING-PACK INCONSISTENCIES", "Drawing-Pack Inconsistencies"),
    ("KEY RISKS", "Key Risks"),
    ("MISSING INFORMATION", "Missing Information"),
    ("RECOMMENDED ACTIONS", "Recommended Actions"),
    ("SUBMISSION READINESS", "Submission Readiness"),
]

BUILDING_SPECIAL_KEY_VALUE_SECTIONS = {
    "PROJECT CLASSIFICATION",
    "TOP SUMMARY",
    "BUILDING CONTROL SUBMISSION READINESS",
}

PLANNING_SPECIAL_KEY_VALUE_SECTIONS = {
    "PROJECT CLASSIFICATION",
    "TOP SUMMARY",
    "LOCAL AUTHORITY CONTEXT",
    "PD / PRIOR APPROVAL / PLANNING ROUTE",
    "SUBMISSION READINESS",
}

BUILDING_DISCLAIMER_TEXT = (
    "Beta: This report is an AI-assisted preliminary review. "
    "It does not replace professional Building Control approval, structural engineering design, "
    "or statutory review by the Local Authority or Approved Inspector."
)

PLANNING_DISCLAIMER_TEXT = (
    "Beta: This report is an AI-assisted preliminary planning review. "
    "It does not replace a formal planning appraisal, lawful development advice, "
    "or the decision of the Local Planning Authority."
)

MODULE_CONFIG = {
    "Building Regulations Review": {
        "required_headings": BUILDING_REQUIRED_HEADINGS,
        "section_order": BUILDING_SECTION_ORDER,
        "special_key_value_sections": BUILDING_SPECIAL_KEY_VALUE_SECTIONS,
        "disclaimer": BUILDING_DISCLAIMER_TEXT,
        "title": "AI Building Regulations Compliance Review",
        "readiness_key": "BUILDING CONTROL SUBMISSION READINESS",
    },
    "Planning Review": {
        "required_headings": PLANNING_REQUIRED_HEADINGS,
        "section_order": PLANNING_SECTION_ORDER,
        "special_key_value_sections": PLANNING_SPECIAL_KEY_VALUE_SECTIONS,
        "disclaimer": PLANNING_DISCLAIMER_TEXT,
        "title": "AI Planning Route and Risk Review",
        "readiness_key": "SUBMISSION READINESS",
    },
}

PROJECT_TYPE_OPTIONS = [
    "Ground Floor Rear Extension",
    "Ground Floor Side Extension",
    "Ground Floor Infill Extension",
    "Porch",
    "First Floor Rear Extension",
    "First Floor Side Extension",
    "Loft Extension",
    "Flat Conversion",
    "House Conversion",
]

PROPERTY_TYPE_OPTIONS = [
    "Not stated",
    "Detached House",
    "Semi-Detached House",
    "Terraced House",
    "End of Terrace House",
    "Bungalow",
    "Chalet Bungalow",
    "Flat",
    "Maisonette",
    "Other",
]

PROPERTY_TYPE_OPTIONS = [
    "Not stated",
    "Detached House",
    "Semi-Detached House",
    "Terraced House",
    "End of Terrace House",
    "Bungalow",
    "Chalet Bungalow",
    "Flat",
    "Maisonette",
    "Other",
]


PLANNING_STATEMENT_TYPES = ["Planning Statement", "Design & Access Statement", "Prior Approval Statement"]


def inject_custom_css():
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
        .sy-hero {
            padding: 1.2rem 1.25rem;
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            background: linear-gradient(135deg, rgba(31,59,115,0.22), rgba(10,14,23,0.85));
            margin-bottom: 1rem;
        }
        .sy-step {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 0.85rem 1rem;
            background: rgba(255,255,255,0.03);
            min-height: 92px;
        }
        .sy-card {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 18px;
            padding: 1rem 1rem 0.85rem 1rem;
            background: rgba(255,255,255,0.03);
            margin-bottom: 0.9rem;
        }
        .sy-mini-card {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 16px;
            padding: 0.85rem 1rem;
            background: rgba(255,255,255,0.03);
            min-height: 146px;
        }
        .sy-kpi {
            font-size: 0.78rem;
            text-transform: uppercase;
            letter-spacing: 0.06em;
            opacity: 0.75;
            margin-bottom: 0.35rem;
        }
        .sy-kpi-value {
            font-size: 1.02rem;
            font-weight: 700;
            line-height: 1.35;
        }
        .sy-upload-item {
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 14px;
            padding: 0.65rem 0.8rem;
            background: rgba(255,255,255,0.02);
            margin-bottom: 0.5rem;
        }
        .sy-muted {opacity: 0.78; font-size: 0.93rem;}
        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
            padding: 0.75rem 0.9rem;
            border-radius: 16px;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def smooth_progress(progress_bar, status_text, start, end, message, duration=0.8):
    steps = max(1, end - start)
    sleep_time = duration / steps
    for value in range(start, end + 1):
        progress_bar.progress(value)
        status_text.text(f"{message} {value}%")
        time.sleep(sleep_time)


def clean_input_value(value, fallback):
    if value is None:
        return fallback
    cleaned = str(value).strip()
    bad_values = {"", "a", "aa", "as", "s", "sa", "test", "xx", "ww", "w", "m", "na", "n/a"}
    if cleaned.lower() in bad_values:
        return fallback
    return cleaned


def parse_key_value_lines(section_text: str) -> List[Tuple[str, str]]:
    rows = []
    for raw_line in section_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("- "):
            line = line[2:].strip()
        if ":" in line:
            label, value = line.split(":", 1)
            rows.append((label.strip(), value.strip()))
        else:
            rows.append(("", line))
    return rows


def parse_compliance_rows(content: str) -> List[Dict[str, str]]:
    rows = []
    lines = [line.strip() for line in content.splitlines() if line.strip()]
    i = 0
    while i < len(lines):
        line = lines[i]
        if line.startswith("- "):
            line = line[2:].strip()
        if line.startswith("Part "):
            why_text = ""
            if ":" in line:
                left, status = line.split(":", 1)
                status = status.strip()
            else:
                left, status = line, ""

            if "–" in left:
                part_code, part_title = left.split("–", 1)
            elif "-" in left:
                part_code, part_title = left.split("-", 1)
            else:
                part_code, part_title = left, ""

            if i + 1 < len(lines):
                next_line = lines[i + 1]
                if next_line.startswith("- "):
                    next_line = next_line[2:].strip()
                if next_line.lower().startswith("why:"):
                    why_text = next_line[4:].strip()
                    i += 1

            rows.append(
                {
                    "part": part_code.replace("Part", "").strip(),
                    "title": part_title.strip(),
                    "status": status,
                    "why": why_text,
                }
            )
        i += 1
    return rows


def parse_report_sections(report_text: str, required_headings: List[str]) -> Dict[str, str]:
    headings = set(required_headings)
    sections: Dict[str, str] = {}
    current_heading = None
    current_lines: List[str] = []

    for line in report_text.splitlines():
        stripped = line.strip()
        if stripped.upper() in headings:
            if current_heading:
                sections[current_heading] = "\n".join(current_lines).strip()
            current_heading = stripped.upper()
            current_lines = []
        else:
            if current_heading:
                current_lines.append(line)

    if current_heading:
        sections[current_heading] = "\n".join(current_lines).strip()
    return sections


def validate_report_headings(report_text: str, required_headings: List[str]) -> Tuple[bool, List[str]]:
    report_upper = report_text.upper()
    missing = [heading for heading in required_headings if heading not in report_upper]
    return len(missing) == 0, missing


def get_pdf_page_count(pdf_path: str) -> int:
    doc = fitz.open(pdf_path)
    try:
        return len(doc)
    finally:
        doc.close()


def set_cell_shading(cell, fill):
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement("w:shd")
    shd.set(qn("w:fill"), fill)
    tc_pr.append(shd)


def set_cell_text(cell, text, bold=False, font_size=9):
    cell.text = ""
    p = cell.paragraphs[0]
    run = p.add_run(text)
    run.bold = bold
    run.font.size = Pt(font_size)


def set_table_col_widths(table, widths_inches):
    for row in table.rows:
        for idx, width in enumerate(widths_inches):
            row.cells[idx].width = Inches(width)


def add_page_number(paragraph):
    run = paragraph.add_run()
    fld_begin = OxmlElement("w:fldChar")
    fld_begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = " PAGE "
    fld_end = OxmlElement("w:fldChar")
    fld_end.set(qn("w:fldCharType"), "end")
    run._r.append(fld_begin)
    run._r.append(instr)
    run._r.append(fld_end)


def add_word_header(section, title_text="ArchLens AI", subtitle_text="AI Review"):
    header = section.header
    para = header.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = para.add_run(title_text + "\n")
    run1.bold = True
    run1.font.size = Pt(11)
    run2 = para.add_run(subtitle_text)
    run2.font.size = Pt(9)


def add_word_footer(section, practice_name="ArchLens AI", report_title="AI Review"):
    footer = section.footer
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run(f"{practice_name} | {report_title} | Page ")
    run.font.size = Pt(9)
    add_page_number(para)


def build_pdf_report(file_name, address, client, date, practice_name, report_id, sections, module_name):
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfbase.pdfmetrics import stringWidth
    from reportlab.pdfgen import canvas

    config = MODULE_CONFIG[module_name]
    section_order = config["section_order"]
    special_key_value_sections = config["special_key_value_sections"]
    disclaimer_text = config["disclaimer"]
    report_title = config["title"]

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)

    width, height = A4
    left_margin = 40
    right_margin = 40
    top_margin = 50
    bottom_margin = 40
    usable_width = width - left_margin - right_margin
    y = height - top_margin

    def wrap_text(text, font_name="Helvetica", font_size=10, max_width=None):
        max_width = max_width or usable_width
        words = text.split()
        lines = []
        current_line = ""
        for word in words:
            test_line = word if not current_line else current_line + " " + word
            if stringWidth(test_line, font_name, font_size) <= max_width:
                current_line = test_line
            else:
                if current_line:
                    lines.append(current_line)
                current_line = word
        if current_line:
            lines.append(current_line)
        return lines

    def draw_page_header():
        c.setStrokeColor(colors.HexColor("#D9E2F3"))
        c.line(left_margin, height - 28, width - right_margin, height - 28)
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 9)
        c.drawString(left_margin, height - 20, practice_name or "ArchLens AI")
        c.setFont("Helvetica", 8)
        c.drawRightString(width - right_margin, height - 20, f"{report_title} | {report_id}")
        c.setFillColor(colors.black)

    def draw_page_footer():
        c.setStrokeColor(colors.HexColor("#D9E2F3"))
        c.line(left_margin, 22, width - right_margin, 22)
        c.setFillColor(colors.grey)
        c.setFont("Helvetica", 8)
        c.drawString(left_margin, 10, "Beta: AI-assisted preliminary review")
        c.drawRightString(width - right_margin, 10, f"Page {c.getPageNumber()}")
        c.setFillColor(colors.black)

    def new_page():
        nonlocal y
        c.showPage()
        y = height - top_margin
        draw_page_header()
        draw_page_footer()

    def draw_section_banner(title):
        nonlocal y
        banner_height = 24
        if y < bottom_margin + banner_height + 10:
            new_page()
        top = y
        bottom = top - banner_height
        c.setFillColor(colors.HexColor("#E9EEF5"))
        c.roundRect(left_margin, bottom, usable_width, banner_height, 4, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 11)
        c.drawString(left_margin + 10, bottom + 7, title)
        c.setFillColor(colors.black)
        y = bottom - 8

    def get_status_color(status):
        upper = status.upper()
        if "PASS" in upper and "PARTLY" not in upper:
            return colors.green
        if "FAIL" in upper:
            return colors.red
        if "PARTLY" in upper or "REVIEW REQUIRED" in upper:
            return colors.orange
        return colors.grey

    def draw_status_badge(x, y_top, status):
        badge_height = 14
        padding_x = 6
        text_width = stringWidth(status, "Helvetica-Bold", 8)
        badge_width = text_width + padding_x * 2
        c.setFillColor(get_status_color(status))
        c.roundRect(x, y_top - badge_height + 2, badge_width, badge_height, 3, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8)
        c.drawString(x + padding_x, y_top - 9, status)
        c.setFillColor(colors.black)

    def draw_part_badge(x, y_top, part_code):
        radius = 8
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.circle(x + radius, y_top - 6, radius, fill=1, stroke=0)
        c.setFillColor(colors.white)
        c.setFont("Helvetica-Bold", 8)
        c.drawCentredString(x + radius, y_top - 9, part_code)
        c.setFillColor(colors.black)

    def draw_compliance_table(content):
        nonlocal y
        rows = parse_compliance_rows(content)
        if not rows:
            c.setFont("Helvetica", 10)
            for wrapped in wrap_text("No compliance status detected."):
                if y < bottom_margin:
                    new_page()
                c.drawString(left_margin, y, wrapped)
                y -= 14
            return

        col_part = left_margin
        col_doc = left_margin + 40
        col_status = left_margin + 250
        col_why = left_margin + 390
        doc_width = 185
        why_width = usable_width - (col_why - left_margin) - 10

        def draw_table_header():
            nonlocal y
            c.setFillColor(colors.HexColor("#EAEFF7"))
            c.rect(left_margin, y - 18, usable_width, 20, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(col_part, y - 5, "Part")
            c.drawString(col_doc, y - 5, "Approved Document")
            c.drawString(col_status, y - 5, "Status")
            c.drawString(col_why, y - 5, "Why")
            y -= 28

        if y < bottom_margin + 60:
            new_page()
        draw_table_header()

        for row in rows:
            doc_lines = wrap_text(row["title"], "Helvetica-Bold", 9, doc_width)
            why_lines = wrap_text(row["why"], "Helvetica", 8, why_width)
            row_height = max(38, 16 + max(len(doc_lines), max(2, len(why_lines))) * 10)

            if y < bottom_margin + row_height + 20:
                new_page()
                draw_table_header()

            row_top = y
            c.setFillColor(colors.whitesmoke)
            c.rect(left_margin, y - row_height + 6, usable_width, row_height, fill=1, stroke=0)
            c.setFillColor(colors.black)

            draw_part_badge(col_part, row_top, row["part"])

            c.setFont("Helvetica-Bold", 9)
            doc_y = row_top - 5
            for wrapped in doc_lines:
                c.drawString(col_doc, doc_y, wrapped)
                doc_y -= 10

            status_text = row["status"]
            status_upper = status_text.upper()
            if "PARTLY DEMONSTRATED" in status_upper:
                status_text = "REVIEW REQUIRED"
            elif "PASS" in status_upper and "PARTLY" not in status_upper:
                status_text = "PASS"
            elif "FAIL" in status_upper:
                status_text = "FAIL"

            draw_status_badge(col_status, row_top, status_text)

            c.setFont("Helvetica", 8)
            why_y = row_top - 5
            for wrapped in why_lines:
                c.drawString(col_why, why_y, wrapped)
                why_y -= 10

            y -= row_height + 8

    c.setFillColor(colors.HexColor("#1F3B73"))
    c.setFont("Helvetica-Bold", 24)
    c.drawCentredString(width / 2, height - 180, practice_name or "ArchLens AI")
    c.setFont("Helvetica-Bold", 18)
    c.drawCentredString(width / 2, height - 215, report_title)
    c.setFillColor(colors.black)
    c.setFont("Helvetica", 11)
    cover_lines = [
        f"Project Address: {address}",
        f"Client: {client}",
        f"Drawing Pack Reviewed: {file_name}",
        f"Date: {date}",
        f"Report ID: {report_id}",
        f"Prepared by: {practice_name or 'ArchLens AI'}",
    ]
    for i, line in enumerate(cover_lines):
        c.drawCentredString(width / 2, height - 300 - (i * 18), line)
    top_summary_rows = parse_key_value_lines(sections.get("TOP SUMMARY", ""))
    c.setFont("Helvetica-Bold", 12)
    c.drawCentredString(width / 2, height - 430, "Client Summary")
    c.setFont("Helvetica", 10)
    for i, (label, value) in enumerate(top_summary_rows[:4]):
        c.drawCentredString(width / 2, height - 452 - (i * 16), f"{label}: {value}" if label else value)
    c.setFillColor(colors.grey)
    c.setFont("Helvetica-Oblique", 9)
    c.drawCentredString(width / 2, 80, disclaimer_text)
    c.showPage()
    y = height - top_margin
    draw_page_header()
    draw_page_footer()

    for key, title in section_order:
        if y < bottom_margin:
            new_page()
        draw_section_banner(title)
        content = sections.get(key, "Not detected")

        if module_name == "Building Regulations Review" and key == "COMPLIANCE STATUS BY APPROVED DOCUMENT":
            draw_compliance_table(content)
            y -= 10
            continue

        if key in special_key_value_sections:
            rows = parse_key_value_lines(content)
            for label, value in rows:
                if label:
                    c.setFont("Helvetica-Bold", 10)
                    for wrapped in wrap_text(f"{label}:"):
                        if y < bottom_margin:
                            new_page()
                        c.drawString(left_margin, y, wrapped)
                        y -= 14
                    c.setFont("Helvetica", 10)
                    for wrapped in wrap_text(value):
                        if y < bottom_margin:
                            new_page()
                        c.drawString(left_margin + 12, y, wrapped)
                        y -= 14
                else:
                    c.setFont("Helvetica", 10)
                    for wrapped in wrap_text(value):
                        if y < bottom_margin:
                            new_page()
                        c.drawString(left_margin, y, wrapped)
                        y -= 14
                y -= 4
        else:
            c.setFont("Helvetica", 10)
            for raw_line in content.splitlines():
                line = raw_line.strip()
                if not line:
                    y -= 6
                    continue
                if line.startswith("- "):
                    line = "• " + line[2:]
                for wrapped in wrap_text(line):
                    if y < bottom_margin:
                        new_page()
                    c.drawString(left_margin, y, wrapped)
                    y -= 14
        y -= 10

    c.save()
    buffer.seek(0)
    return buffer


def build_word_report(file_name, address, client, date, practice_name, report_id, sections, module_name):
    config = MODULE_CONFIG[module_name]
    section_order = config["section_order"]
    special_key_value_sections = config["special_key_value_sections"]
    disclaimer_text = config["disclaimer"]
    report_title = config["title"]

    doc = Document()

    section = doc.sections[0]
    add_word_header(section, practice_name or "ArchLens AI", report_title)
    add_word_footer(section, practice_name or "ArchLens AI", report_title)

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(practice_name or "ArchLens AI")
    run.bold = True
    run.font.size = Pt(20)

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run(report_title)
    run.bold = True
    run.font.size = Pt(16)

    disclaimer = doc.add_paragraph()
    disclaimer.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = disclaimer.add_run(disclaimer_text)
    run.font.size = Pt(10)

    doc.add_paragraph("")
    info = doc.add_paragraph()
    info.alignment = WD_ALIGN_PARAGRAPH.CENTER
    info.add_run(f"Project Address: {address}\n").bold = True
    info.add_run(f"Client: {client}\n")
    info.add_run(f"Drawing Pack Reviewed: {file_name}\n")
    info.add_run(f"Date: {date}\n")
    info.add_run(f"Report ID: {report_id}\n")
    info.add_run(f"Prepared by: {practice_name or 'ArchLens AI'}")

    doc.add_paragraph("")
    summary_heading = doc.add_paragraph()
    summary_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = summary_heading.add_run("Client Summary")
    run.bold = True
    run.font.size = Pt(13)

    for label, value in parse_key_value_lines(sections.get("TOP SUMMARY", "")):
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        if label:
            p.add_run(f"{label}: ").bold = True
            p.add_run(value)
        else:
            p.add_run(value)

    doc.add_page_break()

    for section in doc.sections:
        add_word_header(section, practice_name or "ArchLens AI", report_title)
        add_word_footer(section, practice_name or "ArchLens AI", report_title)

    meta_table = doc.add_table(rows=6, cols=2)
    meta_table.style = "Table Grid"
    meta_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_table_col_widths(meta_table, [2.0, 4.8])
    meta_rows = [
        ("Project Address", str(address)),
        ("Client", str(client)),
        ("Drawing Pack Reviewed", str(file_name)),
        ("Date", str(date)),
        ("Report ID", str(report_id)),
        ("Prepared by", str(practice_name or "ArchLens AI")),
    ]
    for i, (label, value) in enumerate(meta_rows):
        left_cell = meta_table.rows[i].cells[0]
        right_cell = meta_table.rows[i].cells[1]
        set_cell_text(left_cell, label, bold=True, font_size=10)
        set_cell_text(right_cell, value, font_size=10)
        set_cell_shading(left_cell, "EAEFF7")
        left_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        right_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    doc.add_paragraph("")

    for key, title in section_order:
        doc.add_paragraph("")
        heading = doc.add_paragraph()
        heading_run = heading.add_run(title)
        heading_run.bold = True
        heading_run.font.size = Pt(13)
        content = sections.get(key, "Not detected")

        if key in special_key_value_sections:
            rows = parse_key_value_lines(content)
            for label, value in rows:
                if label:
                    p = doc.add_paragraph()
                    p.add_run(f"{label}: ").bold = True
                    p.add_run(value)
                else:
                    doc.add_paragraph(value)

        elif module_name == "Building Regulations Review" and key == "COMPLIANCE STATUS BY APPROVED DOCUMENT":
            rows = parse_compliance_rows(content)
            table = doc.add_table(rows=1, cols=4)
            table.style = "Table Grid"
            set_table_col_widths(table, [0.8, 2.3, 1.6, 3.3])

            headers = table.rows[0].cells
            set_cell_text(headers[0], "Part", bold=True, font_size=9)
            set_cell_text(headers[1], "Approved Document", bold=True, font_size=9)
            set_cell_text(headers[2], "Status", bold=True, font_size=9)
            set_cell_text(headers[3], "Why", bold=True, font_size=9)
            for cell in headers:
                set_cell_shading(cell, "D9E2F3")

            for row in rows:
                cells = table.add_row().cells
                set_cell_text(cells[0], row["part"], bold=True, font_size=9)
                set_cell_text(cells[1], row["title"], bold=True, font_size=9)
                set_cell_text(cells[2], row["status"], bold=True, font_size=9)
                set_cell_text(cells[3], row["why"], font_size=9)
                status_upper = row["status"].upper()
                if "PASS" in status_upper and "PARTLY" not in status_upper:
                    set_cell_shading(cells[2], "C6EFCE")
                elif "FAIL" in status_upper:
                    set_cell_shading(cells[2], "FFC7CE")
                elif "PARTLY" in status_upper or "REVIEW REQUIRED" in status_upper:
                    set_cell_shading(cells[2], "FCE4D6")
        else:
            for line in content.splitlines():
                stripped = line.strip()
                if not stripped:
                    continue
                if stripped.startswith("- "):
                    doc.add_paragraph(stripped[2:], style="List Bullet")
                else:
                    doc.add_paragraph(stripped)

    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


def extract_summary_value(sections: Dict[str, str], module_name: str):
    top_summary_rows = {k.upper(): v for k, v in parse_key_value_lines(sections.get("TOP SUMMARY", "")) if k}
    if module_name == "Planning Review":
        return (
            top_summary_rows.get("OVERALL PLANNING RISK RATING", "Unknown"),
            top_summary_rows.get("LIKELY ROUTE", "Unknown"),
            top_summary_rows.get("LOCAL AUTHORITY USED", "Unknown"),
        )
    return (
        top_summary_rows.get("OVERALL RISK RATING", "Unknown"),
        top_summary_rows.get("SUBMISSION STATUS", "Unknown"),
        top_summary_rows.get("REVIEW CONFIDENCE", "Unknown"),
    )


def render_kpi_cards(sections: Dict[str, str], report_id: str, module_name: str):
    v1, v2, v3 = extract_summary_value(sections, module_name)
    label_2 = "Likely Route" if module_name == "Planning Review" else "Submission Status"
    label_3 = "Local Authority" if module_name == "Planning Review" else "Review Confidence"

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Report ID", report_id)
    c2.metric("Risk Rating", v1)
    c3.metric(label_2, v2)
    c4.metric(label_3, v3)


def extract_summary_values(sections: Dict[str, str], module_name: str):
    top_summary_rows = {k.upper(): v for k, v in parse_key_value_lines(sections.get("TOP SUMMARY", "")) if k}
    if module_name == "Planning Review":
        return {
            "risk": top_summary_rows.get("OVERALL PLANNING RISK RATING", "Unknown"),
            "route": top_summary_rows.get("LIKELY ROUTE", "Unknown"),
            "authority": top_summary_rows.get("LOCAL AUTHORITY USED", "Unknown"),
            "probability": top_summary_rows.get("PLANNING APPROVAL PROBABILITY", "Unknown"),
        }
    return {
        "risk": top_summary_rows.get("OVERALL RISK RATING", "Unknown"),
        "route": top_summary_rows.get("SUBMISSION STATUS", "Unknown"),
        "authority": top_summary_rows.get("REVIEW CONFIDENCE", "Unknown"),
        "probability": "N/A",
    }


def inject_custom_css():
    st.markdown(
        """
        <style>
        .block-container {padding-top: 1.2rem; padding-bottom: 2rem;}
        .sy-hero {
            padding: 1.2rem 1.25rem;
            border: 1px solid rgba(255,255,255,0.08);
            border-radius: 22px;
            background: linear-gradient(135deg, rgba(31,59,115,0.24), rgba(8,11,18,0.92));
            margin-bottom: 1rem;
        }
        .sy-step, .sy-card, .sy-mini-card, .sy-upload-item {
            border: 1px solid rgba(255,255,255,0.08);
            background: rgba(255,255,255,0.03);
        }
        .sy-step { border-radius: 18px; padding: 0.9rem 1rem; min-height: 96px; }
        .sy-card { border-radius: 20px; padding: 1rem 1rem 0.85rem 1rem; margin-bottom: 0.9rem; }
        .sy-mini-card { border-radius: 16px; padding: 0.9rem 1rem; min-height: 148px; }
        .sy-kpi { font-size: 0.78rem; text-transform: uppercase; letter-spacing: 0.06em; opacity: 0.75; margin-bottom: 0.35rem; }
        .sy-upload-item { border-radius: 14px; padding: 0.75rem 0.9rem; margin-bottom: 0.55rem; }
        .sy-muted { opacity: 0.78; font-size: 0.93rem; }
        div[data-testid="stMetric"] {
            background: rgba(255,255,255,0.03);
            border: 1px solid rgba(255,255,255,0.08);
            padding: 0.75rem 0.9rem;
            border-radius: 16px;
        }
        .stDownloadButton button, .stButton button { border-radius: 14px; }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_kpi_cards(sections: Dict[str, str], report_id: str, module_name: str):
    values = extract_summary_values(sections, module_name)
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Report ID", report_id)
    c2.metric("Risk Rating", values["risk"])
    c3.metric("Likely Route" if module_name == "Planning Review" else "Submission Status", values["route"])
    c4.metric("Approval Probability" if module_name == "Planning Review" else "Review Confidence", values["probability"] if module_name == "Planning Review" else values["authority"])


def render_at_a_glance(sections: Dict[str, str], report_id: str, module_name: str):
    config = MODULE_CONFIG[module_name]
    readiness_key = config["readiness_key"]
    middle_key = "PROJECT DETAILS" if module_name == "Building Regulations Review" else "SITE AND PROPOSAL OVERVIEW"
    middle_title = "Project Details" if module_name == "Building Regulations Review" else "Site and Proposal Overview"

    render_kpi_cards(sections, report_id, module_name)
    st.markdown("")

    col1, col2, col3 = st.columns(3)
    with col1:
        st.markdown('<div class="sy-mini-card"><div class="sy-kpi">Project Classification</div></div>', unsafe_allow_html=True)
        st.markdown(sections.get("PROJECT CLASSIFICATION", "Not detected"))
    with col2:
        st.markdown(f'<div class="sy-mini-card"><div class="sy-kpi">{middle_title}</div></div>', unsafe_allow_html=True)
        st.markdown(sections.get(middle_key, "Not detected"))
    with col3:
        st.markdown('<div class="sy-mini-card"><div class="sy-kpi">Submission Readiness</div></div>', unsafe_allow_html=True)
        st.markdown(sections.get(readiness_key, "Not detected"))

    st.markdown("")
    top_summary_rows = {k.upper(): v for k, v in parse_key_value_lines(sections.get("TOP SUMMARY", "")) if k}
    risk_summary = top_summary_rows.get("OVERALL RISK RATING") or top_summary_rows.get("OVERALL PLANNING RISK RATING") or sections.get("TOP SUMMARY", "")
    summary_hint = top_summary_rows.get("PLANNING APPROVAL PROBABILITY") or top_summary_rows.get("LIKELY ROUTE", "Unknown")
    if "HIGH" in str(risk_summary).upper():
        st.error(f"High risk detected | Summary: {summary_hint}")
    elif "MEDIUM" in str(risk_summary).upper():
        st.warning(f"Moderate risk detected | Summary: {summary_hint}")
    else:
        st.success(f"Lower risk indicated | Summary: {summary_hint}")


def render_section_content(content: str, is_key_value: bool):
    if is_key_value:
        rows = parse_key_value_lines(content)
        for label, value in rows:
            if label:
                st.markdown(f"**{label}:** {value}")
            else:
                st.markdown(value)
    else:
        st.markdown(content)


def render_sections(sections: Dict[str, str], report_text: str, module_name: str):
    config = MODULE_CONFIG[module_name]
    for key, title in config["section_order"]:
        content = sections.get(key, "Not detected")
        with st.expander(title, expanded=key in {"TOP SUMMARY", config["readiness_key"]}):
            render_section_content(content, key in config["special_key_value_sections"])
    with st.expander("Show full AI report"):
        st.text(report_text)


def build_simple_word_doc(title: str, body_text: str) -> BytesIO:
    doc = Document()
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(title)
    run.bold = True
    run.font.size = Pt(16)
    doc.add_paragraph("")
    for line in body_text.splitlines():
        stripped = line.strip()
        if not stripped:
            doc.add_paragraph("")
        elif stripped.startswith("- "):
            doc.add_paragraph(stripped[2:], style="List Bullet")
        else:
            doc.add_paragraph(stripped)
    buffer = BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


inject_custom_css()
st.set_page_config(page_title="ArchLens AI", layout="wide")
st.markdown(
    """
    <div class="sy-hero">
        <h1 style="margin:0;">ArchLens AI</h1>
        <div class="sy-muted">AI planning feasibility, planning route, and Building Regulations review for drawing packs, homeowner sketches, and professional reports.</div>
    </div>
    """,
    unsafe_allow_html=True,
)

step1, step2, step3 = st.columns(3)
with step1:
    st.markdown('<div class="sy-step"><strong>Step 1 — Project setup</strong><br><span class="sy-muted">Choose the review module, project type, and report mode.</span></div>', unsafe_allow_html=True)
with step2:
    st.markdown('<div class="sy-step"><strong>Step 2 — Upload drawing pack</strong><br><span class="sy-muted">Upload one or more PDFs such as plans, elevations, sections, or homeowner sketches.</span></div>', unsafe_allow_html=True)
with step3:
    st.markdown('<div class="sy-step"><strong>Step 3 — Generate outputs</strong><br><span class="sy-muted">Review the report cards, expand sections, download reports, and draft a planning statement.</span></div>', unsafe_allow_html=True)

with st.sidebar:
    st.header("Project Setup")
    review_module = st.selectbox(
        "Review Module",
        ["Building Regulations Review", "Planning Review"],
        index=0 if st.session_state.active_module == "Building Regulations Review" else 1,
    )
    project_types = st.multiselect("Project Type", PROJECT_TYPE_OPTIONS, default=[])

    if review_module == "Planning Review":
        property_type = st.selectbox("Property Type", PROPERTY_TYPE_OPTIONS, index=0)
    else:
        property_type = "Not stated"

    proposal_summary = st.text_area(
        "Proposal Description",
        height=110,
        placeholder="Briefly describe the proposal and what the client wants.",
    )

    review_mode = st.selectbox("Report Mode", ["Architect / Professional", "Homeowner Summary"])
    project_address = st.text_input("Project Address")
    local_authority = st.text_input("Local Authority (optional)") if review_module == "Planning Review" else ""

    if review_mode != "Homeowner Summary":
        practice_name = st.text_input("Practice / Company Name (optional)")
    else:
        practice_name = ""

    client_name = st.text_input("Client")
    review_date = st.date_input("Report Date")

    if st.button("Clear Report", key="clear_report_btn"):
        for key, value in DEFAULT_STATE.items():
            st.session_state[key] = value
        st.session_state["planning_statement_text"] = None
        st.session_state["planning_statement_file"] = None
        st.info("Stored report cleared.")

st.session_state.active_module = review_module
config = MODULE_CONFIG[review_module]

setup_tab, upload_tab, report_tab = st.tabs(["Project Setup", "Upload Drawing Pack", "AI Review Report"])

with setup_tab:
    st.markdown(f'<div class="sy-card"><h3 style="margin-top:0;">{config["title"]}</h3><div class="sy-muted">{config["disclaimer"]}</div></div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.2, 1])
    with c1:
        st.markdown("**Current setup**")
        st.write(f"Review module: {review_module}")
        st.write(f"Report mode: {review_mode}")
        st.write(f"Project type: {', '.join(project_types) if project_types else 'Not stated'}")
        if review_module == "Planning Review":
            st.write(f"Property type: {property_type or 'Not stated'}")
            st.write(f"Local authority: {local_authority or 'Auto-detect from drawing/address/client input'}")
        st.write(f"Project address: {project_address or 'Not provided'}")
        st.write(f"Proposal description: {proposal_summary or 'Not provided'}")
        st.write(f"Client: {client_name or 'Not provided'}")
        if review_mode != "Homeowner Summary":
            st.write(f"Practice / Company: {practice_name or 'Not provided'}")
    with c2:
        if review_module == "Planning Review":
            st.info("Use this module for PD checks, prior approval, full planning fallback risk, planning approval probability, and planning statement drafting.")
        else:
            st.info("Use this module for technical Building Regulations review including plans, sections, details, specifications, and structural sheets.")

with upload_tab:
    st.markdown('<div class="sy-card"><h3 style="margin-top:0;">Upload drawing pack</h3><div class="sy-muted">Upload one or more PDFs. Complete drawing packs usually produce more accurate outputs, but homeowner sketches can still support a preliminary feasibility review.</div></div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Upload Drawing PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="drawing_pdf_upload",
    )

    if review_module == "Planning Review" and review_mode == "Homeowner Summary":
        st.info("Homeowners can upload a simple sketch or basic PDF. ArchLens AI will frame the output as a preliminary planning feasibility review, not a formal planning decision.")

    if uploaded_file:
        preview_col, action_col = st.columns([1.5, 1])
        with preview_col:
            st.markdown("**Uploaded files**")
            for file in uploaded_file:
                file_size_mb = round(file.size / (1024 * 1024), 2)
                st.markdown(
                    f'<div class="sy-upload-item"><strong>{file.name}</strong><br><span class="sy-muted">{file_size_mb} MB</span></div>',
                    unsafe_allow_html=True,
                )
        with action_col:
            st.markdown("**Drawing pack summary**")
            st.metric("Files uploaded", len(uploaded_file))
            total_mb = round(sum(f.size for f in uploaded_file) / (1024 * 1024), 2)
            st.metric("Total size", f"{total_mb} MB")
            st.metric("Selected project types", len(project_types))
            run_analysis = st.button(f"Run {review_module}", key="run_review_btn", use_container_width=True)

        if run_analysis:
            progress_bar = st.progress(0)
            status_text = st.empty()
            temp_pdf_path = None
            file = uploaded_file[-1]

            for file in uploaded_file:
                if file.size > MAX_FILE_SIZE_MB * 1024 * 1024:
                    st.error(f"PDF too large. Maximum file size is {MAX_FILE_SIZE_MB}MB.")
                    st.stop()

                with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp_file:
                    tmp_file.write(file.getbuffer())
                    temp_pdf_path = tmp_file.name

            try:
                page_count = get_pdf_page_count(temp_pdf_path)
                if page_count > MAX_PAGE_COUNT:
                    st.error(f"PDF has {page_count} pages. Maximum allowed is {MAX_PAGE_COUNT} pages.")
                    os.remove(temp_pdf_path)
                    st.stop()

                smooth_progress(progress_bar, status_text, 10, 25, "Preparing drawing analysis...", 0.6)

                def update_analysis_progress(current_batch, total_batches):
                    start_pct = 25
                    end_pct = 85
                    progress = start_pct + int((current_batch / max(1, total_batches)) * (end_pct - start_pct))
                    progress_bar.progress(progress)
                    status_text.text(f"Step 3 of 4 — Analyzing drawing pages... Batch {current_batch} of {total_batches} ({progress}%)")

                try:
                    if review_module == "Building Regulations Review":
                        report = pdf_summary.analyze_pdf(
                            temp_pdf_path,
                            client_project_type=("Project type: " + (", ".join(project_types) or "Not stated") + "\nProposal summary: " + (proposal_summary or "Not stated")),
                            review_mode=review_mode,
                            progress_callback=update_analysis_progress,
                        )
                    else:
                        progress_bar.progress(40)
                        status_text.text("Step 3 of 4 — Analyzing planning route and risks... 40%")
                        report = pdf_summary.analyze_planning_pdf(
                            temp_pdf_path,
                            client_project_types=project_types,
                            property_type=property_type,
                            proposal_summary=proposal_summary,
                            project_address=project_address,
                            local_authority=local_authority,
                            review_mode=review_mode,
                        )
                        progress_bar.progress(85)
                        status_text.text("Step 3 of 4 — Planning analysis complete... 85%")
                except Exception as e:
                    msg = str(e).lower()
                    if "insufficient_quota" in msg or "quota" in msg:
                        st.error("OpenAI API quota exceeded. Please add credits in your OpenAI billing dashboard.")
                    elif "rate limit" in msg or "429" in msg:
                        st.error("The AI analysis service is temporarily rate-limited. Please try again shortly.")
                    else:
                        st.error(f"Could not analyze this PDF: {e}")
                    st.stop()

                valid, missing = validate_report_headings(report, config["required_headings"])
                if not valid:
                    st.error(f"AI report validation failed. Missing headings: {', '.join(missing)}")
                    st.stop()

                sections = parse_report_sections(report, config["required_headings"])

                clean_project_address = clean_input_value(project_address, "Not provided")
                clean_client_name = clean_input_value(client_name, "Not provided")
                clean_practice_name = clean_input_value(practice_name, "ArchLens AI")
                report_id = str(uuid.uuid4())[:8].upper()

                smooth_progress(progress_bar, status_text, 85, 95, "Preparing report files...", 0.6)

                word_file = build_word_report(
                    file.name,
                    clean_project_address,
                    clean_client_name,
                    review_date,
                    clean_practice_name,
                    report_id,
                    sections,
                    review_module,
                )

                pdf_file = build_pdf_report(
                    file.name,
                    clean_project_address,
                    clean_client_name,
                    review_date,
                    clean_practice_name,
                    report_id,
                    sections,
                    review_module,
                )

                st.session_state.report = report
                st.session_state.sections = sections
                st.session_state.word_file = word_file
                st.session_state.pdf_file = pdf_file
                st.session_state.last_filename = file.name
                st.session_state.last_error = None
                st.session_state.report_id = report_id
                st.session_state.active_module = review_module
                st.session_state["planning_statement_text"] = None
                st.session_state["planning_statement_file"] = None

                smooth_progress(progress_bar, status_text, 95, 100, "Finalising report...", 0.4)
                status_text.text("Analysis complete. 100%")
                progress_bar.progress(100)
                st.success("Report created successfully. Open the AI Review Report tab.")
            finally:
                if temp_pdf_path:
                    try:
                        os.remove(temp_pdf_path)
                    except Exception:
                        pass
    else:
        st.info("No files uploaded yet. Add one or more PDFs to preview the drawing pack.")

with report_tab:
    if st.session_state.sections and st.session_state.active_module == review_module:
        sections = st.session_state.sections
        report = st.session_state.report
        word_file = st.session_state.word_file
        pdf_file = st.session_state.pdf_file
        report_id = st.session_state.report_id or "N/A"

        panel_title = "Professional report summary"
        panel_note = "Use the cards below for a quick read, then open the collapsible sections for the detailed report."
        if review_module == "Planning Review" and review_mode == "Homeowner Summary":
            panel_title = "Homeowner planning feasibility summary"
            panel_note = "This is a preliminary feasibility-style review based on the uploaded sketch or drawing pack. It is not a formal planning decision."

        st.markdown(f'<div class="sy-card"><h3 style="margin-top:0;">{panel_title}</h3><div class="sy-muted">{panel_note}</div></div>', unsafe_allow_html=True)

        render_at_a_glance(sections, report_id, review_module)
        st.markdown("")
        render_sections(sections, report, review_module)

        base_filename = (st.session_state.last_filename or "drawing_pack").rsplit(".", 1)[0]
        suffix = "Planning" if review_module == "Planning Review" else "BuildingRegs"

        d1, d2 = st.columns(2)
        with d1:
            st.download_button(
                label=("Download Homeowner Feasibility Report (.docx)" if review_module == "Planning Review" and review_mode == "Homeowner Summary" else "Download Professional Report (.docx)"),
                data=word_file,
                file_name=f"{base_filename}_{suffix}_AI_Review_Report.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="download_docx",
                use_container_width=True,
            )
        with d2:
            st.download_button(
                label=("Download Homeowner Feasibility Report (.pdf)" if review_module == "Planning Review" and review_mode == "Homeowner Summary" else "Download Professional Report (.pdf)"),
                data=pdf_file,
                file_name=f"{base_filename}_{suffix}_AI_Review_Report.pdf",
                mime="application/pdf",
                key="download_pdf",
                use_container_width=True,
            )

        if review_module == "Planning Review":
            st.markdown("")
            st.markdown('<div class="sy-card"><h3 style="margin-top:0;">Automatic Planning Statement</h3><div class="sy-muted">Generate a draft planning statement from the ArchLens review and download it as Word.</div></div>', unsafe_allow_html=True)
            statement_type = st.selectbox(
                "Statement Type",
                PLANNING_STATEMENT_TYPES,
                index=PLANNING_STATEMENT_TYPES.index(st.session_state.get("planning_statement_type", "Planning Statement")),
                key="planning_statement_type_select",
            )
            st.session_state["planning_statement_type"] = statement_type

            if st.button(f"Generate {statement_type}", key="generate_planning_statement_btn", use_container_width=True):
                statement_text = pdf_summary.generate_planning_statement(
                    report_text=report,
                    sections=sections,
                    project_address=project_address or "Not provided",
                    client_name=client_name or "Not provided",
                    local_authority=local_authority or "",
                    review_mode=review_mode,
                    statement_type=statement_type,
                )
                st.session_state["planning_statement_text"] = statement_text
                st.session_state["planning_statement_file"] = build_simple_word_doc(f"Draft {statement_type}", statement_text)

            if st.session_state.get("planning_statement_text"):
                with st.expander(f"Show {statement_type.lower()} draft", expanded=False):
                    st.text(st.session_state["planning_statement_text"])
                safe_statement_name = statement_type.replace(" & ", "_").replace(" ", "_")
                st.download_button(
                    label=f"Download {statement_type} (.docx)",
                    data=st.session_state["planning_statement_file"],
                    file_name=f"{base_filename}_{safe_statement_name}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_planning_statement_docx",
                    use_container_width=True,
                )
    else:
        st.info("No report generated yet. Complete the setup, upload the drawing pack, and run the review.")
