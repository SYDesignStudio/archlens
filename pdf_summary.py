import base64
import gc
import os
import re
import time
from typing import Callable, Dict, List, Optional

import fitz
import pdfplumber
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

LIVE_ANALYSIS_MAX_PAGES = 12
IMAGE_BATCH_SIZE = 2
IMAGE_RENDER_SCALE = 1.0

REQUIRED_HEADINGS = [
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
    "PLANNING ASSESSMENT",
    "DRAWING-PACK INCONSISTENCIES",
    "KEY RISKS",
    "MISSING INFORMATION",
    "RECOMMENDED ACTIONS",
    "SUBMISSION READINESS",
]

_client: Optional[OpenAI] = None


def get_openai_client() -> OpenAI:
    global _client
    if _client is None:
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            raise RuntimeError("OPENAI_API_KEY not found. Set it in your environment or .env file.")
        _client = OpenAI(api_key=api_key)
    return _client


def _call_responses_api(model: str, input_payload, retries: int = 3):
    last_error = None
    client = get_openai_client()
    for attempt in range(retries):
        try:
            return client.responses.create(model=model, input=input_payload)
        except Exception as e:
            last_error = e
            msg = str(e).lower()
            if "insufficient_quota" in msg or "quota" in msg:
                raise
            if attempt < retries - 1 and ("rate limit" in msg or "429" in msg or "timeout" in msg):
                time.sleep(2 * (attempt + 1))
                continue
            raise
    raise last_error


def clean_extracted_text(text: str) -> str:
    if not text:
        return ""
    junk_markers = ("%PDF-", "xref", "endxref", "obj", "endobj", "stream", "endstream", "/Type", "/Length")
    cleaned_lines: List[str] = []
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if any(line.startswith(marker) for marker in junk_markers):
            continue
        if len(line) > 240 and sum(ch.isalnum() for ch in line) / max(1, len(line)) < 0.45:
            continue
        cleaned_lines.append(line)
    return re.sub(r"\n{3,}", "\n\n", "\n".join(cleaned_lines)).strip()


def extract_text_from_pdf(pdf_path: str) -> str:
    all_text: List[str] = []
    with pdfplumber.open(pdf_path) as pdf:
        for page_number, page in enumerate(pdf.pages, start=1):
            page_text = clean_extracted_text(page.extract_text() or "")
            if page_text.strip():
                all_text.append(f"\n--- Page {page_number} ---\n{page_text}")
    return "\n".join(all_text)


def detect_sheet_type(page_text: str) -> str:
    text = (page_text or "").upper()

    if (
        "BUILDING REGULATION SPECIFICATION" in text
        or "BUILDING REGULATIONS SPECIFICATION" in text
        or (
            "SPECIFICATION" in text
            and any(
                key in text
                for key in [
                    "VENTILATION",
                    "SMOKE DETECTION",
                    "FIRE RESISTANCE",
                    "CAVITY BARRIERS",
                    "WINDOWS",
                    "DOORS",
                    "THERMAL",
                    "U-VALUE",
                    "INSULATION",
                ]
            )
        )
    ):
        return "BUILDING REGULATION SPECIFICATION"
    if (
        "FIRE PLAN" in text
        or "FIRE PLANS" in text
        or "KEY TO FIRE STRATEGY" in text
        or "FD30" in text
        or "SMOKE DETECTOR" in text
        or "HEAT DETECTOR" in text
    ):
        return "FIRE PLAN"
    if "PROPOSED ROOF PLAN" in text or "ROOF PLAN" in text:
        return "ROOF PLAN"
    if "PROPOSED PLAN" in text or "PROPOSED PLANS" in text:
        return "PROPOSED PLAN"
    if "PROPOSED SECTION" in text or "PROPOSED SECTIONS" in text:
        return "PROPOSED SECTION"
    if "PROPOSED ELEVATION" in text or "PROPOSED ELEVATIONS" in text:
        return "PROPOSED ELEVATION"
    if "SITE PLAN" in text or "PROPOSED SITE" in text or "BLOCK PLAN" in text:
        return "SITE PLAN"
    if "STRUCTURAL" in text and ("PLAN" in text or "DETAIL" in text or "CALC" in text):
        return "STRUCTURAL"
    if "DETAIL" in text or "DETAILS" in text:
        return "DETAILS"
    if "GENERAL NOTES" in text:
        return "GENERAL NOTES"
    if "COVER SHEET" in text or "SHEET LIST" in text or "TITLE SHEET" in text:
        return "COVER SHEET"
    if "EXISTING PLAN" in text or "EXISTING PLANS" in text:
        return "EXISTING PLAN"
    if "EXISTING SECTION" in text or "EXISTING SECTIONS" in text:
        return "EXISTING SECTION"
    if "DEMOLITION" in text or "DEMOLATION" in text:
        return "DEMOLITION"
    return "OTHER"


def extract_text_by_page(pdf_path: str) -> List[Dict[str, str]]:
    pages: List[Dict[str, str]] = []
    doc = fitz.open(pdf_path)
    try:
        max_pages = min(len(doc), LIVE_ANALYSIS_MAX_PAGES)
        for i in range(max_pages):
            page = doc.load_page(i)
            page_text = clean_extracted_text(page.get_text("text") or "")
            lines = [ln.strip() for ln in page_text.splitlines() if ln.strip()]
            first_line = lines[0] if lines else "Untitled sheet"
            pages.append(
                {
                    "page_number": i + 1,
                    "text": page_text,
                    "sheet_type": detect_sheet_type(page_text),
                    "sheet_title": first_line,
                }
            )
            del page
            gc.collect()
    finally:
        doc.close()
    return pages



def render_pdf_page_batch_to_images(pdf_path: str, start_index: int, end_index: int, scale: float = IMAGE_RENDER_SCALE) -> List[str]:
    doc = fitz.open(pdf_path)
    image_paths: List[str] = []
    try:
        for i in range(start_index, min(end_index, len(doc))):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
            image_path = f"temp_page_{os.getpid()}_{i + 1}.png"
            pix.save(image_path)
            image_paths.append(image_path)
            del pix
            gc.collect()
    finally:
        doc.close()
    return image_paths


def chunk_list(items, chunk_size):
    for i in range(0, len(items), chunk_size):
        yield items[i : i + chunk_size]


def image_to_base64(image_path: str) -> str:
    with open(image_path, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")


def detect_local_authority(address: str = "", text: str = "") -> str:
    combined = f"{address}\n{text}".lower()
    authorities = {
        "Hounslow": ["hounslow", "tw3", "tw4", "tw5", "tw7", "tw13", "tw14"],
        "Ealing": ["ealing", "w5", "w7", "ub1", "ub2", "ub5"],
        "Hillingdon": ["hillingdon", "uxbridge", "hayes", "ub3", "ub4", "ub7", "ha4", "ruislip"],
        "Richmond upon Thames": ["richmond", "twickenham", "isleworth", "tw1", "tw2", "tw10", "tw11"],
        "Brent": ["brent", "wembley", "harlesden", "nw10", "ha9"],
        "Barnet": ["barnet", "edgware", "n20", "en4", "nw7"],
        "Enfield": ["enfield", "n13", "n14", "en1", "en2", "en3", "en4", "en8", "berkshire gardens"],
        "Slough": ["slough", "sl1", "sl2", "sl3"],
        "Reading": ["reading", "rg1", "rg2", "rg30", "rg31"],
        "Surrey Heath": ["surrey heath", "camberley", "gu15", "gu16"],
    }
    for authority, needles in authorities.items():
        if any(needle in combined for needle in needles):
            return authority
    return "Not clearly identified"



def extract_project_address(text: str) -> str:
    if not text:
        return ""
    cleaned = text.replace("\r", "\n")
    patterns = [
        r"(?:address|site)\s*[:\-]\s*([^\n]+(?:\n[^\n]+){0,2}?[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})",
        r"(\d+[A-Za-z]?\s+[^\n,]+(?:,\s*[^\n,]+){0,3},?\s*[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2})",
    ]
    for pattern in patterns:
        m = re.search(pattern, cleaned, re.IGNORECASE)
        if m:
            value = " ".join(part.strip(" ,") for part in m.group(1).splitlines() if part.strip())
            value = re.sub(r"\s{2,}", " ", value).strip(" ,")
            if len(value) > 8:
                return value
    return ""

def infer_fire_statement_status(text: str, page_summary: str) -> str:
    combined = f"{text}\n{page_summary}".lower()
    # A fire plan is not the same as a formal fire statement.
    return "submitted" if "fire statement" in combined else "not evident in the drawing pack"



def build_project_summary_from_inputs(project_types_text: str, proposal_summary_text: str, property_type_text: str) -> str:
    base = proposal_summary_text if proposal_summary_text and proposal_summary_text.lower() != "not stated" else project_types_text
    text = (base or "").strip()
    if not text or text.lower() == "not stated":
        return "Residential development works to the host property as shown on the submitted drawings."
    text = re.sub(r"\|", ", ", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,")
    lowered = text.lower()
    if lowered.startswith("proposed "):
        sentence = text[0].upper() + text[1:]
    else:
        sentence = "Proposed " + text[0].lower() + text[1:]
    if not sentence.endswith("."):
        sentence += "."
    return sentence


def infer_application_type(project_types_text: str, proposal_summary_text: str, property_type_text: str) -> str:
    combined = f"{project_types_text} {proposal_summary_text} {property_type_text}".lower()
    if property_type_text.lower() in {"flat", "maisonette"}:
        return "FULL PLANNING"
    if "ground floor rear extension" in combined and not any(term in combined for term in ["side extension", "wraparound", "first floor", "loft", "dormer", "gable"]):
        return "PRIOR APPROVAL"
    if any(term in combined for term in ["loft", "dormer", "gable", "first floor", "conversion", "side extension", "wraparound", "rooflight", "roof light"]):
        return "FULL PLANNING"
    return "FULL PLANNING"


def infer_submission_readiness_from_context(
    application_type: str,
    project_types_text: str,
    proposal_summary_text: str,
    text: str,
    page_summary: str,
) -> tuple[str, str]:
    combined = f"{project_types_text} {proposal_summary_text} {text} {page_summary}".lower()
    if "superseded" in combined:
        return (
            "FURTHER INFORMATION REQUIRED",
            "A current drawing / site plan set should be confirmed before submission because superseded information appears within the pack.",
        )
    if application_type == "PRIOR APPROVAL" and "ground floor rear extension" in combined:
        if ("6m" in combined or "6000" in combined) and ("3m" in combined or "3000" in combined or "4m" in combined or "4000" in combined):
            return (
                "READY TO SUBMIT",
                "The pack appears to provide the key prior approval information typically required for a larger home extension submission, subject to final dimensional confirmation on the submitted drawings.",
            )
        return (
            "LIKELY READY WITH MINOR AMENDMENTS",
            "The proposal appears capable of proceeding by prior approval, but key depth / height information should be confirmed clearly from the original rear wall before submission.",
        )
    return (
        "LIKELY READY WITH MINOR AMENDMENTS",
        "Initial AI assessment based on drawing pack completeness and clarity.",
    )




def infer_officer_recommendation(readiness_status: str, route_text: str, proposal_features: Dict[str, bool]) -> str:
    route_lower = (route_text or "").lower()
    if "further information required" in (readiness_status or "").lower():
        return "FURTHER INFORMATION REQUIRED"
    if proposal_features.get("rear_dormer") and not proposal_features.get("first_floor_extension"):
        if "full planning" in route_lower:
            return "LIKELY APPROVE"
    if "ready to submit" in (readiness_status or "").lower():
        return "LIKELY APPROVE"
    return "LIKELY APPROVE WITH MINOR AMENDMENTS"


def detect_proposal_features(project_types_text: str, proposal_summary_text: str, text: str, page_summary: str) -> Dict[str, bool]:
    combined = f"{project_types_text}\n{proposal_summary_text}\n{text}\n{page_summary}".lower()
    return {
        "gable": any(term in combined for term in ["gable", "hip to gable", "side gable"]),
        "rear_dormer": any(term in combined for term in ["rear dormer", "dormer", "rear roof enlargement"]),
        "front_rooflights": any(term in combined for term in ["front rooflight", "front rooflights", "rooflight", "rooflights"]),
        "single_storey_rear_extension": "ground floor rear extension" in combined or "single-storey rear extension" in combined or "single storey rear extension" in combined,
        "first_floor_extension": (
        "first floor rear extension" in combined
        or "first floor side extension" in combined
        or "two storey rear extension" in combined
        or "second storey extension" in combined
    ),
        "side_extension": "side extension" in combined,
        "wraparound": "wraparound" in combined or "wrap around" in combined,
        "loft_extension": "loft extension" in combined or "loft conversion" in combined or "rear dormer" in combined or "dormer" in combined,
        "flat_or_maisonette": any(term in combined for term in ["flat", "maisonette"]),
    }


def build_detected_proposal_label(features: Dict[str, bool], fallback_project_types: str) -> str:
    labels = []
    if features.get("gable"):
        labels.append("side gable roof extension")
    if features.get("rear_dormer"):
        labels.append("rear dormer")
    if features.get("front_rooflights"):
        labels.append("front rooflights")
    if features.get("single_storey_rear_extension"):
        labels.append("single-storey rear extension")
    if features.get("first_floor_extension"):
        labels.append("first-floor extension")
    if features.get("side_extension") and "side gable roof extension" not in labels:
        labels.append("side extension")
    if features.get("wraparound"):
        labels.append("wraparound form")
    if not labels:
        return fallback_project_types or "residential alterations"
    return ", ".join(labels)


def detect_street_precedent_signal(text: str, page_summary: str) -> str:
    combined = f"{text}\n{page_summary}".lower()
    precedent_terms = [
        "existing 3d", "proposed 3d", "street scene", "front elevation", "rear elevation",
        "terraced", "end-terraced", "end terrace", "mid-terrace", "mid terrace",
        "similar roof extensions", "similar dormers", "surrounding area", "street scene"
    ]
    score = sum(1 for term in precedent_terms if term in combined)
    if score >= 5:
        return "STRONG"
    if score >= 3:
        return "MODERATE"
    return "LIMITED"


def calculate_planning_route_confidence_score(
    application_type: str,
    project_types_text: str,
    proposal_summary_text: str,
    property_type_text: str,
    text: str,
    page_summary: str,
) -> int:
    combined = f"{project_types_text}\n{proposal_summary_text}\n{property_type_text}\n{text}\n{page_summary}".lower()
    score = 52
    if application_type in {"PRIOR APPROVAL", "FULL PLANNING"}:
        score += 8
    if "location plan" in combined:
        score += 5
    if "block plan" in combined:
        score += 5
    if "site plan" in combined:
        score += 4
    if "proposed plans" in combined or "proposed plan" in combined:
        score += 6
    if "proposed elevations" in combined or "proposed elevation" in combined:
        score += 6
    if "proposed sections" in combined or "proposed section" in combined:
        score += 5
    if "3d" in combined or "isometric" in combined:
        score += 3
    if any(term in combined for term in ["ridge", "eaves", "rooflights", "dormer", "gable"]):
        score += 4
    if property_type_text.lower() in {"terraced house", "semi-detached house", "detached house", "end of terrace house"}:
        score += 3
    if "not clearly identified" in combined:
        score -= 8
    if "conservation area" in combined or "article 4" in combined:
        score -= 4
    return max(35, min(96, score))


def analyze_image_batch(image_paths: List[str], text: str, checks: List[str], review_mode: str) -> str:
    audience_hint = (
        "Use concise technical language suitable for architects and Building Control reviewers."
        if review_mode == "Architect / Professional"
        else "Use plain English suitable for homeowners and avoid unnecessary jargon."
    )

    prompt = f"""
You are an AI architectural drawing reviewer for UK residential projects.

{audience_hint}

Detected compliance checks:
{chr(10).join(checks) if checks else "No specific checks detected"}

You are reviewing a batch of drawing pages from a larger PDF pack.

Focus on:
- sheet titles
- plans
- sections
- elevations
- fire strategy
- stairs
- room labels
- dimensions
- windows
- doors
- annotations
- structural references
- ventilation notes
- thermal notes
- drainage notes
- title block / status notes
- sheet numbering and cross references

Return a concise page-batch summary with:
- what these pages contain
- important rooms/spaces
- visible dimensions/levels
- visible fire strategy information
- visible stair information
- visible ventilation information
- visible thermal / build-up information
- visible structural / steel information
- visible drainage information
- visible drawing-pack QA issues (title blocks / wrong notes / numbering)
- anything important for UK Building Regulations review
- keep commentary tight and avoid generic filler

Extracted text from full PDF:
{text[:30000]}
"""
    content = [{"type": "input_text", "text": prompt}]
    for image_path in image_paths:
        content.append(
            {
                "type": "input_image",
                "image_url": f"data:image/png;base64,{image_to_base64(image_path)}",
            }
        )
    response = _call_responses_api("gpt-5", [{"role": "user", "content": content}])
    return response.output_text


def build_checks(text: str) -> List[str]:
    lower_text = text.lower()
    checks = []

    if "stair" in lower_text:
        checks.append("Check Part K: stair pitch, rise/going, headroom, landings, guarding")
    if "bedroom" in lower_text or "sleep" in lower_text:
        checks.append("Check Part B: protected route, fire doors, alarms, escape provisions")
    if any(word in lower_text for word in ["wc", "bathroom", "ensuite", "shower room", "kitchen", "utility"]):
        checks.append("Check Part F: ventilation requirements to wet rooms and affected habitable rooms")
    if any(word in lower_text for word in ["extension", "rear extension", "side extension", "wraparound"]):
        checks.append("Check extension-related requirements: structure, thermal performance, ventilation, drainage")
    if any(word in lower_text for word in ["dormer", "rooflight", "roof plan", "loft"]):
        checks.append("Check roof / loft-related requirements where truly relevant")
    if any(word in lower_text for word in ["beam", "steel", "padstone", "lintel", "bearing"]):
        checks.append("Check structural engineer calculations and support details")
    if any(word in lower_text for word in ["u-value", "u value", "insulation", "thermal", "roof build-up"]):
        checks.append("Check Part L: thermal build-ups, U-values, and junction coordination")
    if any(word in lower_text for word in ["drain", "svp", "rwp", "soil stack", "gully", "waste pipe"]):
        checks.append("Check Part H: drainage and rainwater coordination")
    return checks


def estimate_confidence(page_data: List[Dict[str, str]], text: str) -> str:
    sheet_types = {p["sheet_type"] for p in page_data}
    score = 0
    if len(page_data) >= 8:
        score += 1
    if "PROPOSED PLAN" in sheet_types:
        score += 1
    if "PROPOSED SECTION" in sheet_types or "PROPOSED ELEVATION" in sheet_types:
        score += 1
    if "BUILDING REGULATION SPECIFICATION" in sheet_types:
        score += 1
    if "FIRE PLAN" in sheet_types:
        score += 1
    if len(text.strip()) > 2000:
        score += 1

    if score >= 5:
        return "HIGH"
    if score >= 3:
        return "MEDIUM"
    return "LOW"


def analyze_pdf(
    pdf_path: str,
    client_project_type: str = "",
    review_mode: str = "Architect / Professional",
    progress_callback: Optional[Callable[[int, int], None]] = None,
) -> str:
    text = extract_text_from_pdf(pdf_path)
    page_data = extract_text_by_page(pdf_path)
    max_pages_for_analysis = LIVE_ANALYSIS_MAX_PAGES
    if len(page_data) > max_pages_for_analysis:
        page_data = page_data[:max_pages_for_analysis]

    if not text.strip():
        return """PROJECT CLASSIFICATION
- Primary Project Type: Not identified
- Secondary Project Type: None
- Project Scope: Unknown
- Affected Storeys: Unknown
- Relevant Building Regulations Focus: Unknown
- Why this classification was chosen: No readable text found.
- Comparison with client-stated description: Could not confirm.

PROJECT DETAILS
- No readable text found in this PDF.

TOP SUMMARY
- Overall Risk Rating: HIGH
- Submission Status: NOT READY TO SUBMIT
- Review Confidence: LOW
- Main Missing Items: no readable text extracted

DRAWING-PACK INCONSISTENCIES
- No readable content available to compare.

EXECUTIVE SUMMARY
- No readable text found in this PDF.
- A text-based exported PDF is recommended.

DRAWING PACK SUMMARY
- Cover Sheet: Not detected
- General Notes: Not detected
- Existing Plans: Not detected
- Proposed Plans: Not detected
- Sections / Elevations: Not detected
- Fire Plans: Not detected
- Building Regulation Specification: Not detected
- Structural Drawings: Not detected
- Details: Not detected

COMPLIANCE STATUS BY APPROVED DOCUMENT
- Part A – Structure: FAIL
  - Why: No readable structural information found.
- Part B – Fire Safety: FAIL
  - Why: No readable fire information found.
- Part F – Ventilation: FAIL
  - Why: No readable ventilation information found.
- Part K – Stairs: FAIL
  - Why: No readable stair information found.
- Part L – Conservation of Fuel and Power: FAIL
  - Why: No readable thermal information found.
- Part P – Electrical Safety: FAIL
  - Why: No readable electrical information found.

KEY RISKS
- HIGH RISK: Entire drawing text could not be extracted.

MISSING INFORMATION
- Full readable drawing content not available.

RECOMMENDED ACTIONS
- Provide a text-based PDF export.
- Check whether the current file is image-based or corrupted.

BUILDING CONTROL SUBMISSION READINESS
- Status: NOT READY TO SUBMIT
- Reason: The pack could not be read reliably enough for a compliance review.
"""

    checks = build_checks(text)
    batch_summaries = []
    total_pages = len(page_data)
    batch_size = IMAGE_BATCH_SIZE
    total_batches = max(1, (total_pages + batch_size - 1) // batch_size)

    for idx, start in enumerate(range(0, total_pages, batch_size), start=1):
        batch_paths = render_pdf_page_batch_to_images(pdf_path, start, start + batch_size, scale=IMAGE_RENDER_SCALE)
        try:
            if progress_callback:
                progress_callback(idx, total_batches)
            batch_summaries.append(analyze_image_batch(batch_paths, text, checks, review_mode))
        finally:
            for image_path in batch_paths:
                if os.path.exists(image_path):
                    os.remove(image_path)
            gc.collect()

    combined_batch_text = "\n\n".join(batch_summaries)

    sheet_summary_lines = [
        f"Page {page['page_number']}: {page['sheet_type']} | Sheet title: {page['sheet_title']}"
        for page in page_data
    ]
    sheet_summary = "\n".join(sheet_summary_lines)
    confidence = estimate_confidence(page_data, text)
    audience_rules = (
        "Write for architects and Building Control reviewers using professional wording."
        if review_mode == "Architect / Professional"
        else "Write for homeowners in simple plain English and avoid unnecessary jargon."
    )

    final_prompt = f"""
You are an AI Building Regulations drawing reviewer for UK residential projects.

Client stated project description:
{client_project_type}

Report mode:
{review_mode}

Audience rule:
{audience_rules}

You have reviewed all pages of the drawing pack using:
1. Full extracted text from the PDF
2. Image-based page batch summaries

Detected compliance checks:
{chr(10).join(checks) if checks else "No specific checks detected"}

Detected sheet types:
{sheet_summary}

Estimated review confidence from pack completeness:
{confidence}

PROJECT SCOPE DETECTION
First determine the TRUE project scope before performing compliance checks.

Possible scopes:
- Ground floor rear extension
- Side extension
- Rear + side extension
- Loft conversion
- Dormer / roof alteration
- Garage conversion
- Internal alterations
- Mixed scope residential project

Rules:
- Only classify as loft conversion if a new habitable loft storey is clearly proposed.
- Do not classify as loft conversion based only on roof plans or the word "loft".
- If extension and loft works are present, classify as Mixed Scope.
- Rooflights, lanterns, and flat roof construction that belong only to the proposed extension are part of the extension scope, not a separate roof alteration.
- Only assign a separate roof alteration scope if the existing main roof is clearly altered beyond the extension roof itself.
- Determine Affected Storeys and only apply storey-specific compliance concerns where truly relevant.

DRAWING PACK QA
Check for coordination issues including:
- inconsistent drawing numbers
- missing scale bars
- missing north arrows
- duplicated sheets
- sheet titles inconsistent with content
- planning notes inside building regulation packs
- specification sheets referencing unrelated project types
- contradictory references between GA, Fire Plans, details, schedules, and specification sheets

GENERAL RULES
- Write in plain professional English.
- Keep sentences short and easy to understand.
- Use bullet points under every heading.
- Be direct, practical, and easy to read.
- Focus on UK residential Building Regulations context.
- Use all page summaries, not just the first pages.
- Do not invent dimensions, schedules, or details that are not shown or specified.
- If something is shown on the drawings, say "Shown on drawing".
- If something is stated in notes or specification, say "Specified in notes/specification".
- If something cannot be verified, say "Not clearly shown".
- Do not say an item is missing if it is clearly shown on drawings, fire plans, or specification sheets.
- Keep commentary tight and avoid irrelevant AI filler.

Return the report using EXACT headings in this order:
PROJECT CLASSIFICATION
PROJECT DETAILS
TOP SUMMARY
DRAWING-PACK INCONSISTENCIES
EXECUTIVE SUMMARY
DRAWING PACK SUMMARY
COMPLIANCE STATUS BY APPROVED DOCUMENT
KEY RISKS
MISSING INFORMATION
RECOMMENDED ACTIONS
BUILDING CONTROL SUBMISSION READINESS

Full PDF text:
{text[:30000]}

Page batch summaries:
{combined_batch_text}
"""

    response = _call_responses_api("gpt-5", final_prompt)
    output_text = response.output_text

    missing = [h for h in REQUIRED_HEADINGS if h not in output_text.upper()]
    if missing:
        repair_prompt = f"""
Rewrite the following report so it contains ALL of these exact headings in this exact order:
{chr(10).join(REQUIRED_HEADINGS)}

Keep the substance, but repair structure and heading order only.

Report to repair:
{output_text}
"""
        repaired = _call_responses_api("gpt-5", repair_prompt)
        output_text = repaired.output_text

    gc.collect()
    return output_text



def polish_planning_report_text(report_text: str, address_text: str, fire_status: str, authority_value: str) -> str:
    text = report_text
    text = text.replace("Comparison with client-stated description:\nAligned.", "")
    text = text.replace("Comparison with client-stated description:\nCould not confirm.", "")
    text = text.replace("Comparison with client-stated description:", "")
    text = text.replace("PLANNING OFFICER STYLE REASONING", "PLANNING ASSESSMENT")
    text = text.replace("Main Constraints / Uncertainties:", "Key Planning Considerations:")
    text = text.replace("Authority was user-entered as Hounslow.", "This review has been prepared against the policy framework of the London Borough of Hounslow.")
    text = re.sub(r"Authority was (?:user-entered|inferred) as\s+([^\.]+)\.", r"This review has been prepared against the policy framework of \1.", text)
    text = text.replace("• Client did not state an address; drawings appear to reference a specific property. Ensure the application address and red line boundary match the intended site.", "• Ensure the application address and red/blue line plans are shown consistently across the submitted drawing set.")
    text = text.replace("This is an initial feasibility opinion only based on the drawing pack and client inputs; it is not a guarantee of planning approval.", "")
    if address_text and "Project Address: Not provided" in text:
        text = text.replace("Project Address: Not provided", f"Project Address: {address_text}")
    if fire_status == "submitted":
        text = text.replace("Fire Statement provided.", "A Fire Statement is evident within the submitted pack.")
        text = text.replace("A fire statement may be required where applicable.", "A Fire Statement is evident within the submitted pack.")
        text = text.replace("A fire statement may be required where applicable under London Plan Policy D12.", "A Fire Statement is evident within the submitted pack and has been considered against London Plan Policy D12 where relevant.")
    else:
        text = text.replace("Fire Statement provided.", "A fire statement may be required where applicable.")
        text = text.replace("A Fire Statement has been supplied in line with London Plan D12.", "A fire statement may be required where applicable under London Plan Policy D12.")
        text = text.replace("A Fire Statement has been provided.", "A fire statement may be required where applicable.")
        text = text.replace("Fire Statement already noted but ensure it is signed and site-specific.", "Where a fire statement is required, ensure it is signed and site-specific.")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text

def analyze_planning_pdf(
    pdf_path: str,
    client_project_types: Optional[List[str]] = None,
    property_type: str = "",
    proposal_summary: str = "",
    project_address: str = "",
    local_authority: str = "",
    review_mode: str = "Architect / Professional",
) -> str:
    text = extract_text_from_pdf(pdf_path)
    page_data = extract_text_by_page(pdf_path)
    project_types_text = ", ".join(client_project_types or []) or "Not stated"
    property_type_text = property_type.strip() or "Not stated"
    proposal_summary_text = proposal_summary.strip() or "Not stated"
    extracted_address = extract_project_address(text)
    address_text = project_address.strip() or extracted_address or "Not stated"
    inferred_authority = detect_local_authority(project_address, f"{text}\n{proposal_summary}")
    authority_value = inferred_authority
    page_summary = "\n".join(
        f"Page {page['page_number']}: {page['sheet_type']} | {page['sheet_title']}" for page in page_data
    )
    audience_hint = (
        "Write like a concise UK planning consultant and delegated officer note."
        if review_mode == "Architect / Professional"
        else "Write in plain English suitable for a homeowner, keep the officer-style reasoning structure, and frame the output as a preliminary planning feasibility report rather than a formal planning decision."
    )

    proposal_features = detect_proposal_features(project_types_text, proposal_summary_text, text, page_summary)
    detected_proposal_label = build_detected_proposal_label(proposal_features, project_types_text)
    project_summary_value = build_project_summary_from_inputs(detected_proposal_label, proposal_summary_text, property_type_text)
    application_type_value = infer_application_type(project_types_text, proposal_summary_text, property_type_text)
    route_confidence_score = calculate_planning_route_confidence_score(
        application_type_value, project_types_text, proposal_summary_text, property_type_text, text, page_summary
    )
    street_precedent_signal = detect_street_precedent_signal(text, page_summary)
    readiness_status, readiness_reason = infer_submission_readiness_from_context(
        application_type_value, project_types_text, proposal_summary_text, text, page_summary
    )

    prompt = f"""
You are reviewing a UK residential planning drawing pack.

{audience_hint}

Client-stated project types:
{project_types_text}

Client-stated property type:
{property_type_text}

Project address input:
{address_text}

Client proposal summary:
{proposal_summary_text}

Local authority input:
{authority_value}

Planning reasoning requirements:
- First identify the proposal accurately from the drawing pack and text. Recognise whether the scheme includes a side gable, rear dormer, rooflights, single-storey extension, side extension, wraparound form or mixed works.
- If the drawings indicate a roof extension to side to form gable, rear dormer and front rooflights, describe that exact combination rather than only referring to a loft extension.
- Include officer-style reasoning using concise delegated report language.
- Include a short street precedent conclusion where the pack suggests similar roof forms, terraced context, repeated dormer patterns, 3D views or wider roofscape context.
- Use the following detected inputs to stabilise the assessment: Detected proposal label = {detected_proposal_label}; Street precedent signal = {street_precedent_signal}; Planning route confidence score = {route_confidence_score}%.
- If review mode is Homeowner Summary, the report should work as a preliminary planning feasibility review based on a simple sketch, basic PDF, or drawing pack.
- Make clear that the output is an initial feasibility opinion only and does not guarantee planning approval.
- Where the sketch or drawing lacks enough information, state the likely route and the main items that still need confirming.
- Apply bungalow logic where relevant. If the dwelling appears to be a bungalow or chalet bungalow, assess scale, ridge/eaves relationship, roof form, bulk and whether side/rear additions read as subordinate.
- Apply PD vs full planning logic. If the works look capable of falling under permitted development, say so. If PD rules are not met or look doubtful, state that full planning is likely required.
- Apply prior approval larger home extension logic where a larger single-storey rear extension appears relevant.
- Flag when the proposal looks more like householder planning than PD.
- Include rear extension risk logic: projection, height, relationship to neighbours, wraparound effects, depth, outlook and design balance.
- Infer local authority from the drawing pack, project address, or client proposal summary if the user did not enter it.
- Use officer-style reasoning and concise delegated report wording patterns such as: "By reason of", "On balance", "Taken cumulatively", "The proposal is likely to", "Insufficient information is shown to confirm".
- Reduce irrelevant commentary. Stay practical and decision-focused.

Important route logic:
- If PD criteria appear clearly met, say PD may be available subject to full dimensional confirmation.
- If PD criteria are not met, are uncertain, or mixed works go beyond PD, state that full planning is likely required.
- For larger home extensions, distinguish between standard PD and prior approval larger home extension.
- If the scheme appears to include side extension, wraparound form, roof changes beyond PD limits, front-facing changes, flats, or other non-PD triggers, explain why full planning is likely required.

Return the report using EXACT headings in this order:
PROJECT CLASSIFICATION
SITE AND PROPOSAL OVERVIEW
TOP SUMMARY
LOCAL AUTHORITY CONTEXT
PD / PRIOR APPROVAL / PLANNING ROUTE
PLANNING ASSESSMENT
DRAWING-PACK INCONSISTENCIES
KEY RISKS
MISSING INFORMATION
RECOMMENDED ACTIONS
SUBMISSION READINESS

Within PLANNING ASSESSMENT, write in a delegated-report tone and cover:
- Site / proposal
- Design, scale and massing
- Neighbouring amenity
- Character and appearance
- Fire safety
- Overall planning balance
Use concise officer-style wording such as "By reason of", "On balance", "The proposal would", "The proposal is likely to".

Section guidance:
PROJECT CLASSIFICATION
- Keep this concise and professional.
- Include only:
  - Primary Project Type
  - Secondary Works
  - Dwelling Type
- Do not include any system-style commentary such as "Comparison with client-stated description" or "Aligned".

SITE AND PROPOSAL OVERVIEW
- Summarise the apparent proposal and affected parts of the property in 2 to 4 clean bullets.
- If a Fire Statement is not evident in the pack, do not say one has been submitted.
- Only mention a fire statement where it is genuinely relevant.

TOP SUMMARY
- Do not include "Overall Planning Risk Rating" or "Planning Approval Probability".
- Include only:
  - Project Summary: {project_summary_value}
  - Application Type: {application_type_value}
  - Planning Route Confidence Score: {route_confidence_score}%
  - {authority_value}
- Present 3 to 6 concise "Key Planning Considerations" bullets only.
- Do not add informal caveat wording here.

LOCAL AUTHORITY CONTEXT
- State the local authority professionally, e.g. "This review has been prepared against the policy framework of the London Borough of Hounslow."
- Do not say whether the authority was user-entered or inferred.
- Refer to the relevant local plan / SPD / London Plan policies only.
- Where constraints mapping is not available, state that conservation area, Article 4 and other site constraints should still be confirmed.

PD / PRIOR APPROVAL / PLANNING ROUTE
- State the most likely route.
- Give a short route explanation in formal professional wording.
- For homeowner mode, explain the likely route in simple plain English.
- State clearly if PD rules do not appear to be met and full planning is likely required.
- State clearly if a larger home extension prior approval route may be relevant.

PLANNING ASSESSMENT
- Write this as a professional planning assessment, not as an AI or third-party reasoning section.
- Use concise delegated-report style bullets or short paragraphs covering:
  - Site / proposal
  - Design, scale and massing
  - Street precedent / surrounding roofscape
  - Neighbouring amenity
  - Privacy / overlooking
  - Character and appearance
  - Fire safety where relevant
  - Overall planning balance
- Where street precedent appears evident, say so directly in a professional way, for example: "Several similar roof extensions appear to exist within the surrounding terrace and, on balance, the proposal is likely to read as part of the established roofscape pattern."
- Do not say a Fire Statement has been submitted unless it is actually evident in the pack.

DRAWING-PACK INCONSISTENCIES
- Flag only genuine inconsistencies between stated project type, notes and drawings.
- Do not refer to what the client did or did not state.

KEY RISKS
- HIGH / MEDIUM / LOW risk bullets only.

MISSING INFORMATION
- Only list specific items needed to confirm route or planning risk.

RECOMMENDED ACTIONS
- Start each bullet with Provide / Confirm / Revise / Check / Submit.

SUBMISSION READINESS
- Status: use this indicative position unless the drawings strongly justify otherwise: {readiness_status}
- Reason: use this indicative reason unless the drawings strongly justify otherwise: {readiness_reason}
- If a similar rear extension / prior approval scheme shows the typical dimensional and policy information clearly and no major contradictions are evident, "READY TO SUBMIT" can be used.
- In homeowner mode, this should reflect preliminary feasibility readiness rather than formal submission certainty.

Full PDF text:
{text[:26000]}

Detected pages:
{page_summary}
"""

    response = _call_responses_api("gpt-5", prompt)
    output_text = response.output_text
    fire_status = infer_fire_statement_status(text, page_summary)
    output_text = polish_planning_report_text(output_text, address_text, fire_status, authority_value)

    top_summary_pattern = r"TOP SUMMARY\n([\s\S]*?)(?=\n[A-Z][A-Z /\-]+\n)"
    top_summary_replacement = (
        "TOP SUMMARY\n"
        f"Project Summary: {project_summary_value}\n"
        f"Application Type: {application_type_value}\n"
        f"Planning Route Confidence Score: {route_confidence_score}%\n"
        f"{authority_value}\n"
    )
    output_text = re.sub(top_summary_pattern, top_summary_replacement, output_text, count=1)
    output_text = re.sub(r"^.*Overall Planning Risk Rating:.*$\n?", "", output_text, flags=re.MULTILINE)
    output_text = re.sub(r"^.*Planning Approval Probability:.*$\n?", "", output_text, flags=re.MULTILINE)

    missing = [h for h in PLANNING_REQUIRED_HEADINGS if h not in output_text.upper()]
    if missing:
        repaired = _call_responses_api(
            "gpt-5",
            f"Rewrite the report so it uses these exact headings in this exact order only:\n{chr(10).join(PLANNING_REQUIRED_HEADINGS)}\n\nReport:\n{output_text}",
        )
        output_text = repaired.output_text
        output_text = polish_planning_report_text(output_text, address_text, fire_status, authority_value)
        top_summary_pattern = r"TOP SUMMARY\n([\s\S]*?)(?=\n[A-Z][A-Z /\-]+\n)"
        top_summary_replacement = (
            "TOP SUMMARY\n"
            f"Project Summary: {project_summary_value}\n"
            f"Application Type: {application_type_value}\n"
            f"Planning Route Confidence Score: {route_confidence_score}%\n"
            f"{authority_value}\n"
        )
        output_text = re.sub(top_summary_pattern, top_summary_replacement, output_text, count=1)
        output_text = re.sub(r"^.*Overall Planning Risk Rating:.*$\n?", "", output_text, flags=re.MULTILINE)
        output_text = re.sub(r"^.*Planning Approval Probability:.*$\n?", "", output_text, flags=re.MULTILINE)
    recommendation_value = infer_officer_recommendation(
        readiness_status,
        output_text,
        proposal_features,
    )
    if "Recommendation:" not in output_text:
        output_text = output_text.rstrip() + f"\n\nRecommendation: {recommendation_value}\n"

    gc.collect()
    return output_text




def infer_planning_statement_mode(report_text: str, sections: Optional[Dict[str, str]] = None) -> str:
    sections = sections or {}
    combined = f"{report_text}\n{sections.get('PD / PRIOR APPROVAL / PLANNING ROUTE', '')}\n{sections.get('PROJECT CLASSIFICATION', '')}".lower()
    if "prior approval" in combined:
        return "prior_approval"
    if any(term in combined for term in ["householder", "full planning", "planning permission", "gable", "dormer", "rooflights", "roof lights", "side extension", "wraparound", "first-floor"]):
        return "householder"
    if "pd may be available" in combined or "permitted development" in combined or "lawful development" in combined:
        return "pd"
    return "householder"


def build_planning_statement_structure(statement_mode: str) -> str:
    if statement_mode == "prior_approval":
        return """Use this approved-style structure:
1. Introduction
2. Site Context and Surroundings
3. Proposal Details
4. Compliance with Prior Approval Requirements
5. Design Considerations
6. Neighbour Amenity
7. Conclusion"""
    if statement_mode == "pd":
        return """Use this approved-style structure:
1. Introduction
2. Site Context and Surroundings
3. Proposal Details
4. Permitted Development Assessment
5. Design Considerations
6. Conclusion"""
    return """Use this approved-style structure:
1. Introduction
2. Site Context and Surroundings
3. Proposal Details
4. Design and Character
5. Neighbour Amenity
6. Planning Policy Context
7. Conclusion"""


def generate_planning_statement(
    report_text: str,
    sections: Optional[Dict[str, str]] = None,
    project_address: str = "",
    client_name: str = "",
    local_authority: str = "",
    review_mode: str = "Architect / Professional",
) -> str:
    sections = sections or {}
    statement_mode = infer_planning_statement_mode(report_text, sections)
    structure_text = build_planning_statement_structure(statement_mode)
    route_context = sections.get("PD / PRIOR APPROVAL / PLANNING ROUTE", "")
    classification_context = sections.get("PROJECT CLASSIFICATION", "")
    overview_context = sections.get("SITE AND PROPOSAL OVERVIEW", "")
    policy_context = sections.get("LOCAL AUTHORITY CONTEXT", "")
    readiness_context = sections.get("SUBMISSION READINESS", "")

    audience_hint = (
        "Write a concise professional planning statement suitable for a UK submission prepared by an architectural practice."
        if review_mode == "Architect / Professional"
        else "Write a plain-English homeowner-friendly planning statement draft while keeping the structure professional."
    )

    prompt = f"""
You are drafting a UK planning statement using an ArchLens AI planning review.

{audience_hint}

Project Address: {project_address or 'Not provided'}
Client: {client_name or 'Not provided'}
Local Authority: {local_authority or 'Not provided'}
Detected statement mode: {statement_mode}

Use the report findings below to draft a practical planning statement.
Keep it factual, clean, concise and application-ready.
Do not invent measurements or policy references that are not supported by the report.
Write like an approved planning statement prepared by a UK architectural practice.

Critical route rules:
- Always follow the detected statement mode.
- If statement mode is "householder", write this as a Householder Planning Application statement and do not describe the proposal as an LDC, PD-only or Prior Approval scheme unless the report clearly says that forms part of the proposal.
- If statement mode is "prior_approval", write this as a Prior Approval statement and explain the larger home extension route clearly.
- If statement mode is "pd", write this as a Permitted Development / Lawful Development style statement only where the report clearly indicates that route.
- If the report refers to side gable, rear dormer and front rooflights, describe that exact combination.
- Always align the proposal description with the detected report content rather than generic wording.

Design reasoning rules:
- If a rear extension is located to the rear of the property and is not visible from the public highway, explain clearly that it would not be visible from the public highway and would therefore not impact the character or appearance of the street scene.
- Use concise planning officer style reasoning where relevant.
- Do not mention fire statements unless they are genuinely relevant to the scheme.
- Keep the statement clean and submission-ready.

{structure_text}

Use these report sections:
PROJECT CLASSIFICATION:
{classification_context}

SITE AND PROPOSAL OVERVIEW:
{overview_context}

LOCAL AUTHORITY CONTEXT:
{policy_context}

ROUTE CONTEXT:
{route_context}

SUBMISSION READINESS:
{readiness_context}

Full report text:
{report_text[:18000]}
"""

    try:
        response = _call_responses_api("gpt-5", prompt)
        return response.output_text
    except Exception:
        if statement_mode == "prior_approval":
            fallback_parts = [
                "1. Introduction",
                f"This Statement has been prepared in support of the proposed development at {project_address or 'the subject property'}. It should be read alongside the submitted drawings.",
                "",
                "2. Site Context and Surroundings",
                sections.get("SITE AND PROPOSAL OVERVIEW", "The site forms part of an established residential setting."),
                "",
                "3. Proposal Details",
                sections.get("PROJECT CLASSIFICATION", "The proposal should be read alongside the submitted drawings and supporting information."),
                "",
                "4. Compliance with Prior Approval Requirements",
                sections.get("PD / PRIOR APPROVAL / PLANNING ROUTE", "The likely statutory route should be confirmed before submission."),
                "",
                "5. Design Considerations",
                "The proposal has been assessed in the context of the host dwelling and surrounding built form, with regard to scale, visual impact and relationship to neighbouring properties.",
                "",
                "6. Neighbour Amenity",
                "The proposal should be considered with regard to outlook, enclosure, daylight and relationship to adjoining occupiers.",
                "",
                "7. Conclusion",
                sections.get("SUBMISSION READINESS", "Further confirmation of route and supporting information may be required prior to submission."),
            ]
            return "\n".join(fallback_parts)

        if statement_mode == "pd":
            fallback_parts = [
                "1. Introduction",
                f"This Statement has been prepared in support of the proposed development at {project_address or 'the subject property'}. It should be read alongside the submitted drawings.",
                "",
                "2. Site Context and Surroundings",
                sections.get("SITE AND PROPOSAL OVERVIEW", "The site forms part of an established residential setting."),
                "",
                "3. Proposal Details",
                sections.get("PROJECT CLASSIFICATION", "The proposal should be read alongside the submitted drawings and supporting information."),
                "",
                "4. Permitted Development Assessment",
                sections.get("PD / PRIOR APPROVAL / PLANNING ROUTE", "The likely statutory route should be confirmed before submission."),
                "",
                "5. Design Considerations",
                "The proposal should be assessed in the context of the host dwelling and surrounding built form, with regard to scale, materials and visual impact.",
                "",
                "6. Conclusion",
                sections.get("SUBMISSION READINESS", "Further confirmation of route and supporting information may be required prior to submission."),
            ]
            return "\n".join(fallback_parts)

        fallback_parts = [
            "1. Introduction",
            f"This Planning Statement has been prepared in support of the proposed development at {project_address or 'the subject property'}. It should be read alongside the submitted drawings.",
            "",
            "2. Site Context and Surroundings",
            sections.get("SITE AND PROPOSAL OVERVIEW", "The application site forms part of an established residential setting and should be assessed in that context."),
            "",
            "3. Proposal Details",
            sections.get("PROJECT CLASSIFICATION", "The proposal should be read alongside the submitted drawings and supporting information."),
            "",
            "4. Design and Character",
            "The proposal should be assessed in the context of the host dwelling and surrounding built form, with regard to scale, visual impact, roof form and relationship to the established pattern of development.",
            "",
            "5. Neighbour Amenity",
            "The proposal should be considered with regard to outlook, enclosure, daylight, privacy and relationship to adjoining occupiers.",
            "",
            "6. Planning Policy Context",
            sections.get("LOCAL AUTHORITY CONTEXT", "The proposal should be assessed against the relevant local and strategic planning policy framework."),
            "",
            "7. Conclusion",
            sections.get("SUBMISSION READINESS", "Further confirmation of route and supporting information may be required prior to submission."),
        ]
        return "\n".join(fallback_parts)
