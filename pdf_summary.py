import base64
import os
import re
import time
from typing import Callable, Dict, List, Optional

import fitz
import pdfplumber
from dotenv import load_dotenv
from openai import OpenAI

load_dotenv()

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
    pages = []
    with pdfplumber.open(pdf_path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_text = clean_extracted_text(page.extract_text() or "")
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
    return pages


def render_all_pdf_pages_to_images(pdf_path: str, scale: float = 1.5) -> List[str]:
    doc = fitz.open(pdf_path)
    image_paths = []
    try:
        for i in range(len(doc)):
            page = doc.load_page(i)
            pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
            image_path = f"temp_page_{i + 1}.png"
            pix.save(image_path)
            image_paths.append(image_path)
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
    return "submitted" if ("fire statement" in combined or "fire plan" in combined) else "not evident in the drawing pack"

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
    image_paths = render_all_pdf_pages_to_images(pdf_path)
    try:
        batch_summaries = []
        batches = list(chunk_list(image_paths, 2))
        total_batches = len(batches)

        for idx, batch in enumerate(batches, start=1):
            if progress_callback:
                progress_callback(idx, total_batches)
            batch_summaries.append(analyze_image_batch(batch, text, checks, review_mode))

        combined_batch_text = "\n\n".join(batch_summaries)
    finally:
        for image_path in image_paths:
            if os.path.exists(image_path):
                os.remove(image_path)

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
    inferred_authority = detect_local_authority(local_authority or project_address, f"{text}\n{proposal_summary}")
    authority_value = local_authority.strip() or inferred_authority
    page_summary = "\n".join(
        f"Page {page['page_number']}: {page['sheet_type']} | {page['sheet_title']}" for page in page_data
    )
    # Default AI estimation placeholders used in the planning prompt
    readiness_status = "LIKELY READY WITH MINOR AMENDMENTS"
    readiness_reason = "Initial AI assessment based on drawing pack completeness and clarity."
    approval_probability = "Medium"

    audience_hint = (
        "Write like a concise UK planning consultant and delegated officer note."
        if review_mode == "Architect / Professional"
        else "Write in plain English suitable for a homeowner, keep the officer-style reasoning structure, and frame the output as a preliminary planning feasibility report rather than a formal planning decision."
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
- Where fire safety documentation is likely to be needed, say "A fire statement may be required where applicable."

TOP SUMMARY
- Overall Planning Risk Rating: use the indicative risk unless the drawings strongly justify otherwise.
- Likely Route: PD / PRIOR APPROVAL / FULL PLANNING / MIXED OR UNCLEAR
- Local Authority Used: {authority_value}
- Planning Approval Probability: {approval_probability}
- Replace "Main Constraints / Uncertainties" with "Key Planning Considerations"
- Present 3 to 6 concise bullets only.
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
  - Neighbouring amenity
  - Privacy / overlooking
  - Character and appearance
  - Fire safety where relevant
  - Overall planning balance
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
    missing = [h for h in PLANNING_REQUIRED_HEADINGS if h not in output_text.upper()]
    if missing:
        repaired = _call_responses_api(
            "gpt-5",
            f"Rewrite the report so it uses these exact headings in this exact order only:\n{chr(10).join(PLANNING_REQUIRED_HEADINGS)}\n\nReport:\n{output_text}",
        )
        output_text = repaired.output_text
    return output_text


def generate_planning_statement(
    report_text: str,
    sections: Optional[Dict[str, str]] = None,
    project_address: str = "",
    client_name: str = "",
    local_authority: str = "",
    review_mode: str = "Architect / Professional",
) -> str:
    sections = sections or {}
    audience_hint = (
        "Write a concise professional planning statement suitable for a householder planning or prior approval submission."
        if review_mode == "Architect / Professional"
        else "Write a plain-English homeowner-friendly planning statement draft."
    )

    prompt = f"""
You are drafting a UK planning statement using an ArchLens AI planning review.

{audience_hint}

Project Address: {project_address or 'Not provided'}
Client: {client_name or 'Not provided'}
Local Authority: {local_authority or 'Not provided'}

Use the report findings below to draft a practical planning statement.
Keep it factual, clean, and application-ready.
Do not invent measurements or policy references that are not supported by the report.
Where the route is PD / prior approval, explain that clearly.
Where planning permission is likely required, explain that clearly.

Use these headings:
1. Site and Surroundings
2. Proposed Development
3. Relevant Planning Route
4. Design and Appearance
5. Impact on Residential Amenity
6. Conclusion

Report text:
{report_text[:18000]}
"""

    try:
        response = _call_responses_api("gpt-5", prompt)
        return response.output_text
    except Exception:
        fallback_parts = [
            "1. Site and Surroundings",
            f"The application site is {project_address or 'the subject property'}. This draft statement is based on the ArchLens AI review and uploaded drawings.",
            "",
            "2. Proposed Development",
            sections.get("SITE AND PROPOSAL OVERVIEW", "The proposal should be read alongside the submitted drawings."),
            "",
            "3. Relevant Planning Route",
            sections.get("PD / PRIOR APPROVAL / PLANNING ROUTE", "The likely statutory route should be confirmed before submission."),
            "",
            "4. Design and Appearance",
            "The proposal should be assessed in the context of the host dwelling and surrounding built form, with particular regard to scale, subservience, materials, and visual integration.",
            "",
            "5. Impact on Residential Amenity",
            "Residential amenity should be considered with regard to outlook, enclosure, daylight, privacy, and the relationship to adjoining occupiers.",
            "",
            "6. Conclusion",
            sections.get("SUBMISSION READINESS", "Further confirmation of route and supporting information may be required prior to submission."),
        ]
        return "\n".join(fallback_parts)
