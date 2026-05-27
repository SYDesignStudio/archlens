import base64
import gc
import os
import re
import time
from typing import Callable, Dict, List, Optional, Tuple

import fitz
import pdfplumber
from dotenv import load_dotenv
from openai import OpenAI

import planning_rules

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

TARGET_REPORT_STYLE_RULES = """
Target writing style:
- Use the rewritten example report as the reference tone: concise, natural, client-friendly and professional.
- Write like a UK architect or planning consultant preparing a clear appraisal for a homeowner.
- Prefer short paragraphs and focused bullets over dense legal wording.
- Keep officer-style reasoning, but avoid robotic, repetitive or backend-sounding phrases.
- Avoid phrases such as "appears broadly capable", "subject to final dimensional confirmation",
  "planning / permitted development requirements", "rear extension dormer", and repeated "likely compliant" wording.
- Use direct wording such as "appears likely to comply", "suitable for an LDC application",
  "the drawings indicate", "should be confirmed", and "before submission".
- Do not include filler, confidence scores, AI/system language, or raw questionnaire labels.
"""

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


def apply_target_report_language(report_text: str) -> str:
    """Remove recurring stiff phrases while preserving report headings and substance."""
    if not report_text:
        return report_text
    text = report_text
    text = text.replace("\u25a0", "-").replace("\uf0b7", "-").replace("\u00a0", " ")
    text = re.sub(r"\ben\s*[-\u25a0]\s*suite\b", "en-suite", text, flags=re.IGNORECASE)
    replacements = [
        (r"\bappears broadly capable of complying with\b", "appears likely to comply with"),
        (r"\bappears broadly capable of being carried out\b", "appears likely to be carried out"),
        (r"\bsubject to final dimensional confirmation\b", "once the key dimensions are confirmed"),
        (r"\bsubject to final confirmation of dimensions\b", "once the key dimensions are confirmed"),
        (r"\bplanning\s*/\s*permitted development requirements\b", "planning or permitted development requirements"),
        (r"\brear extension dormer\b", "rear dormer loft conversion"),
        (r"\bextension dormer\b", "dormer extension"),
        (r"\blikely compliant subject to minor checks\b", "likely suitable for an LDC application with minor checks outstanding"),
        (r"\bInitial AI assessment based on drawing pack completeness and clarity\.\b", "The drawing pack is generally clear, with a small number of items to confirm before submission."),
    ]
    for pattern, replacement in replacements:
        text = re.sub(pattern, replacement, text, flags=re.IGNORECASE)
    text = re.sub(r"(?im)^(\s*[-•]?\s*[^:\n]+:\s*)Unknown\s*$", r"\1To be confirmed", text)
    text = re.sub(r"(?im)^(\s*[-•]?\s*[^:\n]+:\s*)Not shown\s*$", r"\1Not clearly identified in the submitted information", text)
    text = re.sub(r"(?im)^Unknown\s*$", "To be confirmed", text)
    text = re.sub(r"(?im)^Not shown\s*$", "Not clearly identified in the submitted information", text)
    text = re.sub(r"(?im)^(ROUTE POSITION\s*:?)\s*\n\s*\1\s*$", r"\1", text)
    text = re.sub(r"(?im)^(ROUTE POSITION\s*:?)\s*\n\s*(ROUTE POSITION\s*:?)\s*\n", r"\1\n", text)
    text = re.sub(r",{2,}", ",", text)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",\s*\.", ".", text)
    return text


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
    postcode_match = re.search(r"\b([a-z]{1,2})\d[a-z\d]?\s*\d[a-z]{2}\b", combined, flags=re.IGNORECASE)
    outward = ""
    if postcode_match:
        outward_match = re.match(r"([a-z]{1,2}\d[a-z\d]?)", postcode_match.group(0).replace(" ", ""), flags=re.IGNORECASE)
        outward = outward_match.group(1).lower() if outward_match else ""
    authorities = {
        "Hounslow": ["hounslow", "london borough of hounslow"],
        "Ealing": ["ealing", "london borough of ealing"],
        "Hillingdon": ["hillingdon", "uxbridge", "hayes", "ruislip", "london borough of hillingdon"],
        "Richmond upon Thames": ["richmond upon thames", "twickenham", "isleworth"],
        "Brent": ["brent", "wembley", "harlesden"],
        "Barnet": ["barnet", "edgware"],
        "Enfield": ["enfield", "berkshire gardens"],
        "Slough": ["slough", "slough borough council"],
        "Reading": ["reading", "reading borough council"],
        "Surrey Heath": ["surrey heath", "camberley"],
    }
    postcode_prefixes = {
        "Hounslow": ["tw3", "tw4", "tw5", "tw7", "tw13", "tw14"],
        "Ealing": ["w5", "w7", "ub1", "ub2", "ub5", "ub6"],
        "Hillingdon": ["ub3", "ub4", "ub7", "ub8", "ub9", "ub10", "ha4"],
        "Richmond upon Thames": ["tw1", "tw2", "tw9", "tw10", "tw11", "tw12"],
        "Brent": ["nw10", "ha0", "ha9"],
        "Barnet": ["n2", "n3", "n11", "n12", "n20", "en4", "en5", "nw7"],
        "Enfield": ["n9", "n13", "n14", "n21", "en1", "en2", "en3"],
        "Slough": ["sl1", "sl2", "sl3"],
        "Reading": ["rg1", "rg2", "rg4", "rg6", "rg30", "rg31"],
        "Surrey Heath": ["gu15", "gu16", "gu18", "gu19", "gu20", "gu24"],
    }
    for authority, needles in authorities.items():
        if any(needle in combined for needle in needles):
            return authority
    if outward:
        for authority, prefixes in postcode_prefixes.items():
            if outward in prefixes:
                return authority
    return "Not clearly identified"





# -----------------------------------------------------------------------------
# Local planning policy pack support
# -----------------------------------------------------------------------------
POLICY_FOLDER = os.getenv("ARCHLENS_POLICY_FOLDER", "planning_policies")

def _normalise_policy_text(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", " ", str(value or "").lower()).strip()

def _policy_keywords_for_project(project_types_text: str, proposal_summary_text: str) -> List[str]:
    combined = _normalise_policy_text(project_types_text + " " + proposal_summary_text)
    keywords = ["local plan", "development management", "design", "householder", "residential"]
    if any(x in combined for x in ["loft", "dormer", "roof", "rooflight", "gable"]):
        keywords += ["house extension", "householder", "residential extension", "design", "roof", "dormer"]
    if any(x in combined for x in ["rear extension", "side extension", "infill", "wraparound", "first floor"]):
        keywords += ["house extension", "residential extension", "householder", "design", "daylight", "amenity"]
    if any(x in combined for x in ["flat conversion", "house conversion", "hmo", "residential conversion"]):
        keywords += ["ndss", "space standard", "housing", "conversion", "amenity", "refuse", "parking"]
    if "shop" in combined or "commercial" in combined:
        keywords += ["shopfront", "commercial", "town centre", "parking", "servicing"]
    return list(dict.fromkeys(keywords))

def find_relevant_policy_files(authority: str, project_types_text: str, proposal_summary_text: str, max_files: int = 5) -> List[str]:
    folder = POLICY_FOLDER
    if not os.path.isdir(folder):
        return []
    authority_key = _normalise_policy_text(authority)
    project_keys = _policy_keywords_for_project(project_types_text, proposal_summary_text)
    scored: List[Tuple[int, str]] = []
    for name in os.listdir(folder):
        if not name.lower().endswith(".pdf"):
            continue
        full_path = os.path.join(folder, name)
        name_key = _normalise_policy_text(name)
        score = 0
        if authority_key and authority_key != "not clearly identified":
            for part in authority_key.split():
                if part and part in name_key:
                    score += 4
        for key in project_keys:
            key_norm = _normalise_policy_text(key)
            if key_norm and all(part in name_key for part in key_norm.split()[:3]):
                score += 3
            elif any(part in name_key for part in key_norm.split()):
                score += 1
        if score > 0:
            scored.append((score, full_path))
    scored.sort(key=lambda item: item[0], reverse=True)
    return [path for _, path in scored[:max_files]]

def extract_policy_context(authority: str, project_types_text: str, proposal_summary_text: str, max_chars: int = 4200) -> str:
    files = find_relevant_policy_files(authority, project_types_text, proposal_summary_text)
    if not files:
        return "No matching local planning policy PDFs were found in the planning_policies folder for this authority/project type."
    snippets: List[str] = []
    keywords = _policy_keywords_for_project(project_types_text, proposal_summary_text)
    for file_path in files:
        try:
            with pdfplumber.open(file_path) as pdf:
                file_lines: List[str] = []
                for page in pdf.pages[:6]:
                    page_text = clean_extracted_text(page.extract_text() or "")
                    if not page_text:
                        continue
                    for line in page_text.splitlines():
                        line_l = line.lower()
                        if any(k.lower() in line_l for k in keywords) or len(file_lines) < 8:
                            file_lines.append(line.strip())
                        if len("\n".join(file_lines)) > 900:
                            break
                    if len("\n".join(file_lines)) > 900:
                        break
                if file_lines:
                    snippets.append(f"POLICY FILE: {os.path.basename(file_path)}\n" + "\n".join(file_lines[:18]))
        except Exception as exc:
            snippets.append(f"POLICY FILE: {os.path.basename(file_path)}\nCould not read policy PDF: {exc}")
    context = "\n\n".join(snippets).strip()
    return context[:max_chars] if context else "Policy files were found but no readable relevant text could be extracted."

def summarise_planning_history_for_prompt(pd_context: Optional[Dict[str, str]]) -> str:
    if not pd_context:
        return "No previous planning history was provided by the user."
    known = str(pd_context.get("planning_history_known", "Not sure"))
    ref = str(pd_context.get("previous_application_ref", "")).strip()
    decision = str(pd_context.get("previous_decision_type", "Unknown"))
    pd_removed = str(pd_context.get("pd_rights_removed", "Not sure"))
    implemented = str(pd_context.get("previous_permission_implemented", "Not sure"))
    notes = str(pd_context.get("planning_history_notes", "")).strip()
    parts = [
        f"Known previous applications: {known}",
        f"Reference: {ref or 'not provided'}",
        f"Decision type: {decision}",
        f"PD rights removal condition: {pd_removed}",
        f"Previous permission implemented: {implemented}",
    ]
    if notes:
        parts.append(f"Notes: {notes}")
    if pd_removed.lower() == "yes" and implemented.lower() == "no":
        parts.append("Important reasoning point: if a permission that removed PD rights was never implemented, the PD-removal condition may not have taken effect. Verify against the council planning register and site evidence.")
    return "\n".join(parts)


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



def clean_user_context_text(value: str) -> str:
    """Remove app/UI helper text from user supplied context before it appears in reports."""
    text = (value or "").strip()
    if not text:
        return ""
    text = text.replace("_", " ")
    text = re.sub(r"\|\s*Improve Accuracy\s*:\s*.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Improve Accuracy\s*:\s*.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Optional review focus\s*:\s*Not stated", "", text, flags=re.IGNORECASE)
    # Remove internal app/control text that must never appear in client-facing reports.
    cleanup_patterns = [
        r"PD answers\s*:[^|.]*[|.]?",
        r"Above highest roof\s*:[^|.]*[|.]?",
        r"200mm eaves setback\s*:[^|.]*[|.]?",
        r"Side windows obscure glazed\s*:[^|.]*[|.]?",
        r"Roof volume\s*:[^|.]*[|.]?",
        r"Selected scope items to cross-check\s*:[^|.]*[|.]?",
        r"Scope noted\s*:[^|.]*[|.]?",
        r"Important instruction\s*:[^|]*",
        r"Rule intake answers\s*:[^|]*",
        r"PD answers\s*:[^|]*",
        r"Specific review focus / notes\s*:\s*Not stated",
        r"Drawing dimensions take priority if different\.?",
    ]
    for pattern in cleanup_patterns:
        text = re.sub(pattern, "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",\s*,+", ",", text)
    text = re.sub(r"\s*\|\s*", ", ", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,;|.")
    return text


def sentence_case_project_text(text: str) -> str:
    text = clean_user_context_text(text)
    if not text:
        return text
    replacements = {
        "GROUND FLOOR": "ground floor",
        "REAR EXTENSION": "rear extension",
        "SIDE EXTENSION": "side extension",
        "LOFT EXTENSION": "loft extension",
        "PRIOR APPROVAL": "prior approval",
    }
    for a, b in replacements.items():
        text = text.replace(a, b)
    return text[:1].upper() + text[1:]


def build_project_summary_from_inputs(project_types_text: str, proposal_summary_text: str, property_type_text: str) -> str:
    base = proposal_summary_text if proposal_summary_text and proposal_summary_text.lower() != "not stated" else project_types_text
    text = sentence_case_project_text(base)
    if not text or text.lower() == "not stated":
        return "Residential works to the property as shown on the submitted drawings."
    text = re.sub(r"\|", ", ", text)
    text = re.sub(r"\s{2,}", " ", text).strip(" ,")
    lowered = text.lower()
    if lowered.startswith("proposed "):
        sentence = text
    else:
        sentence = "Proposed " + text[0].lower() + text[1:]
    if not sentence.endswith("."):
        sentence += "."
    return sentence


def infer_application_type(project_types_text: str, proposal_summary_text: str, property_type_text: str) -> str:
    combined = f"{project_types_text} {proposal_summary_text} {property_type_text}".lower()
    if property_type_text.lower() in {"flat", "maisonette"}:
        return "FULL PLANNING"
    if any(term in combined for term in ["loft", "dormer", "rooflight", "roof light", "gable"]):
        return "PD / LDC"
    if "ground floor rear extension" in combined and not any(term in combined for term in ["side extension", "wraparound", "first floor", "loft", "dormer", "gable"]):
        return "PRIOR APPROVAL"
    if any(term in combined for term in ["first floor", "conversion", "side extension", "wraparound"]):
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
            "FURTHER INFORMATION RECOMMENDED",
            "A current drawing / site plan set should be confirmed before submission because superseded information appears within the pack.",
        )
    if application_type == "PRIOR APPROVAL" and "ground floor rear extension" in combined:
        if ("6m" in combined or "6000" in combined) and ("3m" in combined or "3000" in combined or "4m" in combined or "4000" in combined):
            return (
                "READY TO SUBMIT",
                "The pack appears to provide the key prior approval information typically required for a larger home extension submission, with dimensions to be checked on the final drawings.",
            )
        return (
            "LIKELY READY WITH MINOR AMENDMENTS",
            "The proposal appears capable of proceeding by prior approval, but key depth / height information should be confirmed clearly from the original rear wall before submission.",
        )
    return (
        "LIKELY READY WITH MINOR AMENDMENTS",
        "Initial AI assessment based on drawing pack completeness and clarity.",
    )




def detect_proposal_features(project_types_text: str, proposal_summary_text: str, text: str, page_summary: str) -> Dict[str, bool]:
    combined = f"{project_types_text}\n{proposal_summary_text}\n{text}\n{page_summary}".lower()
    has_loft = any(term in combined for term in ["loft", "dormer", "rooflight", "roof light", "roof enlargement"])
    has_extension = any(term in combined for term in ["ground floor rear extension", "single-storey rear extension", "single storey rear extension", "side extension", "wraparound", "wrap around", "first floor extension"])
    is_loft_only = has_loft and not has_extension
    if is_loft_only:
        return {
            "gable": any(term in combined for term in ["gable", "hip to gable", "side gable"]),
            "rear_dormer": any(term in combined for term in ["rear dormer", "dormer", "rear roof enlargement"]),
            "front_rooflights": any(term in combined for term in ["front rooflight", "front rooflights", "rooflight", "rooflights"]),
            "single_storey_rear_extension": False,
            "first_floor_extension": False,
            "side_extension": False,
            "wraparound": False,
            "loft_extension": True,
            "flat_or_maisonette": any(term in combined for term in ["flat", "maisonette"]),
        }
    return {
        "gable": any(term in combined for term in ["gable", "hip to gable", "side gable"]),
        "rear_dormer": any(term in combined for term in ["rear dormer", "dormer", "rear roof enlargement"]),
        "front_rooflights": any(term in combined for term in ["front rooflight", "front rooflights", "rooflight", "rooflights"]),
        "single_storey_rear_extension": "ground floor rear extension" in combined or "single-storey rear extension" in combined or "single storey rear extension" in combined,
        "first_floor_extension": "first floor extension" in combined,
        "side_extension": "side extension" in combined,
        "wraparound": "wraparound" in combined or "wrap around" in combined,
        "loft_extension": has_loft,
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


def format_pd_context_for_prompt(pd_context: Optional[Dict[str, str]]) -> str:
    if not pd_context:
        return "No structured PD questionnaire answers were provided."
    lines = []
    for key, value in pd_context.items():
        pretty_key = key.replace("_", " ").strip().title()
        lines.append(f"- {pretty_key}: {value}")
    return "\n".join(lines)



def _ctx_value(pd_context: Optional[Dict[str, str]], key: str) -> str:
    if not pd_context:
        return ""
    return str(pd_context.get(key, "") or "").strip()


def _ctx_yes(pd_context: Optional[Dict[str, str]], key: str) -> bool:
    return _ctx_value(pd_context, key).lower() == "yes"


def infer_route_from_pd_context(
    pd_context: Optional[Dict[str, str]],
    project_types_text: str,
    property_type_text: str,
) -> Tuple[Optional[str], str, str]:
    if not pd_context:
        return None, "", "MEDIUM"

    property_lower = (property_type_text or "").lower()
    project_lower = (project_types_text or "").lower()
    pd_family = _ctx_value(pd_context, "pd_question_family").lower()

    if _ctx_value(pd_context, "is_single_dwellinghouse").lower() == "no":
        return "FULL PLANNING", "The questionnaire indicates the property is not a single dwellinghouse, so standard householder permitted development rights are unlikely to apply.", "HIGH"

    if property_lower in {"flat", "maisonette"}:
        return "FULL PLANNING", "Flats and maisonettes do not normally benefit from the standard householder permitted development rights used in this review.", "HIGH"

    constraints_value = _ctx_value(pd_context, "site_constraints").lower()
    has_article_23 = any(term in constraints_value for term in ["conservation", "article 2(3)", "national park", "world heritage", "aonb", "site of special scientific interest"])
    has_article4 = "article 4" in constraints_value
    has_listed = "listed" in constraints_value

    if has_article4 or has_listed:
        return "FULL PLANNING", "The questionnaire indicates site constraints that may remove or materially restrict normal householder permitted development rights.", "HIGH"

    if _ctx_yes(pd_context, "forward_of_principal_elevation"):
        return "FULL PLANNING", "Works projecting forward of the principal elevation are unlikely to benefit from normal householder permitted development rights.", "HIGH"

    if pd_family == "class_a" or any(term in project_lower for term in ["rear extension", "side extension", "infill extension", "first floor"]):
        try:
            depth = float(_ctx_value(pd_context, "rear_extension_depth_m") or 0)
        except Exception:
            depth = 0.0
        try:
            overall_h = float(_ctx_value(pd_context, "rear_extension_overall_height_m") or 0)
        except Exception:
            overall_h = 0.0

        if overall_h and overall_h > 4.0:
            return "FULL PLANNING", "The stated overall height exceeds the usual 4.0m limit for a single-storey rear extension under Class A.", "HIGH"

        if _ctx_yes(pd_context, "within_2m_of_boundary") and _ctx_value(pd_context, "eaves_height_within_2m").lower() == "no":
            return "FULL PLANNING", "The stated eaves height within 2m of the boundary exceeds the usual 3.0m limit under Class A.", "HIGH"

        if _ctx_value(pd_context, "side_extension_width").lower() == "yes":
            return "FULL PLANNING", "The questionnaire indicates the side extension is more than half the width of the original house, which falls outside the usual Class A side extension limit.", "HIGH"

        if any(term in project_lower for term in ["first floor rear extension", "first floor side extension", "ground floor side extension", "ground floor infill extension"]):
            return "FULL PLANNING", "The selected project type includes side, infill, or first-floor enlargement works that commonly fall outside the simplest Class A routes and normally need fuller planning assessment.", "HIGH"

        detached = property_lower in {"detached", "detached house"} or property_lower.startswith("detached ")
        terrace_or_other = any(term in property_lower for term in ["terraced", "terrace", "semi-detached", "semi detached", "end of terrace", "semi"]) or not detached

        if "ground floor rear extension" in project_lower and depth > 0:
            if detached:
                if depth <= 4.0:
                    risk = "MEDIUM" if _ctx_value(pd_context, "materials_similar").lower() == "no" else "LOW"
                    return "PD / LDC", "The stated detached house rear extension depth sits within the normal Class A rear extension range, subject to full dimensional confirmation.", risk
                if depth <= 8.0:
                    if has_article_23:
                        return "FULL PLANNING", "The larger home extension prior approval route is restricted on article 2(3) land or similar constrained sites, so full planning is more likely required.", "HIGH"
                    return "PRIOR APPROVAL", "The stated detached house rear extension depth is above the normal Class A threshold but may proceed through the larger home extension prior approval route.", "MEDIUM"
                return "FULL PLANNING", "The stated detached house rear extension depth exceeds the larger home extension threshold.", "HIGH"

            if terrace_or_other:
                if depth <= 3.0:
                    risk = "MEDIUM" if _ctx_value(pd_context, "materials_similar").lower() == "no" else "LOW"
                    return "PD / LDC", "The stated rear extension depth sits within the normal Class A rear extension range for a non-detached house, subject to full dimensional confirmation.", risk
                if depth <= 6.0:
                    if has_article_23:
                        return "FULL PLANNING", "The larger home extension prior approval route is restricted on article 2(3) land or similar constrained sites, so full planning is more likely required.", "HIGH"
                    return "PRIOR APPROVAL", "The stated rear extension depth is above the normal Class A threshold for a non-detached house but may proceed through the larger home extension prior approval route.", "MEDIUM"
                return "FULL PLANNING", "The stated rear extension depth exceeds the larger home extension threshold for a non-detached house.", "HIGH"

        if has_article_23 and any(term in project_lower for term in ["side extension", "first floor rear extension", "first floor side extension"]):
            return "FULL PLANNING", "Article 2(3) land introduces extra Class A restrictions for side extensions and multi-storey rear enlargements, so full planning is more likely required.", "HIGH"

        return None, "", "MEDIUM"

    if pd_family == "class_b" or any(term in project_lower for term in ["loft", "dormer"]):
        roof_volume = _ctx_value(pd_context, "roof_volume_band").lower()

        if has_article_23:
            return (
                "FULL PLANNING",
                "The property appears to be within a constrained designation area where Class B roof enlargements may not apply.",
                "HIGH",
            )

        if _ctx_yes(pd_context, "front_roof_plane_highway"):
            return (
                "FULL PLANNING",
                "Front-facing roof enlargements are unlikely to comply with Class B permitted development rules.",
                "HIGH",
            )

        if _ctx_yes(pd_context, "above_existing_roof_height"):
            return (
                "FULL PLANNING",
                "The proposal appears to extend above the existing roof ridge which falls outside Class B.",
                "HIGH",
            )

        if "over limit" in roof_volume:
            return (
                "FULL PLANNING",
                "The additional roof volume appears to exceed normal Class B allowances.",
                "HIGH",
            )

        if _ctx_value(pd_context, "materials_similar").lower() == "no":
            return (
                "PD / LDC",
                "The proposal appears likely to meet the main Class B tests, but external materials should match the existing dwelling.",
                "MEDIUM",
            )

        if _ctx_value(pd_context, "eaves_setback_0_2m").lower() == "no":
            return (
                "PD / LDC",
                "The proposal appears likely to meet the main Class B tests, but the 200mm eaves setback should be confirmed.",
                "MEDIUM",
            )

        return (
            "PD / LDC",
            "The roof enlargement appears likely to comply with Class B permitted development requirements subject to standard dimensional confirmation.",
            "LOW",
        )

    if pd_family == "class_d" or "porch" in project_lower:
        if _ctx_value(pd_context, "porch_ground_area_band").lower() == "no":
            return "FULL PLANNING", "The questionnaire indicates the porch exceeds the usual 3m² Class D ground area limit.", "HIGH"
        if _ctx_value(pd_context, "porch_height_band").lower() == "no":
            return "FULL PLANNING", "The questionnaire indicates the porch exceeds the usual 3m Class D height limit.", "HIGH"
        if _ctx_yes(pd_context, "porch_within_2m_highway"):
            return "FULL PLANNING", "The questionnaire indicates part of the porch would be within 2m of a boundary with a highway, which would fall outside Class D.", "HIGH"
        return "PD / LDC", "The porch appears likely to fall within Class D permitted development once the key dimensions are confirmed.", "LOW"

    return None, "", "MEDIUM"


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
    has_loft = any(word in lower_text for word in ["dormer", "rooflight", "roof plan", "loft", "roof enlargement"])
    has_extension = any(word in lower_text for word in ["rear extension", "side extension", "wraparound", "wrap around", "single storey extension", "single-storey extension"])
    is_loft_only = has_loft and not has_extension

    if "stair" in lower_text:
        checks.append("Check Part K: stair pitch, rise/going, headroom, landings, guarding")
    if "bedroom" in lower_text or "sleep" in lower_text:
        checks.append("Check Part B: protected route, fire doors, alarms, escape provisions")
    if any(word in lower_text for word in ["wc", "bathroom", "ensuite", "shower room", "kitchen", "utility"]):
        checks.append("Check Part F: ventilation requirements to wet rooms and affected habitable rooms")
    if has_extension and not is_loft_only:
        checks.append("Check extension-related requirements: structure, thermal performance, ventilation, drainage")
    if has_loft:
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

{TARGET_REPORT_STYLE_RULES}

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
- Keep the report SIMPLE, SHORT and decision-focused.
- Remove unnecessary background commentary.
- Use short bullets and only mention items relevant to the uploaded project.
- User-entered measurements are supporting context only. If drawings show different dimensions, drawing dimensions take priority.
- If the user asks for a specific review focus, focus the report on that issue and keep unrelated commentary minimal.
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

    output_text = apply_target_report_language(output_text)
    output_text = simplify_report_text(output_text, max_bullets_per_section=6)
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



def is_minor_class_b_condition_issue(rule_engine_summary: str, pd_context: Optional[Dict[str, str]] = None, project_types_text: str = "", proposal_summary_text: str = "") -> bool:
    """Return True where a loft/dormer scheme is broadly Class B/C capable but the
    rule engine has escalated a minor condition/annotation issue too harshly.

    Example: side-facing roof window not yet annotated as obscure glazed and
    non-opening below 1.7m should be treated as a minor PD condition/check, not
    as an automatic full-planning route where all core Class B criteria pass.
    """
    combined_project = f"{project_types_text} {proposal_summary_text}".lower()
    summary_u = (rule_engine_summary or "").upper()
    ctx = pd_context or {}
    family = str(ctx.get("pd_question_family", "")).lower()
    is_roof_project = family == "class_b" or any(t in combined_project for t in ["loft", "dormer", "rooflight", "roof light", "roof enlargement"])
    if not is_roof_project:
        return False

    side_issue = any(t in summary_u for t in ["SIDE WINDOWS", "SIDE-FACING", "SIDE ROOF WINDOWS", "OBSCURE", "1.7M"])
    if not side_issue:
        return False

    hard_fail_terms = [
        "FRONT-FACING ROOF ENLARGEMENT",
        "PRINCIPAL ELEVATION AND FRONTS A HIGHWAY",
        "ABOVE THE HIGHEST PART",
        "EXCEEDS HIGHEST ROOF",
        "OVER LIMIT",
        "EXCEEDS NORMAL CLASS B ALLOWANCE",
        "ROOF VOLUME EXCEEDS",
        "BALCONY",
        "VERANDAH",
        "RAISED PLATFORM",
        "ARTICLE 4",
        "LISTED BUILDING",
        "FLAT OR MAISONETTE",
        "NOT A SINGLE DWELLINGHOUSE",
        "CONSERVATION AREA / ARTICLE 2(3)",
    ]
    if any(term in summary_u for term in hard_fail_terms):
        return False

    front_ok = str(ctx.get("front_roof_plane_highway", "")).lower() in {"no", "not applicable", ""}
    highest_ok = str(ctx.get("above_existing_roof_height", "")).lower() in {"no", "not applicable", ""}
    volume_ok = "over limit" not in str(ctx.get("roof_volume_band", "")).lower()
    return front_ok and highest_ok and volume_ok


def normalise_minor_class_b_route_text(text: str) -> str:
    """Clean AI wording where minor Class B condition issues have been wrongly
    phrased as a full planning trigger.
    """
    if not text:
        return text
    replacements = [
        (r"Planning permission likely required as currently shown due to side-facing roof window\(s\)[^\n.]*[.]?", "PD/LDC appears likely subject to adding a clear note that any side-facing roof window is obscure-glazed and non-opening below 1.7m."),
        (r"Likely planning permission required as currently shown due to side-facing roof window\(s\)[^\n.]*[.]?", "Likely compliant subject to minor drawing notes for any side-facing roof window."),
        (r"full planning is likely required due to side-facing roof window\(s\)[^\n.]*[.]?", "a Lawful Development Certificate route appears likely, subject to side-facing roof window notes."),
        (r"FULL PLANNING\s*\n\s*COMPLIANCE POSITION\s*\n\s*Likely planning permission required", "PD / LDC\nCOMPLIANCE POSITION\nLikely compliant subject to minor checks"),
    ]
    for pattern, repl in replacements:
        text = re.sub(pattern, repl, text, flags=re.IGNORECASE)
    return text

def analyze_planning_pdf(
    pdf_path: str,
    client_project_types: Optional[List[str]] = None,
    property_type: str = "",
    proposal_summary: str = "",
    project_address: str = "",
    local_authority: str = "",
    review_mode: str = "Architect / Professional",
    pd_context: Optional[Dict[str, str]] = None,
    scope_items: Optional[List[str]] = None,
    rule_engine_summary: str = "",
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
    policy_context = extract_policy_context(authority_value, project_types_text, proposal_summary_text)
    planning_history_context = summarise_planning_history_for_prompt(pd_context)
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
    pd_route, pd_route_reason, pd_refusal_risk = infer_route_from_pd_context(
        pd_context, project_types_text, property_type_text
    )
    if pd_route:
        application_type_value = pd_route

    # Deterministic PD rule engine. This runs before the AI narrative so PASS / FAIL /
    # NEEDS CONFIRMATION decisions are based on rules, not free-form AI wording.
    try:
        if not rule_engine_summary:
            rule_facts = planning_rules.facts_from_app_context(
                project_types=client_project_types or [],
                property_type=property_type_text,
                proposal_summary=proposal_summary_text,
                pd_context=pd_context or {},
                scope_items=scope_items or [],
            )
            rule_result = planning_rules.run_householder_pd_rules(rule_facts)
            rule_engine_summary = planning_rules.format_rule_result_for_prompt(rule_result)
        else:
            rule_result = None
    except Exception as rule_error:
        rule_result = None
        rule_engine_summary = f"DETERMINISTIC RULE ENGINE RESULT: NEEDS CONFIRMATION\nSUMMARY: Rule engine could not complete: {rule_error}"

    rule_summary_upper = (rule_engine_summary or "").upper()
    minor_class_b_condition_only = is_minor_class_b_condition_issue(
        rule_engine_summary, pd_context, project_types_text, proposal_summary_text
    )
    if minor_class_b_condition_only:
        application_type_value = "PD / LDC"
        pd_refusal_risk = "LOW"
    elif "FULL PLANNING REQUIRED" in rule_summary_upper:
        application_type_value = "FULL PLANNING"
        pd_refusal_risk = "HIGH"
    elif "PRIOR APPROVAL POSSIBLE" in rule_summary_upper or "LARGER HOME EXTENSION PRIOR APPROVAL" in rule_summary_upper:
        application_type_value = "PRIOR APPROVAL"
    elif "PD POSSIBLE" in rule_summary_upper or "PERMITTED DEVELOPMENT / LDC POSSIBLE" in rule_summary_upper:
        application_type_value = "PD / LDC"

    prompt = f"""
You are reviewing a UK residential planning drawing pack.

{audience_hint}

{TARGET_REPORT_STYLE_RULES}

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

Structured PD questionnaire answers:
{format_pd_context_for_prompt(pd_context)}

Previous planning history / PD rights context:
{planning_history_context}

Relevant local planning policy context from planning_policies folder:
{policy_context}

Deterministic rule engine result:
{rule_engine_summary}

Route-check instructions:
- Use the structured rule result to guide the likely planning route, but write the outcome in professional consultant language.
- Do not use harsh PASS / FAIL wording in the report unless there is a clear policy breach.
- For typical dormer/rooflight LDC schemes, say the proposal is likely suitable for an LDC application where only standard confirmations are outstanding.
- Explain the route briefly and avoid backend rule-engine wording.

Planning reasoning requirements:
- Keep the report SIMPLE, SHORT and decision-focused. Do not include background commentary that does not help the reader decide what to do next.
- Use short bullets. Maximum 3 bullets in overview/route sections and maximum 5 bullets in assessment/risk/action sections unless essential.
- Do not repeat the same caveat in multiple sections.
- Do not include "In simple terms" sections.
- Do not include user questionnaire labels such as "Improve Accuracy" or raw user input strings in the final report.
- Translate all user inputs into natural professional English paragraphs.
- Use the planning history context to identify PD-rights/condition risks. If PD rights may have been removed by a condition, check whether the permission was actually implemented before concluding full planning is required.
- Use the local policy context only where relevant to the project type and authority. Do not quote long policy extracts.
- User-entered measurements are supporting context only. If the drawings show different dimensions, drawing dimensions take priority. If dimensions cannot be verified from the drawings, say "Not clearly dimensioned on the drawings".
- If the user asks for a specific review focus, focus the report on that issue and keep unrelated commentary minimal.
- Only give factual conclusions supported by the uploaded drawings, structured user inputs, or relevant planning policy/PD rules.
- First identify the proposal accurately from the drawing pack and text. Recognise whether the scheme includes a side gable, rear dormer, rooflights, single-storey extension, side extension, wraparound form or mixed works.
- If the drawings indicate a roof extension to side to form gable, rear dormer and front rooflights, describe that exact combination rather than only referring to a loft extension.
- Include officer-style reasoning using concise delegated report language.
- Include a short street precedent conclusion where the pack suggests similar roof forms, terraced context, repeated dormer patterns, 3D views or wider roofscape context.
- Use the detected proposal and street precedent only to stabilise the assessment. Do not mention internal detection labels or confidence scores in the final report.
- If review mode is Homeowner Summary, the report should work as a preliminary planning feasibility review based on a simple sketch, basic PDF, or drawing pack.
- Make clear that the output is an initial feasibility opinion only and does not guarantee planning approval.
- Where the sketch or drawing lacks enough information, state the likely route and the main items that still need confirming.
- Apply bungalow logic where relevant. If the dwelling appears to be a bungalow or chalet bungalow, assess scale, ridge/eaves relationship, roof form, bulk and whether side/rear additions read as subordinate.
- Apply PD vs full planning logic. If the works look capable of falling under permitted development, say so. If PD rules are not met or look doubtful, state that full planning is likely required.
- Use the submitted structured PD questionnaire answers and the Permitted Development Rights for Householders Technical Guidance to stabilise route selection and refusal / approval risk.
- Apply prior approval larger home extension logic where a larger single-storey rear extension appears relevant.
- Flag when the proposal looks more like householder planning than PD.
- Include rear extension risk logic: projection, height, relationship to neighbours, wraparound effects, depth, outlook and design balance.
- Infer local authority from the drawing pack, project address, or client proposal summary if the user did not enter it.
- Use officer-style reasoning and concise delegated report wording patterns such as: "By reason of", "On balance", "Taken cumulatively", "The proposal is likely to", "Insufficient information is shown to confirm".
- Reduce irrelevant commentary. Stay practical and decision-focused.

Important route logic:
- If PD criteria appear clearly met, say PD may be available subject to full dimensional confirmation.
- If PD criteria are not met, are uncertain, or mixed works go beyond PD, state that full planning is likely required.
- For Class B loft/dormer schemes, side-facing window obscurity/non-opening wording is normally a minor PD condition/drawing-note issue, not a full planning trigger, where the core roof volume, ridge height, front roof slope and eaves setback tests are otherwise acceptable.
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
- Summarise the apparent proposal and affected parts of the property in 2 to 3 clean bullets.
- If a Fire Statement is not evident in the pack, do not say one has been submitted.
- Only mention a fire statement where it is genuinely relevant.

TOP SUMMARY
- Keep this section extremely clean and professional.
- Do not include:
  - Overall Planning Risk Rating
  - Planning Approval Probability
  - Planning Route Confidence Score
  - AI commentary
  - Confidence percentages
  - Backend logic references
- Include only:
  - Project Summary: {project_summary_value}
  - Likely Planning Route: {application_type_value}
  - Overall Planning Position: The proposal appears likely to comply with the main planning or permitted development requirements once the key dimensions and planning history are confirmed.
  - Local Authority: {authority_value}
- Add a maximum of 3 concise Key Planning Considerations bullets.
- Use professional consultant wording only.
- Do not add informal caveat wording here.

LOCAL AUTHORITY CONTEXT
- State the local authority professionally, e.g. "This review has been prepared against the policy framework of the London Borough of Hounslow."
- Do not say whether the authority was user-entered or inferred.
- Refer to the relevant local plan / SPD / London Plan policies found in the policy context only where they are relevant.
- Where constraints mapping is not available, state that conservation area, Article 4 and other site constraints should still be confirmed.

PD / PRIOR APPROVAL / PLANNING ROUTE
- Start with a short professional route position.
- Use wording such as "Likely compliant subject to minor checks", "Requires further review", or "Planning permission likely required".
- Do not show backend rule codes, deterministic engine wording, or long GPDO checklists.
- If minor confirmation items are missing, do not treat the proposal as failed.
- Give a short route explanation in formal professional wording.
- For homeowner mode, explain the likely route in simple plain English.
- State clearly if PD rules do not appear to be met and full planning is likely required.
- State clearly if a larger home extension prior approval route may be relevant.

PLANNING ASSESSMENT
- Write this section like a real UK planning consultant / delegated officer report.
- Use concise professional wording.
- Maximum 5 bullets only.
- No repetitive wording.
- No AI language.
- No generic filler.
- Focus only on:
  - Design impact
  - Roof form
  - Scale and bulk
  - Streetscene impact
  - Neighbour amenity
  - PD compliance logic where relevant
- Where similar extensions appear nearby, mention this naturally.
- Example wording style: "The proposed roof enlargement appears subordinate to the existing dwelling and would remain visually contained within the established terrace roofscape."
- Do not say a Fire Statement has been submitted unless it is actually evident in the pack.

DRAWING-PACK INCONSISTENCIES
- Flag only genuine inconsistencies between stated project type, notes and drawings.
- Do not refer to what the client did or did not state.

KEY RISKS
- Keep this section short.
- Only include actual planning risks.
- Maximum 4 bullets.
- Categorise every bullet using LOW / MEDIUM / HIGH.
- Do not include unnecessary warnings.

MISSING INFORMATION
- Only list specific items needed to confirm route or planning risk.

RECOMMENDED ACTIONS
- Maximum 5 actions.
- Start every line with Provide / Confirm / Revise / Check / Submit.
- Keep actions practical and submission-focused.

SUBMISSION READINESS
- Replace robotic wording with professional consultant wording.
- Never use:
  - NOT READY
  - AI confidence
  - system language
- Use only:
  - READY TO SUBMIT
  - LIKELY READY WITH MINOR AMENDMENTS
  - FURTHER INFORMATION RECOMMENDED
- Status: use this indicative position unless the drawings strongly justify otherwise: {readiness_status}
- Reason: use this indicative reason unless the drawings strongly justify otherwise: {readiness_reason}
- Add one concise professional explanation only.
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
    output_text = apply_target_report_language(output_text)
    if minor_class_b_condition_only:
        output_text = normalise_minor_class_b_route_text(output_text)

    if (
        "PD / PRIOR APPROVAL / PLANNING ROUTE" in output_text
        and pd_route_reason
        and not re.search(r"(?im)^route position\s*:?", output_text)
    ):
        route_insert = "Route position:\n" + pd_route_reason + "\n\n"
        output_text = output_text.replace(
            "PD / PRIOR APPROVAL / PLANNING ROUTE\n",
            "PD / PRIOR APPROVAL / PLANNING ROUTE\n" + route_insert,
            1,
        )
    top_summary_pattern = r"TOP SUMMARY\n([\s\S]*?)(?=\n[A-Z][A-Z /\-]+\n)"
    compliance_position = "The proposal appears likely to comply with the main permitted development requirements once the key dimensions and planning history are confirmed." if application_type_value == "PD / LDC" else ("The proposal appears suitable for the prior approval route, provided the neighbour consultation and dimensional checks are satisfied." if application_type_value == "PRIOR APPROVAL" else "The proposal is likely to require a formal planning application and should be assessed against the relevant local planning policies.")
    if minor_class_b_condition_only:
        application_type_value = "PD / LDC"
        compliance_position = "Likely suitable for an LDC application with minor checks outstanding"
    top_summary_replacement = (
        "TOP SUMMARY\n"
        f"Project Summary: {project_summary_value}\n"
        f"Likely Planning Route: {application_type_value}\n"
        f"Overall Planning Position: {compliance_position}\n"
        f"Local Authority: {authority_value}\n"
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
        output_text = apply_target_report_language(output_text)
        top_summary_pattern = r"TOP SUMMARY\n([\s\S]*?)(?=\n[A-Z][A-Z /\-]+\n)"
        compliance_position = "The proposal appears likely to comply with the main permitted development requirements once the key dimensions and planning history are confirmed." if application_type_value == "PD / LDC" else ("The proposal appears suitable for the prior approval route, provided the neighbour consultation and dimensional checks are satisfied." if application_type_value == "PRIOR APPROVAL" else "The proposal is likely to require a formal planning application and should be assessed against the relevant local planning policies.")
    if minor_class_b_condition_only:
        application_type_value = "PD / LDC"
        compliance_position = "Likely suitable for an LDC application with minor checks outstanding"
        top_summary_replacement = (
            "TOP SUMMARY\n"
            f"Project Summary: {project_summary_value}\n"
            f"Likely Planning Route: {application_type_value}\n"
            f"Overall Planning Position: {compliance_position}\n"
            f"Local Authority: {authority_value}\n"
        )
        output_text = re.sub(top_summary_pattern, top_summary_replacement, output_text, count=1)
        output_text = re.sub(r"^.*Overall Planning Risk Rating:.*$\n?", "", output_text, flags=re.MULTILINE)
        output_text = re.sub(r"^.*Planning Approval Probability:.*$\n?", "", output_text, flags=re.MULTILINE)
    output_text = apply_target_report_language(output_text)
    output_text = simplify_report_text(output_text, max_bullets_per_section=5)
    gc.collect()
    return output_text




def simplify_report_text(report_text: str, max_bullets_per_section: int = 6) -> str:
    """Keep generated reports concise, remove filler, and preserve required headings."""
    if not report_text:
        return report_text
    text = report_text
    remove_phrases = [
        "This note is an initial feasibility opinion only and does not guarantee planning approval.",
        "This is an initial feasibility opinion only based on the drawing pack and client inputs; it is not a guarantee of planning approval.",
        "IN SIMPLE TERMS",
        "In simple terms",
        "Selected scope items to cross-check:",
        "Deterministic rule engine result:",
        "Rule-based PD check:",
        "Structured PD route logic:",
        "Important route logic:",
        "Rule engine instructions:",
        "Structured PD questionnaire answers:",
        "PASS / FAIL / NEEDS CONFIRMATION",
    ]
    for phrase in remove_phrases:
        text = text.replace(phrase, "")
    text = re.sub(r"Proposed gROUND", "Proposed ground", text)
    text = re.sub(r"gROUND", "ground", text)
    text = re.sub(r"^Not provided$\n?", "", text, flags=re.MULTILINE)
    text = re.sub(r"ROUTE POSITION\s*\n\s*Not provided\s*\n", "ROUTE POSITION\n", text, flags=re.IGNORECASE)
    text = re.sub(r"(?im)^(route position\s*:?)\s*\n\s*(route position\s*:?)\s*\n", r"\1\n", text)
    text = re.sub(r"Project Summary:\s*PROPOSED\s+", "Project Summary: Proposed ", text)
    text = re.sub(r"Improve Accuracy\s*:\s*[^\n.]+[.]?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Selected scope items to cross-check\s*:[^\n.]*[.]?", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Important instruction\s*:[^\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Rule intake answers\s*:[^\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"PD answers\s*:[^\n]*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Planning Route Confidence Score\s*:[^\n]*\n?", "", text, flags=re.IGNORECASE)
    text = normalise_minor_class_b_route_text(text)

    text = re.sub(r"^.*Overall Planning Risk Rating:.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^.*Planning Approval Probability:.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^.*Planning Route Confidence Score:.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^.*Deterministic engine.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^.*Backend logic.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"Detected proposal label\s*=.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"Street precedent signal\s*=.*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"^(Actual|Required):.*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"^.*Class [A-H]\.\d[^\n]*$\n?", "", text, flags=re.MULTILINE | re.IGNORECASE)
    text = re.sub(r"\s+,", ",", text)
    text = re.sub(r",\s*\.", ".", text)
    text = re.sub(r"\n{3,}", "\n\n", text)

    headings = PLANNING_REQUIRED_HEADINGS + REQUIRED_HEADINGS
    heading_pattern = r"^(" + "|".join(re.escape(h) for h in sorted(set(headings), key=len, reverse=True)) + r")$"
    lines = text.splitlines()
    out = []
    current_heading = None
    item_count = 0
    trim_sections = {"PLANNING ASSESSMENT", "RECOMMENDED ACTIONS", "MISSING INFORMATION", "KEY RISKS", "DRAWING-PACK INCONSISTENCIES"}
    for line in lines:
        stripped = line.strip()
        if re.match(heading_pattern, stripped, flags=re.IGNORECASE):
            current_heading = stripped.upper()
            item_count = 0
            out.append(stripped)
            continue
        if current_heading in trim_sections:
            is_item = bool(stripped) and (stripped.startswith("-") or stripped.startswith("•") or ":" in stripped or len(stripped) > 18)
            if is_item:
                item_count += 1
            if item_count > max_bullets_per_section:
                continue
        out.append(line)
    text = "\n".join(out)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    text = apply_target_report_language(text)
    return text


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
    return """Use this exact planning statement structure and order:
PLANNING STATEMENT
[Proposal Title]
[Site Address]

1. Introduction
2. Site and Surroundings
3. Proposed Development
4. Design and Appearance
5. Impact on Neighbouring Amenity
6. Planning Considerations
7. Conclusion"""


PLANNING_STATEMENT_HEADINGS = [
    "1. Introduction",
    "2. Site and Surroundings",
    "3. Proposed Development",
    "4. Design and Appearance",
    "5. Impact on Neighbouring Amenity",
    "6. Planning Considerations",
    "7. Conclusion",
]


def normalise_planning_statement_text(statement_text: str) -> str:
    text = apply_target_report_language(statement_text or "")
    text = (
        text.replace("\u25a0", "-")
        .replace("\uf0b7", "-")
        .replace("\u2022", "-")
        .replace("\u00a0", " ")
    )
    text = re.sub(r"\ben\s*[-\s]*suite\b", "en-suite", text, flags=re.IGNORECASE)
    text = re.sub(r"[ \t]+", " ", text)

    heading_map = {
        "introduction": "1. Introduction",
        "site and surroundings": "2. Site and Surroundings",
        "site context": "2. Site and Surroundings",
        "proposal description": "3. Proposed Development",
        "proposed development": "3. Proposed Development",
        "design and character": "4. Design and Appearance",
        "design and appearance": "4. Design and Appearance",
        "residential amenity": "5. Impact on Neighbouring Amenity",
        "impact on neighbouring amenity": "5. Impact on Neighbouring Amenity",
        "neighbouring amenity": "5. Impact on Neighbouring Amenity",
        "planning considerations": "6. Planning Considerations",
        "relevant planning policy": "6. Planning Considerations",
        "planning history": "6. Planning Considerations",
        "highways and parking": "6. Planning Considerations",
        "highways / parking": "6. Planning Considerations",
        "planning assessment": "6. Planning Considerations",
        "permitted development / lawful development assessment": "6. Planning Considerations",
        "compliance with prior approval requirements": "6. Planning Considerations",
        "conclusion": "7. Conclusion",
    }
    ordered_headings = PLANNING_STATEMENT_HEADINGS

    title_lines: List[str] = []
    section_lines: Dict[str, List[str]] = {heading: [] for heading in ordered_headings}
    seen_headings = set()
    current_heading: Optional[str] = None
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line:
            if current_heading and section_lines[current_heading] and section_lines[current_heading][-1] != "":
                section_lines[current_heading].append("")
            continue
        if re.match(r"(?i)^planning statement$", line):
            continue
        if re.match(r"(?i)^\[(proposal title|site address)\]$", line):
            continue
        candidate = re.sub(r"^\s*\d+[\.)]\s*", "", line).strip()
        candidate_key = re.sub(r"\s+", " ", candidate).lower()
        replacement = heading_map.get(candidate_key)
        if replacement:
            current_heading = replacement
            seen_headings.add(replacement)
            continue
        if current_heading:
            section_lines[current_heading].append(line)
        else:
            title_lines.append(line)

    output_lines: List[str] = ["PLANNING STATEMENT"]
    output_lines.extend(title_lines[:2])
    for heading in ordered_headings:
        content = section_lines.get(heading, [])
        if output_lines and output_lines[-1] != "":
            output_lines.append("")
        output_lines.append(heading)
        output_lines.extend(content)

    text = "\n".join(output_lines)
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    return text


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
        "Write a professional planning statement suitable for a UK householder planning submission prepared by an architectural practice."
        if review_mode == "Architect / Professional"
        else "Write a plain-English planning statement suitable for a homeowner while keeping the document submission-ready."
    )

    prompt = f"""
You are drafting a UK planning statement using an ArchLens AI planning review.

{audience_hint}

{TARGET_REPORT_STYLE_RULES}

Project Address: {project_address or 'Not provided'}
Client: {client_name or 'Not provided'}
Local Authority: {local_authority or 'Not provided'}
Detected statement mode: {statement_mode}

Use the report findings below to draft a practical planning statement.
The master style is a short UK architectural practice planning statement: clear title block, numbered headings, simple professional paragraphs and no checklist tone.
Keep it factual, clean, concise and application-ready.
Do not invent measurements or policy references that are not supported by the report.
Write like a real planning statement prepared by a UK architectural practice for submission with a planning application.

Critical route rules:
- Always follow the detected statement mode.
- If statement mode is "householder", write this as a Householder Planning Application statement and do not describe the proposal as an LDC, PD-only or Prior Approval scheme unless the report clearly says that forms part of the proposal.
- If statement mode is "prior_approval", write this as a Prior Approval statement and explain the larger home extension route clearly.
- If statement mode is "pd", write this as a Permitted Development / Lawful Development style statement only where the report clearly indicates that route.
- If the report refers to side gable, rear dormer and front rooflights, describe that exact combination.
- Always align the proposal description with the detected report content rather than generic wording.
- Do not add extra section headings outside the required Planning Statement structure.
- Where route-specific reasoning is needed, include it naturally under Planning Considerations, Design and Appearance, Impact on Neighbouring Amenity or Conclusion.
- Do not use report-card wording, risk-rating wording, checklist labels, confidence scores or backend/system language.

Design reasoning rules:
- If a rear extension is located to the rear of the property and is not visible from the public highway, explain clearly that it would not be visible from the public highway and would therefore not impact the character or appearance of the street scene.
- Use concise planning officer style reasoning where relevant.
- Do not mention fire statements unless they are genuinely relevant to the scheme.
- Keep the statement clean and submission-ready.
- Use short paragraphs as the main format. Use bullets only where they are genuinely needed.
- The first section should begin naturally, for example: "This Planning Statement has been prepared in support of..."
- Adapt the proposal title and wording to the actual project type and report content. Do not hard-code a rear extension or any example wording.

{structure_text}

Section guidance:
1. Introduction - state the application type, proposal, site address and purpose of the statement.
2. Site and Surroundings - describe the dwelling, residential context, street scene and relevant surrounding pattern only where supported.
3. Proposed Development - describe the actual proposed works, internal layout improvements and drawing information where available.
4. Design and Appearance - explain scale, form, materials, subordination and relationship to the host dwelling.
5. Impact on Neighbouring Amenity - cover privacy, outlook, daylight, overshadowing and overbearing impact in plain language.
6. Planning Considerations - cover the likely planning route, local authority context, policy considerations and planning balance.
7. Conclusion - give a concise professional conclusion on acceptability and submission readiness.

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
        return normalise_planning_statement_text(response.output_text)
    except Exception:
        proposal_title = sections.get("PROJECT CLASSIFICATION", "").splitlines()[0].strip() or "Proposed Residential Development"
        site_address = project_address or "Site address to be confirmed"
        intro_application = "householder planning application"
        if statement_mode == "prior_approval":
            intro_application = "prior approval application"
        elif statement_mode == "pd":
            intro_application = "lawful development certificate application"
        fallback_parts = [
            "PLANNING STATEMENT",
            proposal_title,
            site_address,
            "",
            "1. Introduction",
            f"This Planning Statement has been prepared in support of a {intro_application} for {site_address}.",
            "The proposal seeks to improve the use and functionality of the existing property and should be read alongside the submitted drawings and supporting information.",
            "",
            "2. Site and Surroundings",
            sections.get("SITE AND PROPOSAL OVERVIEW", "The application site forms part of an established residential setting and should be assessed in that context."),
            "",
            "3. Proposed Development",
            sections.get("PROJECT CLASSIFICATION", "The proposal should be read alongside the submitted drawings and supporting information."),
            "",
            "4. Design and Appearance",
            "The proposed works should be assessed in the context of the host dwelling and surrounding built form, with regard to scale, appearance, materials and relationship to the established pattern of development.",
            "",
            "5. Impact on Neighbouring Amenity",
            "The proposal should be considered with regard to outlook, enclosure, daylight, privacy and the relationship with adjoining occupiers.",
            "",
            "6. Planning Considerations",
            "\n\n".join([
                sections.get("LOCAL AUTHORITY CONTEXT", "The proposal should be assessed against the relevant local and strategic planning policy framework."),
                sections.get("PD / PRIOR APPROVAL / PLANNING ROUTE", "The likely planning route should be confirmed before submission."),
                "Planning history, permitted development rights, Article 4 status and any relevant conditions should be checked before submission where they are not already confirmed.",
            ]).strip(),
            "",
            "7. Conclusion",
            sections.get("SUBMISSION READINESS", "Further confirmation of route and supporting information may be required prior to submission."),
        ]
        return normalise_planning_statement_text("\n".join(fallback_parts))
