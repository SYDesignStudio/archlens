
import os
import base64
import gc
import re
import time
import uuid
import tempfile
import jwt
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
    "saved_projects": [],
}
for key, value in DEFAULT_STATE.items():
    if key not in st.session_state:
        st.session_state[key] = value

MAX_FILE_SIZE_MB = 20
MAX_PAGE_COUNT = 30
STARTER_MONTHLY_REVIEW_LIMIT = 10

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
    "COMPLIANCE SNAPSHOT",
    "PLANNING ASSESSMENT",
    "DRAWING-PACK INCONSISTENCIES",
    "KEY RISKS",
    "MISSING INFORMATION",
    "RECOMMENDED ACTIONS",
    "PROFESSIONAL CONCLUSION",
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
    ("COMPLIANCE SNAPSHOT", "Compliance Snapshot"),
    ("PLANNING ASSESSMENT", "Planning Assessment"),
    ("DRAWING-PACK INCONSISTENCIES", "Drawing-Pack Inconsistencies"),
    ("KEY RISKS", "Key Risks"),
    ("MISSING INFORMATION", "Missing Information"),
    ("RECOMMENDED ACTIONS", "Recommended Actions"),
    ("PROFESSIONAL CONCLUSION", "Professional Conclusion"),
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
    "This report provides an initial planning review based on the submitted drawings and information. "
    "It does not replace a formal planning appraisal or the decision of the Local Planning Authority."
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


CLASS_A_PROJECT_TYPES = {
    "Ground Floor Rear Extension",
    "Ground Floor Side Extension",
    "Ground Floor Infill Extension",
    "First Floor Rear Extension",
    "First Floor Side Extension",
}


def get_accuracy_question_group(project_types: List[str]) -> str:
    selected = set(project_types or [])
    if "Loft Extension" in selected:
        return "class_b"
    if selected & CLASS_A_PROJECT_TYPES:
        return "class_a"
    if "Porch" in selected:
        return "class_d"
    return "generic"


def render_improve_accuracy_section(project_types: List[str]):
    answers: Dict[str, str] = {}
    with st.expander("Improve Accuracy (Optional)", expanded=False):
        st.caption("Simple extra questions based on the householder technical guidance. Only the relevant questions for the selected project type are shown.")
        client_name = st.text_input("Client")
        review_date = st.date_input("Report Date")
        group = get_accuracy_question_group(project_types)

        if group == "class_b":
            answers["pd_question_family"] = "class_b"
            st.caption("Class B – additions etc. to the roof")
            answers["site_constraints"] = st.multiselect(
                "Any site constraints that may affect PD rights?",
                ["None", "Conservation area / Article 2(3) land", "Article 4 direction", "Listed building"],
                default=["None"],
                key="acc_b_constraints",
            )
            answers["front_roof_plane_highway"] = st.selectbox(
                "Does any part project from the front roof slope facing the highway?",
                ["Not sure", "No", "Yes"],
                key="acc_b_front",
            )
            answers["roof_volume_band"] = st.selectbox(
                "Added roof volume",
                ["Not sure", "Within normal limit", "Over limit"],
                key="acc_b_volume",
                help="Use 'Within normal limit' where the total added roof volume appears to stay within the normal Class B allowance for the house type.",
            )
            answers["above_existing_roof_height"] = st.selectbox(
                "Does it rise above the highest part of the existing roof?",
                ["Not sure", "No", "Yes"],
                key="acc_b_highest",
            )
            answers["eaves_setback_0_2m"] = st.selectbox(
                "Is the enlargement set back at least 200mm from the original eaves?",
                ["Not sure", "Yes", "No"],
                key="acc_b_eaves",
            )
            answers["side_windows_obscure_glazed"] = st.selectbox(
                "Any side windows? If yes, will they be obscure glazed and non-opening below 1.7m?",
                ["Not applicable", "Yes", "No", "Not sure"],
                key="acc_b_windows",
            )
            answers["materials_similar"] = st.selectbox(
                "Will external materials be similar in appearance to the existing house?",
                ["Not sure", "Yes", "No"],
                key="acc_b_materials",
            )

        elif group == "class_a":
            answers["pd_question_family"] = "class_a"
            st.caption("Class A – enlargement, improvement or alteration")
            answers["site_constraints"] = st.multiselect(
                "Any site constraints that may affect PD rights?",
                ["None", "Conservation area / Article 2(3) land", "Article 4 direction", "Listed building"],
                default=["None"],
                key="acc_a_constraints",
            )
            answers["forward_of_principal_elevation"] = st.selectbox(
                "Does any part project beyond the principal elevation or a side elevation facing a highway?",
                ["Not sure", "No", "Yes"],
                key="acc_a_principal",
            )
            answers["existing_rear_extensions"] = st.selectbox(
                "Have there been previous rear extensions added to the original house?",
                ["Not sure", "No", "Yes"],
                key="acc_a_existing",
            )
            answers["within_2m_of_boundary"] = st.selectbox(
                "Is any part within 2m of a boundary?",
                ["Not sure", "No", "Yes"],
                key="acc_a_boundary",
            )
            answers["eaves_height_within_2m"] = st.selectbox(
                "If within 2m of a boundary, are the eaves 3.0m or lower?",
                ["Not applicable", "Yes", "No", "Not sure"],
                key="acc_a_eaves",
            )
            answers["side_extension_width"] = st.selectbox(
                "For a side extension, is it more than half the width of the original house?",
                ["Not applicable", "No", "Yes", "Not sure"],
                key="acc_a_width",
            )
            answers["materials_similar"] = st.selectbox(
                "Will external materials be similar in appearance to the existing house?",
                ["Not sure", "Yes", "No"],
                key="acc_a_materials",
            )

        elif group == "class_d":
            answers["pd_question_family"] = "class_d"
            st.caption("Class D – porch")
            answers["site_constraints"] = st.multiselect(
                "Any site constraints that may affect PD rights?",
                ["None", "Conservation area / Article 2(3) land", "Article 4 direction", "Listed building"],
                default=["None"],
                key="acc_d_constraints",
            )
            answers["porch_ground_area_band"] = st.selectbox(
                "Is the porch ground area 3m² or less?",
                ["Not sure", "Yes", "No"],
                key="acc_d_area",
            )
            answers["porch_height_band"] = st.selectbox(
                "Is the porch 3.0m high or lower?",
                ["Not sure", "Yes", "No"],
                key="acc_d_height",
            )
            answers["porch_within_2m_highway"] = st.selectbox(
                "Is any part within 2m of a boundary with a highway?",
                ["Not sure", "No", "Yes"],
                key="acc_d_highway",
            )
        else:
            st.caption("Add only anything that clearly changes the planning route.")

        answers["accuracy_notes"] = st.text_area(
            "Anything else that could affect PD, prior approval, or full planning route?",
            height=90,
            placeholder="Example: Article 4 area, listed building, previous extensions already built, or front dormer proposed.",
            key=f"acc_notes_{group}",
        )

    return client_name, review_date, answers


def get_planning_route_snapshot(
    project_types: List[str],
    property_type: str,
    proposal_summary: str,
    rear_extension_depth_m=None,
    rear_extension_height_m=None,
    accuracy_answers: Dict[str, str] | None = None,
):
    selected = set(project_types or [])
    accuracy_answers = accuracy_answers or {}
    summary = (proposal_summary or '').lower()
    property_type_l = (property_type or '').lower()

    route = 'Full Planning likely'
    risk = 'Medium'
    reason = 'Add core project details to improve route accuracy.'

    if 'Loft Extension' in selected:
        front = accuracy_answers.get('front_roof_plane_highway', 'Not sure')
        highest = accuracy_answers.get('above_existing_roof_height', 'Not sure')
        volume = accuracy_answers.get('roof_volume_band', 'Not sure')
        if front == 'No' and highest == 'No' and volume == 'Within normal limit':
            route = 'PD / LDC possible'
            risk = 'Medium'
            reason = 'Class B may apply if the roof enlargement stays off the front roof slope, stays below the highest roof part, and remains within the normal volume allowance.'
        elif front == 'Yes' or highest == 'Yes' or volume == 'Over limit':
            route = 'Full Planning likely'
            risk = 'High'
            reason = 'Class B is less likely where the loft addition affects the front roof slope, exceeds the highest roof part, or appears above the normal volume allowance.'
        else:
            route = 'PD / LDC possible'
            risk = 'Medium'
            reason = 'Loft works may fall under Class B, but the key front roof, highest roof, and volume checks still need confirmation.'

    elif selected & CLASS_A_PROJECT_TYPES:
        depth = rear_extension_depth_m or 0.0
        height = rear_extension_height_m or 0.0
        detached = 'detached' in property_type_l
        attached = any(x in property_type_l for x in ['semi', 'terraced', 'end of terrace'])
        principal = accuracy_answers.get('forward_of_principal_elevation', 'Not sure')
        within_boundary = accuracy_answers.get('within_2m_of_boundary', 'Not sure')
        eaves_test = accuracy_answers.get('eaves_height_within_2m', 'Not sure')

        if principal == 'Yes' or 'first floor' in summary or {'First Floor Rear Extension', 'First Floor Side Extension'} & selected:
            route = 'Full Planning likely'
            risk = 'High'
            reason = 'Upper-floor enlargements and works projecting beyond the principal elevation are usually outside straightforward Class A PD.'
        else:
            pd_limit = 4.0 if detached else 3.0
            pa_limit = 8.0 if detached else 6.0
            if 'Ground Floor Rear Extension' in selected and depth and depth <= pd_limit and height <= 4.0 and not (within_boundary == 'Yes' and eaves_test == 'No'):
                route = 'PD / LDC possible'
                risk = 'Low'
                reason = 'The extension appears within standard Class A depth and height limits, subject to full PD checks.'
            elif 'Ground Floor Rear Extension' in selected and depth and depth <= pa_limit and height <= 4.0:
                route = 'Prior Approval likely'
                risk = 'Medium'
                reason = 'The rear extension appears to fall within the larger home extension range, so the neighbour consultation / prior approval route may apply.'
            else:
                route = 'Full Planning likely'
                risk = 'High'
                reason = 'The extension appears to exceed the usual Class A / larger home extension limits or needs fuller planning assessment.'

    elif 'Porch' in selected:
        area = accuracy_answers.get('porch_ground_area_band', 'Not sure')
        h = accuracy_answers.get('porch_height_band', 'Not sure')
        highway = accuracy_answers.get('porch_within_2m_highway', 'Not sure')
        if area == 'Yes' and h == 'Yes' and highway == 'No':
            route = 'PD possible'
            risk = 'Low'
            reason = 'The porch appears capable of falling within Class D size, height, and highway-distance limits.'
        elif area == 'No' or h == 'No' or highway == 'Yes':
            route = 'Full Planning likely'
            risk = 'High'
            reason = 'The porch appears to fail one or more Class D limits on area, height, or distance to the highway boundary.'
        else:
            route = 'PD possible'
            risk = 'Medium'
            reason = 'A porch may fall under Class D, but the area, height, and highway checks still need confirmation.'

    elif 'Flat Conversion' in selected or 'House Conversion' in selected or 'flat' in property_type_l or 'maisonette' in property_type_l:
        route = 'Full Planning likely'
        risk = 'High'
        reason = 'Conversions and works to flats / maisonettes usually need fuller planning review rather than householder PD.'

    return route, risk, reason


def build_accuracy_context(answers: Dict[str, str]) -> str:
    if not answers:
        return ""
    label_map = {
        'site_constraints': 'Site constraints',
        'front_roof_plane_highway': 'Front roof slope facing highway',
        'roof_volume_band': 'Added roof volume',
        'above_existing_roof_height': 'Above highest roof',
        'eaves_setback_0_2m': '200mm eaves setback',
        'side_windows_obscure_glazed': 'Side windows obscure glazed',
        'materials_similar': 'Materials similar to existing house',
        'forward_of_principal_elevation': 'Projects beyond principal elevation / highway side',
        'existing_rear_extensions': 'Previous rear extensions to original house',
        'within_2m_of_boundary': 'Within 2m of boundary',
        'eaves_height_within_2m': 'Boundary eaves test',
        'side_extension_width': 'Side extension more than half width',
        'porch_ground_area_band': 'Porch ground area 3m² or less',
        'porch_height_band': 'Porch height 3m or lower',
        'porch_within_2m_highway': 'Porch within 2m of highway boundary',
        'accuracy_notes': 'Extra route notes',
    }
    lines = []
    for key, value in answers.items():
        if key == 'pd_question_family':
            continue
        if isinstance(value, list):
            value_str = ", ".join(v for v in value if v and v != 'None').strip()
        else:
            value_str = str(value).strip()
        if not value_str or value_str in {'Not sure', 'Not stated', '', 'None'}:
            continue
        lines.append(f"{label_map.get(key, key)}: {value_str}")
    return " | ".join(lines)


def build_pd_context(project_types: List[str], property_type: str, rear_extension_depth_m, rear_extension_height_m, accuracy_answers: Dict[str, str]) -> Dict[str, str]:
    pd_context: Dict[str, str] = dict(accuracy_answers or {})
    pd_context['is_single_dwellinghouse'] = 'no' if (property_type or '').strip().lower() in {'flat', 'maisonette'} else 'yes'
    constraints = pd_context.get('site_constraints', [])
    if isinstance(constraints, list):
        pd_context['site_constraints'] = ', '.join([c for c in constraints if c and c != 'None']) or 'None'
    if rear_extension_depth_m is not None:
        pd_context['rear_extension_depth_m'] = f"{float(rear_extension_depth_m):.1f}"
    if rear_extension_height_m is not None:
        pd_context['rear_extension_overall_height_m'] = f"{float(rear_extension_height_m):.1f}"
    if not pd_context.get('pd_question_family'):
        pd_context['pd_question_family'] = get_accuracy_question_group(project_types)
    return pd_context



PLAN_LABELS = {
    "starter": "Solo",
    "pro": "Studio",
}
WEBSITE_PRICING_URL = "https://www.sydesignstudio.co.uk/pricing-plans"
WEBSITE_HOME_URL = "https://www.sydesignstudio.co.uk"
WEBSITE_LAUNCH_URL = "https://www.sydesignstudio.co.uk/archlens-launch"
ARCHLENS_SHARED_SECRET = "ArchLens-SYDS-2026-very-long-random-private-secret-839201"

def get_verified_plan_and_user() -> Tuple[str, str, bool]:
    try:
        query_params = st.query_params
        token_value = query_params.get("token", "")
    except Exception:
        token_value = ""

    if isinstance(token_value, list):
        token_value = token_value[0] if token_value else ""

    token_value = str(token_value).strip()
    if not token_value:
        return "starter", "", False

    try:
        payload = jwt.decode(
            token_value,
            ARCHLENS_SHARED_SECRET,
            algorithms=["HS256"],
        )
        plan = str(payload.get("plan", "starter")).strip().lower()
        if plan not in {"starter", "pro"}:
            plan = "starter"

        email = str(payload.get("email", "")).strip()
        display_name = email or str(payload.get("sub", "")).strip()
        return plan, display_name, True
    except Exception:
        return "starter", "", False


def get_allowed_review_modules(plan: str) -> List[str]:
    return ["Planning Review", "Building Regulations Review"] if plan == "pro" else ["Planning Review"]


def get_plan_upgrade_message(feature_name: str) -> str:
    return f"{feature_name} is available on Studio. Upgrade to unlock this feature."

def add_saved_project(project_record: Dict):
    saved = st.session_state.get("saved_projects", [])
    filtered = [item for item in saved if item.get("report_id") != project_record.get("report_id")]
    filtered.insert(0, project_record)
    st.session_state["saved_projects"] = filtered[:25]


def inject_custom_css():
    st.markdown(
        """
        <style>
        :root {
            --sy-bg: var(--background-color);
            --sy-surface: color-mix(in srgb, var(--secondary-background-color) 88%, transparent);
            --sy-surface-2: color-mix(in srgb, var(--secondary-background-color) 76%, transparent);
            --sy-border: color-mix(in srgb, var(--text-color) 14%, transparent);
            --sy-text: var(--text-color);
            --sy-muted: color-mix(in srgb, var(--text-color) 76%, transparent);
            --sy-accent: var(--primary-color);
            --sy-card-shadow: 0 12px 28px rgba(0,0,0,0.16);
        }

        .stApp {
            background: linear-gradient(
                180deg,
                color-mix(in srgb, var(--background-color) 96%, #0A1630 4%) 0%,
                var(--background-color) 100%
            );
            color: var(--sy-text);
        }

        header[data-testid="stHeader"] { display:none !important; }
        [data-testid="stToolbar"] { display:none !important; }
        .stAppDeployButton { display:none !important; }
        #MainMenu { visibility:hidden !important; }
        footer { visibility:hidden !important; }

        .block-container {
            padding-top: 1rem !important;
            padding-bottom: 2rem;
            max-width: 1600px;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(
                180deg,
                color-mix(in srgb, var(--background-color) 92%, #08111F 8%) 0%,
                color-mix(in srgb, var(--background-color) 98%, black 2%) 100%
            );
            border-right: 1px solid var(--sy-border);
        }
        [data-testid="stSidebar"] * { color: var(--sy-text); }

        .sy-topbar {
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:1rem;
            border:1px solid var(--sy-border);
            border-radius:18px;
            background: linear-gradient(
                90deg,
                color-mix(in srgb, var(--secondary-background-color) 92%, transparent),
                color-mix(in srgb, var(--secondary-background-color) 76%, var(--primary-color) 8%)
            );
            margin-bottom:1.1rem;
            box-shadow: var(--sy-card-shadow);
        }
        .sy-topbar-title {
            font-size:0.9rem;
            text-transform:uppercase;
            letter-spacing:0.12em;
            color:var(--sy-muted);
        }
        .sy-topbar-meta { font-size:0.9rem; color:var(--sy-text); }

        .sy-hero {
            padding:1.65rem 1.4rem 1.4rem 1.4rem;
            border:1px solid var(--sy-border);
            border-radius:24px;
            background:
                radial-gradient(circle at 8% 18%, color-mix(in srgb, var(--primary-color) 22%, transparent), transparent 30%),
                radial-gradient(circle at 92% 12%, color-mix(in srgb, var(--primary-color) 14%, #00d4ff 22%), transparent 28%),
                linear-gradient(
                    135deg,
                    color-mix(in srgb, var(--secondary-background-color) 70%, var(--primary-color) 12%) 0%,
                    color-mix(in srgb, var(--background-color) 92%, transparent) 55%,
                    color-mix(in srgb, var(--background-color) 98%, black 2%) 100%
                );
            margin-bottom:1rem;
            box-shadow:0 18px 40px rgba(0,0,0,0.18);
        }
        .sy-hero-grid { display:grid; grid-template-columns:1.4fr 0.9fr; gap:1rem; align-items:start; }
        .sy-hero-copy { padding-top:0.25rem; }
        .sy-hero-copy h1 {
            margin:0 0 0.55rem 0 !important;
            line-height:1.06;
            color:var(--sy-text);
            font-size:3.05rem;
            letter-spacing:-0.03em;
        }
        .sy-hero-copy .sy-muted { margin-top:0.2rem; line-height:1.7; color:var(--sy-muted); }

        .sy-hero-stat,
        .sy-step,
        .sy-card,
        .sy-mini-card,
        .sy-upload-item,
        .sy-sidepanel,
        .sy-workspace {
            border:1px solid var(--sy-border);
            background: linear-gradient(
                180deg,
                color-mix(in srgb, var(--secondary-background-color) 86%, transparent),
                color-mix(in srgb, var(--secondary-background-color) 76%, transparent)
            );
            box-shadow: var(--sy-card-shadow);
            color: var(--sy-text);
        }

        .sy-hero-stat { border-radius:18px; padding:1rem; min-height:82px; }
        .sy-step { border-radius:18px; padding:1rem; min-height:112px; }
        .sy-card { border-radius:20px; padding:1.05rem; margin-bottom:0.95rem; }
        .sy-mini-card { border-radius:18px; padding:1rem; min-height:158px; }
        .sy-sidepanel { border-radius:20px; padding:1rem; position:sticky; top:1rem; }
        .sy-workspace { border-radius:22px; padding:1rem; }

        .sy-panel-title,
        .sy-kpi {
            font-size:0.82rem;
            text-transform:uppercase;
            letter-spacing:0.1em;
            color:var(--sy-muted);
            margin-bottom:0.45rem;
        }

        .sy-workspace-header { display:flex; justify-content:space-between; align-items:center; margin-bottom:0.8rem; }
        .sy-workspace-title { font-size:1.08rem; font-weight:700; color:var(--sy-text); }
        .sy-workspace-meta, .sy-muted { font-size:0.89rem; color:var(--sy-muted); }

        .sy-preview-shell {
            border:1px solid var(--sy-border);
            border-radius:18px;
            overflow:hidden;
            background: var(--secondary-background-color);
        }
        .sy-preview-topbar {
            display:flex;
            justify-content:space-between;
            align-items:center;
            padding:0.82rem 0.95rem;
            border-bottom:1px solid var(--sy-border);
            background: linear-gradient(
                90deg,
                color-mix(in srgb, var(--secondary-background-color) 88%, var(--primary-color) 8%),
                color-mix(in srgb, var(--secondary-background-color) 94%, transparent)
            );
        }
        .sy-preview-title { font-weight:700; font-size:0.98rem; color:var(--sy-text); }
        .sy-preview-meta { font-size:0.84rem; color:var(--sy-muted); }
        .sy-preview-badge {
            padding:0.28rem 0.58rem;
            border-radius:999px;
            font-size:0.76rem;
            font-weight:700;
            background:color-mix(in srgb, var(--primary-color) 18%, transparent);
            color:var(--sy-text);
            border:1px solid var(--sy-border);
        }
        .sy-preview-frame { border:0; background:white; }
        .sy-empty-preview {
            min-height:280px;
            display:flex;
            align-items:center;
            justify-content:center;
            text-align:center;
            border:1px dashed var(--sy-border);
            border-radius:18px;
            background:color-mix(in srgb, var(--secondary-background-color) 72%, transparent);
            padding:1rem;
            color:var(--sy-muted);
        }

        .sy-upload-item { border-radius:14px; padding:0.78rem 0.9rem; margin-bottom:0.55rem; }
        .sy-data-row {
            display:flex;
            justify-content:space-between;
            gap:0.8rem;
            padding:0.52rem 0;
            border-bottom:1px solid var(--sy-border);
            color:var(--sy-text);
        }
        .sy-data-row:last-child { border-bottom:0; }
        .sy-data-row span:first-child { color:var(--sy-muted); }

        div[data-testid="stMetric"] {
            background: linear-gradient(
                180deg,
                color-mix(in srgb, var(--secondary-background-color) 86%, transparent),
                color-mix(in srgb, var(--secondary-background-color) 76%, transparent)
            );
            border:1px solid var(--sy-border);
            padding:0.78rem 0.9rem;
            border-radius:16px;
        }

        .stDownloadButton button, .stButton button, .stLinkButton a {
            border-radius:14px !important;
        }
        .stButton button, .stDownloadButton button, .stLinkButton a {
            background: #D4C29A !important;
            color: #111111 !important;
            border: 1px solid #D4C29A !important;
            box-shadow: 0 10px 24px rgba(212, 194, 154, 0.18) !important;
            font-weight: 600 !important;
        }
        .stButton button:hover, .stDownloadButton button:hover, .stLinkButton a:hover {
            background: #c5b183 !important;
            border-color: #c5b183 !important;
            color: #111111 !important;
            filter:none !important;
        }
        .stButton button:focus, .stDownloadButton button:focus, .stLinkButton a:focus {
            outline: none !important;
            box-shadow: 0 0 0 2px rgba(212, 194, 154, 0.28) !important;
        }

        .stTabs [data-baseweb="tab-list"] { gap:0.42rem; }
        .stTabs [data-baseweb="tab"] {
            border-radius:12px;
            padding:0.5rem 0.9rem;
            background:color-mix(in srgb, var(--secondary-background-color) 80%, transparent);
            border:1px solid var(--sy-border);
        }
        .stTabs [aria-selected="true"] {
            color:var(--sy-text) !important;
            border-color:color-mix(in srgb, var(--primary-color) 28%, transparent) !important;
        }

        .stSelectbox label, .stTextInput label, .stTextArea label, .stNumberInput label, .stDateInput label, .stMultiSelect label {
            color:var(--sy-text) !important;
            font-weight:600 !important;
        }
        [data-baseweb="select"] > div,
        [data-baseweb="tag"] {
            background:color-mix(in srgb, var(--secondary-background-color) 86%, transparent) !important;
            border:1px solid #5D6472 !important;
            color:var(--sy-text) !important;
        }
        .stTextInput input,
        .stTextArea textarea,
        .stNumberInput input,
        .stDateInput input,
        div[data-baseweb="base-input"] > input,
        div[data-baseweb="base-input"] > textarea {
            background:color-mix(in srgb, var(--secondary-background-color) 86%, transparent) !important;
            border:1px solid #5D6472 !important;
            color:var(--sy-text) !important;
            border-radius:12px !important;
        }
        .stTextInput input:focus,
        .stTextArea textarea:focus,
        .stNumberInput input:focus,
        .stDateInput input:focus,
        div[data-baseweb="base-input"] > input:focus,
        div[data-baseweb="base-input"] > textarea:focus,
        [data-baseweb="select"] > div:focus-within {
            border-color:#D4C29A !important;
            box-shadow:0 0 0 1px #D4C29A !important;
        }
        .stTextArea textarea {
            min-height: 90px;
        }
        .streamlit-expanderHeader {
            border:1px solid #5D6472 !important;
            border-radius:12px !important;
        }

                .sy-topbar { margin-bottom: 0.8rem; }
        .sy-hero-simple { display:flex; gap:1rem; align-items:flex-start; }
        .sy-badge-row { display:flex; flex-wrap:wrap; gap:0.5rem; margin-top:0.7rem; }
        .sy-badge {
            padding:0.35rem 0.65rem;
            border:1px solid var(--sy-border);
            border-radius:999px;
            background:color-mix(in srgb, var(--secondary-background-color) 86%, transparent);
            font-size:0.8rem;
            color:var(--sy-text);
        }
        .sy-subtle-card {
            border:1px solid var(--sy-border);
            border-radius:18px;
            background: linear-gradient(
                180deg,
                color-mix(in srgb, var(--secondary-background-color) 86%, transparent),
                color-mix(in srgb, var(--secondary-background-color) 76%, transparent)
            );
            box-shadow: var(--sy-card-shadow);
            padding:0.95rem 1rem;
            margin-bottom:0.85rem;
        }
        .sy-section-label {
            font-size:0.78rem;
            text-transform:uppercase;
            letter-spacing:0.08em;
            color:var(--sy-muted);
            margin-bottom:0.35rem;
        }
.stProgress > div > div > div > div {
            background: linear-gradient(
                90deg,
                var(--primary-color),
                color-mix(in srgb, var(--primary-color) 65%, #00D4FF 35%)
            );
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




def extract_address_from_report(report_text: str, fallback: str = "Not provided") -> str:
    if not report_text:
        return fallback
    m = re.search(r"Project Address:\s*(.+)", report_text, re.IGNORECASE)
    if m:
        value = m.group(1).strip()
        if value and value.lower() != "not provided":
            return value
    m2 = re.search(r"\b\d+[A-Za-z]?\s+[^\n,]+(?:,\s*[^\n,]+){0,3},?\s*[A-Z]{1,2}\d[A-Z\d]?\s*\d[A-Z]{2}\b", report_text)
    if m2:
        return m2.group(0).strip()
    return fallback

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


def render_pdf_preview(uploaded_files):
    if not uploaded_files:
        st.markdown('<div class="sy-empty-preview">Upload a PDF to preview the drawing pack workspace.</div>', unsafe_allow_html=True)
        return

    selected_file = uploaded_files[-1]
    try:
        pdf_bytes = selected_file.getvalue()
        pdf_b64 = base64.b64encode(pdf_bytes).decode("utf-8")
        iframe = f'''
        <div class="sy-preview-shell">
            <div class="sy-preview-topbar">
                <div>
                    <div class="sy-preview-title">{selected_file.name}</div>
                    <div class="sy-preview-meta">{round(selected_file.size / (1024 * 1024), 2)} MB • Live drawing preview</div>
                </div>
                <div class="sy-preview-badge">PDF</div>
            </div>
            <iframe src="data:application/pdf;base64,{pdf_b64}" width="100%" height="760" type="application/pdf" class="sy-preview-frame"></iframe>
        </div>
        '''
        st.markdown(iframe, unsafe_allow_html=True)
    except Exception:
        st.markdown('<div class="sy-empty-preview">Preview not available for this file in the live browser view.</div>', unsafe_allow_html=True)

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
    left_margin = 46
    right_margin = 46
    top_margin = 56
    bottom_margin = 42
    usable_width = width - left_margin - right_margin
    y = height - top_margin

    def wrap_text(text, font_name="Helvetica", font_size=10.5, max_width=None):
        max_width = max_width or usable_width
        words = str(text).split()
        lines = []
        current = ""
        for word in words:
            test = word if not current else current + " " + word
            if stringWidth(test, font_name, font_size) <= max_width:
                current = test
            else:
                if current:
                    lines.append(current)
                current = word
        if current:
            lines.append(current)
        return lines or [""]

    def draw_page_header():
        c.setStrokeColor(colors.HexColor("#D8E0EE"))
        c.line(left_margin, height - 28, width - right_margin, height - 28)
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 10)
        c.drawString(left_margin, height - 18, practice_name or "ArchLens AI")
        c.setFont("Helvetica", 8.5)
        c.drawRightString(width - right_margin, height - 18, f"{report_title} | {report_id}")
        c.setFillColor(colors.black)

    def draw_page_footer():
        c.setStrokeColor(colors.HexColor("#D8E0EE"))
        c.line(left_margin, 24, width - right_margin, 24)
        c.setFillColor(colors.grey)
        c.setFont("Helvetica-Oblique", 8.2)
        footer_lines = wrap_text(disclaimer_text, "Helvetica-Oblique", 8.2, usable_width - 30)[:2]
        yy = 14
        for line in footer_lines:
            c.drawCentredString(width / 2, yy, line)
            yy -= 9
        c.setFillColor(colors.black)

    def new_page():
        nonlocal y
        c.showPage()
        y = height - top_margin
        draw_page_header()
        draw_page_footer()

    def ensure_space(required_height):
        nonlocal y
        if y - required_height < bottom_margin:
            new_page()

    def draw_cover():
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 24)
        c.drawCentredString(width / 2, height - 145, practice_name or "ArchLens AI")
        c.setFillColor(colors.black)
        c.setFont("Helvetica-Bold", 19)
        c.drawCentredString(width / 2, height - 175, report_title)
        c.setFont("Helvetica", 11)
        c.setFillColor(colors.HexColor("#404040"))
        meta_lines = [
            f"Project Address: {address}",
            f"Client: {client}",
            f"Drawing Pack Reviewed: {file_name}",
            f"Date: {date}",
            f"Report ID: {report_id}",
            f"Prepared by: {practice_name or 'ArchLens AI'}",
        ]
        yy = height - 255
        for line in meta_lines:
            c.drawCentredString(width / 2, yy, line)
            yy -= 18
        summary_lines = [ln.strip() for ln in sections.get("TOP SUMMARY", "").splitlines() if ln.strip()]
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 13)
        c.drawCentredString(width / 2, yy - 8, "Executive Summary")
        c.setFillColor(colors.black)
        c.setFont("Helvetica", 11)
        yy -= 30
        for line in summary_lines[:4]:
            c.drawCentredString(width / 2, yy, line)
            yy -= 18
        c.showPage()

    def draw_section_banner(title):
        nonlocal y
        ensure_space(58)
        banner_h = 26
        bottom = y - banner_h
        c.setFillColor(colors.HexColor("#E9EEF5"))
        c.roundRect(left_margin, bottom, usable_width, banner_h, 6, fill=1, stroke=0)
        c.setFillColor(colors.HexColor("#1F3B73"))
        c.setFont("Helvetica-Bold", 12)
        c.drawString(left_margin + 12, bottom + 8, title)
        c.setFillColor(colors.black)
        y = bottom - 12

    def draw_key_value_section(content):
        nonlocal y
        rows = parse_key_value_lines(content)
        for label, value in rows:
            if label:
                ensure_space(40)
                c.setFont("Helvetica-Bold", 10.5)
                for line in wrap_text(f"{label}:"):
                    c.drawString(left_margin, y, line)
                    y -= 15
                y -= 3
                c.setFont("Helvetica", 10.5)
                for line in wrap_text(value, "Helvetica", 10.5, usable_width - 14):
                    ensure_space(20)
                    c.drawString(left_margin + 12, y, line)
                    y -= 15
            else:
                c.setFont("Helvetica", 10.5)
                for line in wrap_text(value):
                    ensure_space(20)
                    c.drawString(left_margin, y, line)
                    y -= 15
            y -= 8

    def draw_bullet_section(content):
        nonlocal y
        for raw in content.splitlines():
            line = raw.strip()
            if not line:
                y -= 8
                continue
            bullet = False
            if line.startswith("- "):
                line = line[2:].strip()
                bullet = True
            elif line.startswith("• "):
                line = line[2:].strip()
                bullet = True
            wrapped = wrap_text(line, "Helvetica", 10.5, usable_width - (18 if bullet else 0))
            ensure_space((len(wrapped) * 15) + 10)
            c.setFont("Helvetica", 10.5)
            if bullet:
                c.drawString(left_margin, y, "•")
                for part in wrapped:
                    c.drawString(left_margin + 14, y, part)
                    y -= 15
                y -= 10
            else:
                for part in wrapped:
                    c.drawString(left_margin, y, part)
                    y -= 15
                y -= 10

    def draw_compliance_table(content):
        nonlocal y
        rows = parse_compliance_rows(content)
        if not rows:
            c.setFont("Helvetica", 10.5)
            c.drawString(left_margin, y, "No compliance status detected.")
            y -= 15
            return

        col_part = left_margin
        col_doc = left_margin + 42
        col_status = left_margin + 258
        col_why = left_margin + 392
        doc_w = 195
        why_w = usable_width - (col_why - left_margin) - 8

        def header():
            nonlocal y
            ensure_space(34)
            c.setFillColor(colors.HexColor("#EAEFF7"))
            c.rect(left_margin, y - 18, usable_width, 20, fill=1, stroke=0)
            c.setFillColor(colors.black)
            c.setFont("Helvetica-Bold", 9)
            c.drawString(col_part, y - 5, "Part")
            c.drawString(col_doc, y - 5, "Approved Document")
            c.drawString(col_status, y - 5, "Status")
            c.drawString(col_why, y - 5, "Why")
            y -= 30

        def status_fill(status):
            up = status.upper()
            if "PASS" in up and "PARTLY" not in up:
                return colors.HexColor("#2E7D32")
            if "FAIL" in up:
                return colors.HexColor("#C62828")
            return colors.HexColor("#EF6C00")

        header()
        for row in rows:
            doc_lines = wrap_text(row["title"], "Helvetica-Bold", 9, doc_w)
            why_lines = wrap_text(row["why"], "Helvetica", 8.5, why_w)
            row_h = max(40, 18 + max(len(doc_lines), len(why_lines), 1) * 10)
            ensure_space(row_h + 12)
            top = y
            c.setFillColor(colors.whitesmoke)
            c.rect(left_margin, y - row_h + 5, usable_width, row_h, fill=1, stroke=0)
            c.setFillColor(colors.HexColor("#1F3B73"))
            c.circle(col_part + 8, top - 6, 8, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            c.drawCentredString(col_part + 8, top - 9, row["part"])
            c.setFillColor(colors.black)

            c.setFont("Helvetica-Bold", 9)
            yy = top - 5
            for line in doc_lines:
                c.drawString(col_doc, yy, line)
                yy -= 10

            badge = row["status"].upper()
            if "PARTLY" in badge:
                badge = "REVIEW REQUIRED"
            elif "PASS" in badge and "PARTLY" not in badge:
                badge = "PASS"
            elif "FAIL" in badge:
                badge = "FAIL"
            bw = stringWidth(badge, "Helvetica-Bold", 8) + 12
            c.setFillColor(status_fill(badge))
            c.roundRect(col_status, top - 14, bw, 14, 3, fill=1, stroke=0)
            c.setFillColor(colors.white)
            c.setFont("Helvetica-Bold", 8)
            c.drawString(col_status + 6, top - 9, badge)
            c.setFillColor(colors.black)

            c.setFont("Helvetica", 8.5)
            yy = top - 5
            for line in why_lines:
                c.drawString(col_why, yy, line)
                yy -= 10

            y -= row_h + 10

    draw_cover()
    y = height - top_margin
    draw_page_header()
    draw_page_footer()

    for key, title in section_order:
        content = sections.get(key, "Not detected")
        estimated_lines = max(4, len([ln for ln in str(content).splitlines() if ln.strip()]))
        estimated_height = 52 + min(estimated_lines, 18) * 16
        ensure_space(estimated_height)
        draw_section_banner(title)
        if module_name == "Building Regulations Review" and key == "COMPLIANCE STATUS BY APPROVED DOCUMENT":
            draw_compliance_table(content)
        elif key in special_key_value_sections:
            draw_key_value_section(content)
        else:
            draw_bullet_section(content)
        y -= 8

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
    run.font.size = Pt(9)
    run.italic = True

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
        authority_value = top_summary_rows.get("LOCAL AUTHORITY", "Unknown")
        if authority_value == "Unknown":
            for line in sections.get("TOP SUMMARY", "").splitlines():
                stripped = line.strip()
                if stripped and ":" not in stripped:
                    authority_value = stripped
                    break
        return (
            top_summary_rows.get("OVERALL PLANNING POSITION", "Not shown"),
            top_summary_rows.get("LIKELY PLANNING ROUTE", top_summary_rows.get("APPLICATION TYPE", top_summary_rows.get("LIKELY ROUTE", "Unknown"))),
            authority_value,
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
    c2.metric("Planning Position" if module_name == "Planning Review" else "Risk Rating", v1)
    c3.metric(label_2, v2)
    c4.metric(label_3, v3)


def extract_summary_values(sections: Dict[str, str], module_name: str):
    top_summary_rows = {k.upper(): v for k, v in parse_key_value_lines(sections.get("TOP SUMMARY", "")) if k}
    if module_name == "Planning Review":
        authority_value = top_summary_rows.get("LOCAL AUTHORITY", "Unknown")
        if authority_value == "Unknown":
            for line in sections.get("TOP SUMMARY", "").splitlines():
                stripped = line.strip()
                if stripped and ":" not in stripped:
                    authority_value = stripped
                    break
        return {
            "risk": top_summary_rows.get("OVERALL PLANNING POSITION", "Not shown"),
            "route": top_summary_rows.get("LIKELY PLANNING ROUTE", top_summary_rows.get("APPLICATION TYPE", top_summary_rows.get("LIKELY ROUTE", "Unknown"))),
            "authority": authority_value,
            "probability": "Not shown",
        }
    return {
        "risk": top_summary_rows.get("OVERALL RISK RATING", "Unknown"),
        "route": top_summary_rows.get("SUBMISSION STATUS", "Unknown"),
        "authority": top_summary_rows.get("REVIEW CONFIDENCE", "Unknown"),
        "probability": "N/A",
    }


def detect_local_authority_for_display(project_address: str, proposal_summary: str, uploaded_files=None) -> str:
    combined_text = proposal_summary or ""
    if uploaded_files:
        names = " ".join(f.name for f in uploaded_files if getattr(f, "name", None))
        combined_text = f"{combined_text}\n{names}"
    return pdf_summary.detect_local_authority(project_address or "", combined_text or "")

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
    if module_name == "Planning Review":
        readiness_text = sections.get("SUBMISSION READINESS", "")
        if "READY TO SUBMIT" in readiness_text.upper():
            st.success("Submission position: Ready to submit")
        elif "LIKELY READY" in readiness_text.upper():
            st.warning("Submission position: Likely ready with minor amendments")
        else:
            st.info("Submission position: Further information required")
    else:
        top_summary_rows = {k.upper(): v for k, v in parse_key_value_lines(sections.get("TOP SUMMARY", "")) if k}
        risk_summary = top_summary_rows.get("OVERALL RISK RATING") or sections.get("TOP SUMMARY", "")
        summary_hint = top_summary_rows.get("REVIEW CONFIDENCE", "Unknown")
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


st.set_page_config(page_title="ArchLens AI", layout="wide")
inject_custom_css()

if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "auth_plan" not in st.session_state:
    st.session_state["auth_plan"] = "starter"
if "auth_user_name" not in st.session_state:
    st.session_state["auth_user_name"] = ""

token_plan, token_user_name, has_valid_token = get_verified_plan_and_user()

if has_valid_token:
    st.session_state["authenticated"] = True
    st.session_state["auth_plan"] = token_plan
    st.session_state["auth_user_name"] = token_user_name

if not st.session_state.get("authenticated", False):
    st.markdown(
        """
        <div class="sy-hero" style="max-width:900px;margin:3rem auto 1rem auto;">
            <div class="sy-hero-copy">
                <h1>ArchLens AI</h1>
                <div class="sy-muted" style="max-width:760px;">
                    Access is managed through your SY Design Studio member account.
                    Please launch ArchLens from your member area to verify your subscription and open the correct plan.
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True,
    )
    col1, col2 = st.columns(2)
    with col1:
        st.link_button("Login to Member Area", WEBSITE_LAUNCH_URL, use_container_width=True)
    with col2:
        st.link_button("View Plans", WEBSITE_PRICING_URL, use_container_width=True)
    st.stop()

current_plan = st.session_state.get("auth_plan", "starter")
current_user_name = st.session_state.get("auth_user_name", "")
allowed_review_modules = get_allowed_review_modules(current_plan)
default_module = st.session_state.active_module if st.session_state.active_module in allowed_review_modules else allowed_review_modules[0]
review_module = default_module
hero_welcome = f'<div style="font-size:0.92rem;color:#C7D7FF;margin-bottom:0.5rem;">Welcome {current_user_name}</div>' if current_user_name else ""

st.markdown(
    f"""
    <div class="sy-topbar">
        <div>
            <div class="sy-topbar-title">Architect AI Workspace</div>
            <div class="sy-topbar-meta">ArchLens AI • Drawing-focused planning and building regulations review</div>
        </div>
        <div class="sy-topbar-meta">Mode: {review_module} | Plan: {PLAN_LABELS.get(current_plan, "Solo")}{(" | User: " + current_user_name) if current_user_name else ""}</div>
    </div>
    <div class="sy-hero">
        <div class="sy-hero-simple">
            <div class="sy-hero-copy" style="max-width:760px;">
                <h1>ArchLens AI</h1>
                {hero_welcome}
                <div class="sy-muted" style="max-width:760px;">
                    Upload a drawing pack, confirm the project details, and generate a cleaner planning or building regulations review from one workspace.
                </div>
                <div class="sy-badge-row">
                    <div class="sy-badge">Drawing Pack Review</div>
                    <div class="sy-badge">Planning Route Logic</div>
                    <div class="sy-badge">Officer-Style Reports</div>
                </div>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

with st.sidebar:
    st.header("Project Setup")
    st.caption("Keep the setup light. Add extra detail only where it improves route accuracy.")
    st.caption(f"Current Plan: {PLAN_LABELS.get(current_plan, 'Solo')}")
    review_module = st.selectbox(
        "Review Module",
        allowed_review_modules,
        index=allowed_review_modules.index(default_module),
    )
    if current_plan == "starter":
        st.caption("Building Regulations Review is available on Studio.")
    project_types = st.multiselect("Project Type", PROJECT_TYPE_OPTIONS, default=[])

    if review_module == "Planning Review":
        property_type = st.selectbox("Property Type", PROPERTY_TYPE_OPTIONS, index=0)
    else:
        property_type = "Not stated"

    proposal_summary = st.text_area(
        "Proposal Description",
        height=110,
        placeholder="Briefly describe the proposal.",
    )


    rear_extension_depth_m = None
    rear_extension_height_m = None
    if review_module == "Planning Review" and "Ground Floor Rear Extension" in project_types:
        rear_extension_depth_m = st.number_input(
            "Rear extension depth from original rear wall (m)",
            min_value=0.0,
            max_value=12.0,
            value=6.0,
            step=0.1,
            help="Enter the proposed rear projection measured from the original rear house wall.",
        )
        rear_extension_height_m = st.number_input(
            "Rear extension overall height (m)",
            min_value=0.0,
            max_value=6.0,
            value=4.0,
            step=0.1,
            help="Enter the proposed overall height of the rear extension.",
        )

    review_mode = st.selectbox("Report Mode", ["Architect / Professional", "Homeowner Summary"])
    project_address = st.text_input("Project Address")
    local_authority = detect_local_authority_for_display(project_address, proposal_summary)
    practice_name = ""

    client_name, review_date, accuracy_answers = render_improve_accuracy_section(project_types)

    pd_route_label = "Not assessed"
    pd_risk_label = "Medium"
    pd_route_reason = "Add project details to improve the route snapshot."
    if review_module == "Planning Review":
        pd_route_label, pd_risk_label, pd_route_reason = get_planning_route_snapshot(
            project_types,
            property_type,
            proposal_summary,
            rear_extension_depth_m,
            rear_extension_height_m,
            accuracy_answers,
        )

    if st.button("Clear Report", key="clear_report_btn"):
        for key, value in DEFAULT_STATE.items():
            st.session_state[key] = value
        st.session_state["planning_statement_text"] = None
        st.session_state["planning_statement_file"] = None
        st.info("Stored report cleared.")

st.session_state.active_module = review_module
if "starter_review_count" not in st.session_state:
    st.session_state["starter_review_count"] = 0
config = MODULE_CONFIG[review_module]
if current_plan == "starter" and review_module == "Building Regulations Review":
    st.warning(get_plan_upgrade_message("Building Regulations Review"))
    st.stop()

setup_tab, upload_tab, report_tab = st.tabs(["Project Setup", "Upload Drawing Pack", "AI Review Report"])

with setup_tab:
    st.markdown(f'<div class="sy-subtle-card"><div class="sy-section-label">Review Summary</div><h3 style="margin:0 0 0.35rem 0;">{config["title"]}</h3><div class="sy-muted">{config["disclaimer"]}</div></div>', unsafe_allow_html=True)
    if review_module == "Planning Review":
        st.markdown(f'<div class="sy-subtle-card"><strong>Auto route:</strong> {pd_route_label} &nbsp;&nbsp; <strong>Risk:</strong> {pd_risk_label}<br><span class="sy-muted">{pd_route_reason}</span></div>', unsafe_allow_html=True)
    c1, c2 = st.columns([1.15, 0.85])
    with c1:
        st.markdown("**Current setup**")
        st.write(f"Review module: {review_module}")
        st.write(f"Report mode: {review_mode}")
        st.write(f"Project type: {', '.join(project_types) if project_types else 'Not stated'}")
        if review_module == "Planning Review":
            st.write(f"Property type: {property_type or 'Not stated'}")
            st.write(f"Local authority: {local_authority}")
        st.write(f"Project address: {project_address or 'Not provided'}")
        st.write(f"Proposal description: {proposal_summary or 'Not provided'}")
        st.write(f"Client: {client_name or 'Not provided'}")
    with c2:
        if review_module == "Planning Review":
            st.info("Use this module for officer-style reasoning, street precedent review, proposal recognition, PD / prior approval / full planning route review, route confidence scoring, and planning statement drafting.")
            if current_plan == "starter":
                st.caption("Solo includes planning review only and PDF exports.")
        else:
            st.info("Use this module for technical Building Regulations review including plans, sections, details, specifications, and structural sheets.")
            if current_plan == "starter":
                st.warning("Building Regulations Review is available on Studio only.")

with upload_tab:
    st.markdown('<div class="sy-subtle-card"><div class="sy-section-label">Upload + Analyse</div><h3 style="margin:0 0 0.35rem 0;">Drawing Workspace</h3><div class="sy-muted">Upload the active drawing pack, preview it, and run the review once the setup is complete.</div></div>', unsafe_allow_html=True)

    uploaded_file = st.file_uploader(
        "Upload Drawing PDF",
        type=["pdf"],
        accept_multiple_files=True,
        key="drawing_pdf_upload",
    )

    workspace_col, assistant_col = st.columns([1.75, 0.85], gap="large")

    with workspace_col:
        st.markdown('<div class="sy-workspace">', unsafe_allow_html=True)
        st.markdown('<div class="sy-workspace-header"><div><div class="sy-workspace-title">Drawing Pack Workspace</div><div class="sy-workspace-meta">Visual preview area for the active PDF and uploaded drawing set.</div></div></div>', unsafe_allow_html=True)

        if review_module == "Planning Review" and review_mode == "Homeowner Summary":
            st.info("Homeowners can upload a simple sketch or basic PDF. ArchLens AI will frame the output as a preliminary planning feasibility review, not a formal planning decision.")

        preview_col, meta_col = st.columns([1.45, 0.75], gap="large")
        with preview_col:
            render_pdf_preview(uploaded_file)
        with meta_col:
            st.markdown("**Uploaded files**")
            if uploaded_file:
                for file in uploaded_file:
                    file_size_mb = round(file.size / (1024 * 1024), 2)
                    st.markdown(
                        f'<div class="sy-upload-item"><strong>{file.name}</strong><br><span class="sy-muted">{file_size_mb} MB</span></div>',
                        unsafe_allow_html=True,
                    )
            else:
                st.markdown('<div class="sy-empty-preview">No drawing pack uploaded yet.</div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

    with assistant_col:
        st.markdown('<div class="sy-sidepanel">', unsafe_allow_html=True)
        st.markdown('<div class="sy-panel-title">Project Snapshot</div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sy-data-row"><span>Review module</span><strong>{review_module}</strong></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sy-data-row"><span>Report mode</span><strong>{review_mode}</strong></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sy-data-row"><span>Project type</span><strong>{", ".join(project_types) if project_types else "Not stated"}</strong></div>', unsafe_allow_html=True)
        if review_module == "Planning Review":
            st.markdown(f'<div class="sy-data-row"><span>Property type</span><strong>{property_type or "Not stated"}</strong></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Local authority</span><strong>{local_authority or "Not clearly identified"}</strong></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sy-data-row"><span>Project address</span><strong>{project_address or "Not provided"}</strong></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="sy-data-row"><span>Client</span><strong>{client_name or "Not provided"}</strong></div>', unsafe_allow_html=True)
        st.markdown("")
        if current_plan == "starter":
            st.info("Solo plan: Planning Review + PDF exports. Upgrade to Studio for Building Regulations Review, Word exports, and unlimited reviews.")
            st.link_button("Upgrade to Studio", WEBSITE_PRICING_URL, use_container_width=True)
            st.link_button("Return to SY Design Studio", WEBSITE_HOME_URL, use_container_width=True)

        if uploaded_file:
            total_uploaded_mb = sum(f.size for f in uploaded_file) / (1024 * 1024)
            if total_uploaded_mb > 20:
                st.error("Drawing pack too large for the live app. Please keep the total upload size to 20MB or less, or split the pack into smaller PDFs.")
                st.stop()

            st.metric("Files uploaded", len(uploaded_file))
            st.metric("Total size", f"{round(total_uploaded_mb, 2)} MB")
            st.metric("Selected project types", len(project_types))
            run_analysis = st.button("Analyse Drawing Pack", key="run_review_btn", use_container_width=True)
        else:
            st.markdown('<div class="sy-muted">Upload a drawing pack to enable the AI review controls.</div>', unsafe_allow_html=True)
            run_analysis = False
        st.markdown('</div>', unsafe_allow_html=True)

    if uploaded_file and run_analysis:
        if current_plan == "starter" and st.session_state.get("starter_review_count", 0) >= STARTER_MONTHLY_REVIEW_LIMIT:
            st.error("You have reached your 10 monthly reviews on Solo. Upgrade to Studio for unlimited project reviews.")
            st.stop()
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
            if page_count > 12:
                st.warning("Large drawing pack detected. The live app will analyse the first 12 pages only to keep processing stable.")

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
                    smooth_progress(progress_bar, status_text, 25, 40,
                                    "Reading drawings and extracting planning data...", 0.8)

                    proposal_summary_for_ai = proposal_summary
                    if "Ground Floor Rear Extension" in project_types:
                        depth_txt = f"{rear_extension_depth_m:.1f}m depth from original rear wall" if rear_extension_depth_m is not None else ""
                        height_txt = f"{rear_extension_height_m:.1f}m overall height" if rear_extension_height_m is not None else ""
                        extra_bits = ", ".join([x for x in [depth_txt, height_txt] if x])
                        if extra_bits:
                            proposal_summary_for_ai = (proposal_summary_for_ai.strip() + " | " + extra_bits).strip(" |")

                    pd_context = build_pd_context(project_types, property_type, rear_extension_depth_m, rear_extension_height_m, accuracy_answers)
                    accuracy_context = build_accuracy_context(accuracy_answers)
                    if accuracy_context:
                        proposal_summary_for_ai = (proposal_summary_for_ai.strip() + " | Improve Accuracy: " + accuracy_context).strip(" |")

                    report = pdf_summary.analyze_planning_pdf(
                        temp_pdf_path,
                        client_project_types=project_types,
                        property_type=property_type,
                        proposal_summary=proposal_summary_for_ai,
                        project_address=project_address,
                        local_authority=local_authority,
                        review_mode=review_mode,
                        pd_context=pd_context,
                    )

                    smooth_progress(progress_bar, status_text, 40, 85, "Analyzing planning route and risks...", 0.8)
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

            extracted_report_address = extract_address_from_report(report, "Not provided")
            clean_project_address = clean_input_value(project_address, extracted_report_address)
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
            if current_plan == "starter":
                st.session_state["starter_review_count"] = st.session_state.get("starter_review_count", 0) + 1

            add_saved_project(
                {
                    "report_id": report_id,
                    "project_address": clean_project_address,
                    "client_name": clean_client_name,
                    "module": review_module,
                    "project_types": ", ".join(project_types) if project_types else "Not stated",
                    "property_type": property_type if review_module == "Planning Review" else "Not stated",
                    "filename": file.name,
                    "date": str(review_date),
                    "plan": PLAN_LABELS.get(current_plan, "Solo"),
                }
            )

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
            gc.collect()

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

        st.markdown(f'<div class="sy-subtle-card"><div class="sy-section-label">Review Output</div><h3 style="margin:0 0 0.35rem 0;">{panel_title}</h3><div class="sy-muted">{panel_note}</div></div>', unsafe_allow_html=True)

        report_col, insight_col = st.columns([1.65, 0.95], gap="large")

        with report_col:
            render_at_a_glance(sections, report_id, review_module)
            st.markdown("")
            render_sections(sections, report, review_module)

        with insight_col:
            values = extract_summary_values(sections, review_module)
            st.markdown('<div class="sy-sidepanel">', unsafe_allow_html=True)
            st.markdown('<div class="sy-panel-title">AI Insights Panel</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Report ID</span><strong>{report_id}</strong></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Review module</span><strong>{review_module}</strong></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Plan</span><strong>{PLAN_LABELS.get(current_plan, "Solo")}</strong></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Application / status</span><strong>{values["route"]}</strong></div>', unsafe_allow_html=True)
            if review_module == "Planning Review":
                st.markdown(f'<div class="sy-data-row"><span>Local authority</span><strong>{values["authority"]}</strong></div>', unsafe_allow_html=True)
                st.markdown(f'<div class="sy-data-row"><span>Route confidence</span><strong>{values["probability"]}</strong></div>', unsafe_allow_html=True)
            else:
                st.markdown(f'<div class="sy-data-row"><span>Authority / confidence</span><strong>{values["authority"]}</strong></div>', unsafe_allow_html=True)
            st.markdown(f'<div class="sy-data-row"><span>Submission position</span><strong>{"Ready to submit" if "READY TO SUBMIT" in sections.get(config["readiness_key"], "").upper() else "Review required"}</strong></div>', unsafe_allow_html=True)
            st.markdown("")
            st.markdown('<div class="sy-panel-title">Quick Actions</div>', unsafe_allow_html=True)

            base_filename = (st.session_state.last_filename or "drawing_pack").rsplit(".", 1)[0]
            suffix = "Planning" if review_module == "Planning Review" else "BuildingRegs"

            if current_plan == "pro":
                st.download_button(
                    label=("Download Homeowner Feasibility Report (.docx)" if review_module == "Planning Review" and review_mode == "Homeowner Summary" else "Download Professional Report (.docx)"),
                    data=word_file,
                    file_name=f"{base_filename}_{suffix}_AI_Review_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key="download_docx",
                    use_container_width=True,
                )
            else:
                st.button("Download Word Report (.docx) 🔒 Studio", key="download_docx_locked", disabled=True, use_container_width=True)
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
                if current_plan == "starter":
                    st.caption("Solo includes planning statement drafting, but Word download is available on Studio only.")
                if st.button("Generate Planning Statement", key="generate_planning_statement_btn", use_container_width=True):
                    statement_text = pdf_summary.generate_planning_statement(
                        report_text=report,
                        sections=sections,
                        project_address=project_address or "Not provided",
                        client_name=client_name or "Not provided",
                        local_authority=local_authority or "",
                        review_mode=review_mode,
                    )
                    st.session_state["planning_statement_text"] = statement_text
                    st.session_state["planning_statement_file"] = build_simple_word_doc("Draft Planning Statement", statement_text)

                if st.session_state.get("planning_statement_text"):
                    if current_plan == "pro":
                        st.download_button(
                            label="Download Planning Statement (.docx)",
                            data=st.session_state["planning_statement_file"],
                            file_name=f"{base_filename}_Planning_Statement.docx",
                            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                            key="download_planning_statement_docx",
                            use_container_width=True,
                        )
                    else:
                        st.button("Download Planning Statement (.docx) 🔒 Studio", key="download_planning_statement_docx_locked", disabled=True, use_container_width=True)
            if current_plan == "starter":
                st.markdown(f'<div class="sy-data-row"><span>Monthly reviews used</span><strong>{st.session_state.get("starter_review_count", 0)} / {STARTER_MONTHLY_REVIEW_LIMIT}</strong></div>', unsafe_allow_html=True)
                st.link_button("Upgrade to Studio", WEBSITE_PRICING_URL, use_container_width=True)
            st.link_button("Return to SY Design Studio", WEBSITE_HOME_URL, use_container_width=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("")
        with st.expander("Saved Projects History", expanded=False):
            saved_projects = st.session_state.get("saved_projects", [])
            if saved_projects:
                for item in saved_projects:
                    st.write(item.get("project_address", "Not provided"))
                    st.caption(f"Report ID: {item.get('report_id', '')}")
                    st.caption(f"Module: {item.get('module', '')}")
                    st.caption(f"Project Type: {item.get('project_types', '')}")
                    st.caption(f"File: {item.get('filename', '')}")
                    st.caption(f"Date: {item.get('date', '')}")
                    st.caption(f"Plan: {item.get('plan', '')}")
                    st.markdown("---")
            else:
                st.info("No saved projects yet. Run a review to build your project history.")

        if review_module == "Planning Review" and st.session_state.get("planning_statement_text"):
            st.markdown("")
            with st.expander("Show planning statement draft", expanded=False):
                st.text(st.session_state["planning_statement_text"])
    else:
        st.info("No report generated yet. Complete the setup, upload the drawing pack, and run the review.")
