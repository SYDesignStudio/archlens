"""
ArchLens AI - Planning Rules Engine
SY Design Studio

Purpose:
- Deterministic pass / fail / needs-confirmation checks for common UK householder
  permitted development and prior approval routes.
- Based on the MHCLG / DLUHC 'Permitted development rights for householders'
  Technical Guidance and GPDO Schedule 2, Part 1.

Important:
- This is not a legal decision engine.
- It should be used to stabilise AI output, reduce hallucination, and provide
  clear rule-based checks before the report is written.
- Where data is missing, the rule returns NEEDS_CONFIRMATION rather than guessing.

Recommended integration:
    from planning_rules import run_planning_rule_checks, format_rule_checks_for_prompt

    rule_result = run_planning_rule_checks(
        project_types=["Ground Floor Rear Extension"],
        property_type="Semi-Detached House",
        measurements={"rear_depth_m": 6.0, "overall_height_m": 4.0, "eaves_height_m": 3.0},
        pd_answers={"within_2m_boundary": "yes", "materials_similar": "yes"},
        detected_features={"single_storey_rear_extension": True},
    )

    prompt_context = format_rule_checks_for_prompt(rule_result)
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Any, Dict, Iterable, List, Optional, Tuple


class RuleStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    NEEDS_CONFIRMATION = "NEEDS CONFIRMATION"
    NOT_APPLICABLE = "NOT APPLICABLE"


class RouteStatus(str, Enum):
    PD_POSSIBLE = "PD / LDC possible"
    PRIOR_APPROVAL_POSSIBLE = "Prior Approval possible"
    FULL_PLANNING_LIKELY = "Full Planning likely"
    NEEDS_CONFIRMATION = "Needs confirmation"


@dataclass
class RuleCheck:
    code: str
    title: str
    status: RuleStatus
    rule: str
    evidence: str = ""
    action: str = ""
    source: str = "Permitted development rights for householders Technical Guidance, GPDO Schedule 2 Part 1"

    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data["status"] = self.status.value
        return data


@dataclass
class RuleEngineResult:
    likely_route: RouteStatus
    confidence: str
    summary: str
    checks: List[RuleCheck] = field(default_factory=list)
    failed_checks: List[RuleCheck] = field(default_factory=list)
    needs_confirmation: List[RuleCheck] = field(default_factory=list)
    passed_checks: List[RuleCheck] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "likely_route": self.likely_route.value,
            "confidence": self.confidence,
            "summary": self.summary,
            "checks": [c.to_dict() for c in self.checks],
            "failed_checks": [c.to_dict() for c in self.failed_checks],
            "needs_confirmation": [c.to_dict() for c in self.needs_confirmation],
            "passed_checks": [c.to_dict() for c in self.passed_checks],
        }


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------


def _normalise(value: Any) -> str:
    return str(value or "").strip().lower().replace("_", "-")


def _contains_any(text: str, needles: Iterable[str]) -> bool:
    text_l = _normalise(text)
    return any(n.lower() in text_l for n in needles)


def _as_float(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().lower().replace("metres", "m").replace("meter", "m")
    text = text.replace(",", "")
    for suffix in ["mm", "m"]:
        if text.endswith(suffix):
            text = text[: -len(suffix)].strip()
            try:
                number = float(text)
                return number / 1000 if suffix == "mm" else number
            except Exception:
                return None
    try:
        return float(text)
    except Exception:
        return None


def _get_measure(measurements: Dict[str, Any], *keys: str) -> Optional[float]:
    for key in keys:
        if key in measurements:
            value = _as_float(measurements.get(key))
            if value is not None:
                return value
    return None


def _get_answer(pd_answers: Dict[str, Any], *keys: str) -> str:
    for key in keys:
        if key in pd_answers and pd_answers.get(key) not in [None, ""]:
            return _normalise(pd_answers.get(key))
    return ""


def _is_yes(value: str) -> bool:
    return value in {"yes", "y", "true", "1"}


def _is_no(value: str) -> bool:
    return value in {"no", "n", "false", "0"}


def _property_category(property_type: str) -> str:
    p = _normalise(property_type)
    if "flat" in p or "maisonette" in p:
        return "flat_or_maisonette"
    if "detached" in p and "semi" not in p:
        return "detached"
    if "semi" in p:
        return "semi"
    if "terrace" in p or "terraced" in p:
        return "terrace"
    if "bungalow" in p:
        return "bungalow"
    return "other"


def _project_text(project_types: List[str], detected_features: Optional[Dict[str, Any]] = None) -> str:
    features = detected_features or {}
    active_features = [k for k, v in features.items() if bool(v)]
    return " ".join(project_types or []) + " " + " ".join(active_features)


def _make_check(
    code: str,
    title: str,
    status: RuleStatus,
    rule: str,
    evidence: str = "",
    action: str = "",
) -> RuleCheck:
    return RuleCheck(code=code, title=title, status=status, rule=rule, evidence=evidence, action=action)


# -----------------------------------------------------------------------------
# General checks applying to householder PD
# -----------------------------------------------------------------------------


def check_general_pd_eligibility(property_type: str, pd_answers: Dict[str, Any]) -> List[RuleCheck]:
    checks: List[RuleCheck] = []
    category = _property_category(property_type)

    if category == "flat_or_maisonette":
        checks.append(
            _make_check(
                "GEN-01",
                "Dwellinghouse eligibility",
                RuleStatus.FAIL,
                "Householder permitted development rights do not normally apply to flats or maisonettes.",
                f"Property type: {property_type}",
                "Use full planning route unless a separate lawful route is confirmed.",
            )
        )
    elif not property_type or _normalise(property_type) in {"not stated", "unknown"}:
        checks.append(
            _make_check(
                "GEN-01",
                "Dwellinghouse eligibility",
                RuleStatus.NEEDS_CONFIRMATION,
                "Householder PD rights apply to dwellinghouses, not flats/maisonettes.",
                "Property type not confirmed.",
                "Confirm whether the site is a single dwellinghouse, flat, maisonette or converted property.",
            )
        )
    else:
        checks.append(
            _make_check(
                "GEN-01",
                "Dwellinghouse eligibility",
                RuleStatus.PASS,
                "Householder PD rights may apply to a single dwellinghouse, subject to restrictions.",
                f"Property type: {property_type}",
            )
        )

    constraints = _get_answer(pd_answers, "site_constraints", "constraints")
    if any(x in constraints for x in ["listed", "article 4"]):
        checks.append(
            _make_check(
                "GEN-02",
                "PD rights restrictions",
                RuleStatus.FAIL,
                "Listed status, Article 4 Directions, or planning conditions may remove/restrict PD rights.",
                f"Constraint input: {constraints}",
                "Confirm constraints and use full planning route if PD rights are removed.",
            )
        )
    elif any(x in constraints for x in ["conservation", "article 2", "aonb", "national park", "world heritage"]):
        checks.append(
            _make_check(
                "GEN-02",
                "Article 2(3) land / conservation constraint",
                RuleStatus.NEEDS_CONFIRMATION,
                "Article 2(3) land restricts several PD classes including side extensions and roof enlargements.",
                f"Constraint input: {constraints}",
                "Confirm conservation area / Article 2(3) status before relying on PD.",
            )
        )
    elif constraints:
        checks.append(
            _make_check(
                "GEN-02",
                "PD rights restrictions",
                RuleStatus.PASS,
                "No PD-rights restriction has been identified from the provided answers.",
                f"Constraint input: {constraints}",
            )
        )
    else:
        checks.append(
            _make_check(
                "GEN-02",
                "PD rights restrictions",
                RuleStatus.NEEDS_CONFIRMATION,
                "Article 4 Directions, conservation area status and planning conditions must be checked before relying on PD.",
                "No constraint answer provided.",
                "Check planning history and constraints map.",
            )
        )

    change_use_origin = _get_answer(pd_answers, "created_by_pd_change_of_use", "class_m_n_p_pa_q")
    if _is_yes(change_use_origin):
        checks.append(
            _make_check(
                "GEN-03",
                "House created through Part 3 change of use",
                RuleStatus.FAIL,
                "PD rights under Part 1 do not apply where the dwellinghouse was created under Classes M, N, P, PA or Q of Part 3.",
                "User indicated the dwelling may have been created through a Part 3 change of use.",
                "Use full planning route unless lawful PD eligibility is confirmed.",
            )
        )
    elif _is_no(change_use_origin):
        checks.append(
            _make_check(
                "GEN-03",
                "House created through Part 3 change of use",
                RuleStatus.PASS,
                "No Part 3 change-of-use origin has been identified.",
                "User indicated no.",
            )
        )
    else:
        checks.append(
            _make_check(
                "GEN-03",
                "House created through Part 3 change of use",
                RuleStatus.NEEDS_CONFIRMATION,
                "PD rights under Part 1 may not apply to houses created under Classes M, N, P, PA or Q of Part 3.",
                "Not confirmed.",
                "Confirm planning history / original use.",
            )
        )
    return checks


# -----------------------------------------------------------------------------
# Class A - extensions / enlargements
# -----------------------------------------------------------------------------


def check_class_a_extension(
    project_types: List[str],
    property_type: str,
    measurements: Dict[str, Any],
    pd_answers: Dict[str, Any],
    detected_features: Optional[Dict[str, Any]] = None,
) -> List[RuleCheck]:
    checks: List[RuleCheck] = []
    text = _project_text(project_types, detected_features)
    category = _property_category(property_type)

    class_a_relevant = _contains_any(
        text,
        [
            "rear extension",
            "side extension",
            "infill extension",
            "wraparound",
            "wrap around",
            "first floor extension",
            "single-storey rear extension",
            "single_storey_rear_extension",
        ],
    )
    if not class_a_relevant:
        return [
            _make_check(
                "A-00",
                "Class A relevance",
                RuleStatus.NOT_APPLICABLE,
                "Class A applies to enlargement, improvement or alteration to a house.",
                "No Class A extension feature identified.",
            )
        ]

    # Forward of principal elevation / side elevation fronting highway
    forward = _get_answer(pd_answers, "forward_of_principal_elevation", "projects_forward_principal")
    if _is_yes(forward):
        checks.append(
            _make_check(
                "A-01",
                "Principal elevation / highway side elevation",
                RuleStatus.FAIL,
                "Class A does not permit enlargement beyond the principal elevation or a side elevation fronting a highway.",
                "User/plans indicate projection beyond the principal elevation or highway-facing side elevation.",
                "Full planning permission is likely required.",
            )
        )
    elif _is_no(forward):
        checks.append(
            _make_check(
                "A-01",
                "Principal elevation / highway side elevation",
                RuleStatus.PASS,
                "The works must not project beyond the principal elevation or a highway-facing side elevation.",
                "No forward/highway-side projection indicated.",
            )
        )
    else:
        checks.append(
            _make_check(
                "A-01",
                "Principal elevation / highway side elevation",
                RuleStatus.NEEDS_CONFIRMATION,
                "The works must not project beyond the principal elevation or a highway-facing side elevation.",
                "Not confirmed from inputs.",
                "Check proposed site/block plan and elevations.",
            )
        )

    overall_height = _get_measure(measurements, "overall_height_m", "height_m", "maximum_height_m")
    if overall_height is None:
        checks.append(
            _make_check(
                "A-02",
                "Overall height",
                RuleStatus.NEEDS_CONFIRMATION,
                "Single-storey Class A rear/side extensions must not exceed 4.0m in height.",
                "Overall height not confirmed.",
                "Confirm height from proposed elevations/sections.",
            )
        )
    elif overall_height <= 4.0:
        checks.append(
            _make_check(
                "A-02",
                "Overall height",
                RuleStatus.PASS,
                "Single-storey Class A rear/side extensions must not exceed 4.0m in height.",
                f"Overall height: {overall_height:.2f}m.",
            )
        )
    else:
        checks.append(
            _make_check(
                "A-02",
                "Overall height",
                RuleStatus.FAIL,
                "Single-storey Class A rear/side extensions must not exceed 4.0m in height.",
                f"Overall height: {overall_height:.2f}m.",
                "Reduce height or use full planning route.",
            )
        )

    # Eaves where within 2m boundary
    within_2m = _get_answer(pd_answers, "within_2m_boundary", "within_2m_of_boundary")
    eaves_height = _get_measure(measurements, "eaves_height_m", "boundary_eaves_height_m")
    if _is_yes(within_2m):
        if eaves_height is None:
            checks.append(
                _make_check(
                    "A-03",
                    "Eaves height within 2m of boundary",
                    RuleStatus.NEEDS_CONFIRMATION,
                    "If any part of the extension is within 2m of the boundary, eaves must not exceed 3.0m.",
                    "Within 2m of boundary indicated, but eaves height not confirmed.",
                    "Confirm eaves height from section/elevation.",
                )
            )
        elif eaves_height <= 3.0:
            checks.append(
                _make_check(
                    "A-03",
                    "Eaves height within 2m of boundary",
                    RuleStatus.PASS,
                    "If any part of the extension is within 2m of the boundary, eaves must not exceed 3.0m.",
                    f"Boundary eaves height: {eaves_height:.2f}m.",
                )
            )
        else:
            checks.append(
                _make_check(
                    "A-03",
                    "Eaves height within 2m of boundary",
                    RuleStatus.FAIL,
                    "If any part of the extension is within 2m of the boundary, eaves must not exceed 3.0m.",
                    f"Boundary eaves height: {eaves_height:.2f}m.",
                    "Reduce eaves height or use full planning route.",
                )
            )
    elif _is_no(within_2m):
        checks.append(
            _make_check(
                "A-03",
                "Eaves height within 2m of boundary",
                RuleStatus.NOT_APPLICABLE,
                "The 3.0m boundary eaves limit applies where the extension is within 2m of the boundary.",
                "Extension not indicated within 2m of boundary.",
            )
        )
    else:
        checks.append(
            _make_check(
                "A-03",
                "Eaves height within 2m of boundary",
                RuleStatus.NEEDS_CONFIRMATION,
                "If any part of the extension is within 2m of the boundary, eaves must not exceed 3.0m.",
                "Boundary relationship not confirmed.",
                "Confirm boundary distances and eaves height.",
            )
        )

    # Rear extension depth / prior approval
    rear_relevant = _contains_any(text, ["rear extension", "single-storey rear extension", "single_storey_rear_extension"])
    if rear_relevant:
        depth = _get_measure(measurements, "rear_depth_m", "depth_m", "projection_m")
        if category == "detached":
            standard_limit, prior_limit = 4.0, 8.0
            prop_label = "detached"
        else:
            standard_limit, prior_limit = 3.0, 6.0
            prop_label = "non-detached"
        if depth is None:
            checks.append(
                _make_check(
                    "A-04",
                    "Single-storey rear extension depth",
                    RuleStatus.NEEDS_CONFIRMATION,
                    f"For a {prop_label} house, standard Class A rear extension depth is {standard_limit:.0f}m; larger home extension prior approval may allow up to {prior_limit:.0f}m.",
                    "Rear extension depth not confirmed.",
                    "Confirm depth from original rear wall, not from existing extension unless it is original.",
                )
            )
        elif depth <= standard_limit:
            checks.append(
                _make_check(
                    "A-04",
                    "Single-storey rear extension depth",
                    RuleStatus.PASS,
                    f"For a {prop_label} house, standard Class A rear extension depth must not exceed {standard_limit:.0f}m.",
                    f"Rear projection: {depth:.2f}m.",
                )
            )
        elif depth <= prior_limit:
            checks.append(
                _make_check(
                    "A-04",
                    "Larger home extension prior approval depth",
                    RuleStatus.PASS,
                    f"For a {prop_label} house, larger home extension prior approval may allow rear projection up to {prior_limit:.0f}m, subject to neighbour consultation and Class A conditions.",
                    f"Rear projection: {depth:.2f}m.",
                    "Prior approval route likely required rather than standard PD/LDC.",
                )
            )
        else:
            checks.append(
                _make_check(
                    "A-04",
                    "Single-storey rear extension depth",
                    RuleStatus.FAIL,
                    f"For a {prop_label} house, the larger home extension prior approval depth limit is {prior_limit:.0f}m.",
                    f"Rear projection: {depth:.2f}m.",
                    "Full planning permission is likely required.",
                )
            )

    # Side extension width/height/storey
    side_relevant = _contains_any(text, ["side extension", "infill extension", "wraparound", "wrap around"])
    if side_relevant:
        side_width_ratio = _get_measure(measurements, "side_width_ratio", "side_width_fraction")
        side_more_than_half = _get_answer(pd_answers, "side_extension_width", "side_more_than_half_width")
        if side_width_ratio is not None:
            if side_width_ratio <= 0.5:
                status = RuleStatus.PASS
                evidence = f"Side extension width ratio: {side_width_ratio:.2f} of original house width."
                action = ""
            else:
                status = RuleStatus.FAIL
                evidence = f"Side extension width ratio: {side_width_ratio:.2f} of original house width."
                action = "Reduce width or use full planning route."
        elif _is_yes(side_more_than_half):
            status = RuleStatus.FAIL
            evidence = "Input indicates side extension is more than half the width of the original house."
            action = "Full planning permission is likely required."
        elif _is_no(side_more_than_half):
            status = RuleStatus.PASS
            evidence = "Input indicates side extension is not more than half the original house width."
            action = ""
        else:
            status = RuleStatus.NEEDS_CONFIRMATION
            evidence = "Side extension width not confirmed."
            action = "Confirm original house width and proposed side extension width."
        checks.append(
            _make_check(
                "A-05",
                "Side extension width",
                status,
                "A side extension must not be wider than half the width of the original dwellinghouse.",
                evidence,
                action,
            )
        )

        more_than_single_storey = _get_answer(pd_answers, "side_more_than_single_storey", "side_extension_more_than_single_storey")
        if _is_yes(more_than_single_storey):
            checks.append(
                _make_check(
                    "A-06",
                    "Side extension storey limit",
                    RuleStatus.FAIL,
                    "A side extension under Class A must be single storey.",
                    "Input indicates more than single storey.",
                    "Full planning permission is likely required.",
                )
            )
        elif _is_no(more_than_single_storey):
            checks.append(
                _make_check(
                    "A-06",
                    "Side extension storey limit",
                    RuleStatus.PASS,
                    "A side extension under Class A must be single storey.",
                    "Input indicates single storey.",
                )
            )
        else:
            checks.append(
                _make_check(
                    "A-06",
                    "Side extension storey limit",
                    RuleStatus.NEEDS_CONFIRMATION,
                    "A side extension under Class A must be single storey.",
                    "Storey height not confirmed.",
                    "Confirm from plans/elevations.",
                )
            )

    # More than one storey rear / first floor
    first_floor_relevant = _contains_any(text, ["first floor extension", "two storey", "2 storey", "upper storey"])
    if first_floor_relevant:
        rear_depth = _get_measure(measurements, "rear_depth_m", "depth_m", "projection_m")
        rear_boundary_distance = _get_measure(measurements, "rear_boundary_distance_m", "distance_to_rear_boundary_m")
        if rear_depth is None:
            depth_status = RuleStatus.NEEDS_CONFIRMATION
            depth_ev = "Upper-storey rear projection not confirmed."
            depth_action = "Confirm projection from original rear wall."
        elif rear_depth <= 3.0:
            depth_status = RuleStatus.PASS
            depth_ev = f"Upper-storey rear projection: {rear_depth:.2f}m."
            depth_action = ""
        else:
            depth_status = RuleStatus.FAIL
            depth_ev = f"Upper-storey rear projection: {rear_depth:.2f}m."
            depth_action = "Full planning permission is likely required."
        checks.append(
            _make_check(
                "A-07",
                "More than single-storey rear projection",
                depth_status,
                "A rear enlargement with more than one storey must not extend beyond the original rear wall by more than 3.0m.",
                depth_ev,
                depth_action,
            )
        )
        if rear_boundary_distance is None:
            status = RuleStatus.NEEDS_CONFIRMATION
            evidence = "Distance to rear boundary not confirmed."
            action = "Confirm distance to rear boundary opposite rear wall."
        elif rear_boundary_distance >= 7.0:
            status = RuleStatus.PASS
            evidence = f"Distance to rear boundary: {rear_boundary_distance:.2f}m."
            action = ""
        else:
            status = RuleStatus.FAIL
            evidence = f"Distance to rear boundary: {rear_boundary_distance:.2f}m."
            action = "Full planning permission is likely required."
        checks.append(
            _make_check(
                "A-08",
                "More than single-storey rear boundary distance",
                status,
                "A rear enlargement with more than one storey must be at least 7.0m from the rear boundary opposite the rear wall.",
                evidence,
                action,
            )
        )

    # Materials
    materials = _get_answer(pd_answers, "materials_similar", "similar_materials")
    if _is_yes(materials):
        checks.append(
            _make_check(
                "A-09",
                "Materials similar appearance",
                RuleStatus.PASS,
                "Class A requires exterior materials to be of similar appearance to the existing dwellinghouse.",
                "Similar materials indicated.",
            )
        )
    elif _is_no(materials):
        checks.append(
            _make_check(
                "A-09",
                "Materials similar appearance",
                RuleStatus.FAIL,
                "Class A requires exterior materials to be of similar appearance to the existing dwellinghouse.",
                "Similar materials not confirmed / indicated no.",
                "Revise materials or use full planning route.",
            )
        )
    else:
        checks.append(
            _make_check(
                "A-09",
                "Materials similar appearance",
                RuleStatus.NEEDS_CONFIRMATION,
                "Class A requires exterior materials to be of similar appearance to the existing dwellinghouse.",
                "Materials not confirmed.",
                "Add materials note to drawings.",
            )
        )

    # 50% curtilage coverage
    coverage = _get_measure(measurements, "curtilage_coverage_percent", "coverage_percent")
    if coverage is None:
        checks.append(
            _make_check(
                "A-10",
                "50% curtilage coverage",
                RuleStatus.NEEDS_CONFIRMATION,
                "The total area of ground covered by buildings within the curtilage, excluding the original house, must not exceed 50% of the curtilage.",
                "Curtilage coverage not confirmed.",
                "Provide garden/curtilage coverage calculation if relying on PD.",
            )
        )
    elif coverage <= 50:
        checks.append(
            _make_check(
                "A-10",
                "50% curtilage coverage",
                RuleStatus.PASS,
                "The total area of ground covered by buildings within the curtilage, excluding the original house, must not exceed 50% of the curtilage.",
                f"Coverage: {coverage:.1f}%.",
            )
        )
    else:
        checks.append(
            _make_check(
                "A-10",
                "50% curtilage coverage",
                RuleStatus.FAIL,
                "The total area of ground covered by buildings within the curtilage, excluding the original house, must not exceed 50% of the curtilage.",
                f"Coverage: {coverage:.1f}%.",
                "Full planning permission is likely required.",
            )
        )

    return checks


# -----------------------------------------------------------------------------
# Class B - loft / dormer / roof enlargement
# -----------------------------------------------------------------------------


def check_class_b_roof_enlargement(
    project_types: List[str],
    property_type: str,
    measurements: Dict[str, Any],
    pd_answers: Dict[str, Any],
    detected_features: Optional[Dict[str, Any]] = None,
) -> List[RuleCheck]:
    checks: List[RuleCheck] = []
    text = _project_text(project_types, detected_features)
    relevant = _contains_any(text, ["loft", "dormer", "roof enlargement", "rear_dormer", "gable", "hip to gable"])
    if not relevant:
        return [
            _make_check("B-00", "Class B relevance", RuleStatus.NOT_APPLICABLE, "Class B applies to roof enlargements such as dormers.", "No Class B feature identified.")
        ]

    constraints = _get_answer(pd_answers, "site_constraints", "constraints")
    if any(x in constraints for x in ["conservation", "article 2", "aonb", "national park", "world heritage"]):
        checks.append(
            _make_check(
                "B-01",
                "Article 2(3) land",
                RuleStatus.FAIL,
                "Class B roof enlargements are not permitted development on Article 2(3) land.",
                f"Constraint input: {constraints}",
                "Use full planning route if confirmed.",
            )
        )
    elif constraints:
        checks.append(_make_check("B-01", "Article 2(3) land", RuleStatus.PASS, "Class B is restricted on Article 2(3) land.", f"Constraint input: {constraints}"))
    else:
        checks.append(_make_check("B-01", "Article 2(3) land", RuleStatus.NEEDS_CONFIRMATION, "Class B is restricted on Article 2(3) land.", "Constraint status not confirmed.", "Check conservation / Article 2(3) status."))

    front_roof = _get_answer(pd_answers, "front_roof_plane_highway", "front_roof_slope_highway")
    if _is_yes(front_roof):
        checks.append(_make_check("B-02", "Principal roof slope facing highway", RuleStatus.FAIL, "Class B does not permit roof enlargement extending beyond the plane of the principal roof slope facing a highway.", "Front/highway roof projection indicated.", "Full planning permission is likely required."))
    elif _is_no(front_roof):
        checks.append(_make_check("B-02", "Principal roof slope facing highway", RuleStatus.PASS, "Class B does not permit roof enlargement extending beyond the plane of the principal roof slope facing a highway.", "No front/highway roof projection indicated."))
    else:
        checks.append(_make_check("B-02", "Principal roof slope facing highway", RuleStatus.NEEDS_CONFIRMATION, "Class B does not permit roof enlargement extending beyond the plane of the principal roof slope facing a highway.", "Not confirmed.", "Confirm dormer/roof enlargement is to rear/acceptable roof plane only."))

    above_roof = _get_answer(pd_answers, "above_existing_roof_height", "above_highest_roof")
    if _is_yes(above_roof):
        checks.append(_make_check("B-03", "Highest roof height", RuleStatus.FAIL, "Class B roof enlargement must not exceed the highest part of the existing roof.", "Input indicates it exceeds the highest roof.", "Full planning permission is likely required."))
    elif _is_no(above_roof):
        checks.append(_make_check("B-03", "Highest roof height", RuleStatus.PASS, "Class B roof enlargement must not exceed the highest part of the existing roof.", "Input indicates it stays below/within the highest roof."))
    else:
        checks.append(_make_check("B-03", "Highest roof height", RuleStatus.NEEDS_CONFIRMATION, "Class B roof enlargement must not exceed the highest part of the existing roof.", "Not confirmed.", "Confirm ridge/highest roof relationship on elevations/sections."))

    roof_volume = _get_measure(measurements, "roof_volume_added_m3", "added_roof_volume_m3")
    category = _property_category(property_type)
    volume_limit = 40.0 if category == "terrace" else 50.0
    if roof_volume is None:
        checks.append(_make_check("B-04", "Roof volume allowance", RuleStatus.NEEDS_CONFIRMATION, f"Class B roof volume allowance is 40m³ for terrace houses and 50m³ for other houses. Applicable limit appears to be {volume_limit:.0f}m³.", "Added roof volume not confirmed.", "Provide roof volume calculation."))
    elif roof_volume <= volume_limit:
        checks.append(_make_check("B-04", "Roof volume allowance", RuleStatus.PASS, f"Class B roof volume allowance appears to be {volume_limit:.0f}m³ for this property type.", f"Added roof volume: {roof_volume:.1f}m³."))
    else:
        checks.append(_make_check("B-04", "Roof volume allowance", RuleStatus.FAIL, f"Class B roof volume allowance appears to be {volume_limit:.0f}m³ for this property type.", f"Added roof volume: {roof_volume:.1f}m³.", "Full planning permission is likely required."))

    eaves_setback = _get_answer(pd_answers, "eaves_setback_0_2m", "eaves_setback_200mm")
    if _is_yes(eaves_setback):
        checks.append(_make_check("B-05", "200mm eaves setback", RuleStatus.PASS, "The roof enlargement should normally be set back at least 0.2m from the original eaves, measured along the roof slope.", "200mm setback indicated."))
    elif _is_no(eaves_setback):
        checks.append(_make_check("B-05", "200mm eaves setback", RuleStatus.FAIL, "The roof enlargement should normally be set back at least 0.2m from the original eaves, unless an exception applies.", "Setback not provided.", "Revise dormer set-back or justify exception."))
    else:
        checks.append(_make_check("B-05", "200mm eaves setback", RuleStatus.NEEDS_CONFIRMATION, "The roof enlargement should normally be set back at least 0.2m from the original eaves.", "Not confirmed.", "Confirm on elevations/sections."))

    return checks


# -----------------------------------------------------------------------------
# Class C - rooflights and non-enlarging roof alterations
# -----------------------------------------------------------------------------


def check_class_c_rooflights(project_types: List[str], measurements: Dict[str, Any], pd_answers: Dict[str, Any], detected_features: Optional[Dict[str, Any]] = None) -> List[RuleCheck]:
    text = _project_text(project_types, detected_features)
    relevant = _contains_any(text, ["rooflight", "roof light", "rooflights", "front_rooflights", "class c"])
    if not relevant:
        return [_make_check("C-00", "Class C relevance", RuleStatus.NOT_APPLICABLE, "Class C applies to rooflights and other non-enlarging roof alterations.", "No Class C feature identified.")]
    checks: List[RuleCheck] = []
    protrusion = _get_measure(measurements, "rooflight_projection_m", "rooflight_protrusion_m")
    if protrusion is None:
        checks.append(_make_check("C-01", "Rooflight projection", RuleStatus.NEEDS_CONFIRMATION, "Class C roof alterations must not protrude more than 0.15m beyond the plane of the original roof slope.", "Projection not confirmed.", "Confirm rooflight projection detail."))
    elif protrusion <= 0.15:
        checks.append(_make_check("C-01", "Rooflight projection", RuleStatus.PASS, "Class C roof alterations must not protrude more than 0.15m beyond the plane of the original roof slope.", f"Projection: {protrusion:.2f}m."))
    else:
        checks.append(_make_check("C-01", "Rooflight projection", RuleStatus.FAIL, "Class C roof alterations must not protrude more than 0.15m beyond the plane of the original roof slope.", f"Projection: {protrusion:.2f}m.", "Full planning permission may be required or revise specification."))

    above_roof = _get_answer(pd_answers, "rooflight_above_highest_roof", "above_highest_roof")
    if _is_yes(above_roof):
        checks.append(_make_check("C-02", "Highest roof height", RuleStatus.FAIL, "Class C alteration must not be higher than the highest part of the original roof.", "Input indicates rooflight/alteration above highest roof.", "Revise or seek planning permission."))
    elif _is_no(above_roof):
        checks.append(_make_check("C-02", "Highest roof height", RuleStatus.PASS, "Class C alteration must not be higher than the highest part of the original roof.", "Input indicates not above highest roof."))
    else:
        checks.append(_make_check("C-02", "Highest roof height", RuleStatus.NEEDS_CONFIRMATION, "Class C alteration must not be higher than the highest part of the original roof.", "Not confirmed.", "Confirm rooflight/alteration height."))
    return checks


# -----------------------------------------------------------------------------
# Class D - porch
# -----------------------------------------------------------------------------


def check_class_d_porch(project_types: List[str], measurements: Dict[str, Any], pd_answers: Dict[str, Any], detected_features: Optional[Dict[str, Any]] = None) -> List[RuleCheck]:
    text = _project_text(project_types, detected_features)
    if not _contains_any(text, ["porch"]):
        return [_make_check("D-00", "Class D relevance", RuleStatus.NOT_APPLICABLE, "Class D applies to porches outside an external door.", "No porch feature identified.")]
    checks: List[RuleCheck] = []
    area = _get_measure(measurements, "porch_area_m2", "porch_ground_area_m2")
    height = _get_measure(measurements, "porch_height_m", "overall_height_m")
    highway_boundary_distance = _get_measure(measurements, "porch_highway_boundary_distance_m", "distance_to_highway_boundary_m")
    if area is None:
        checks.append(_make_check("D-01", "Porch ground area", RuleStatus.NEEDS_CONFIRMATION, "Class D porch ground area must not exceed 3m² measured externally.", "Porch area not confirmed.", "Confirm porch area."))
    elif area <= 3.0:
        checks.append(_make_check("D-01", "Porch ground area", RuleStatus.PASS, "Class D porch ground area must not exceed 3m² measured externally.", f"Area: {area:.2f}m²."))
    else:
        checks.append(_make_check("D-01", "Porch ground area", RuleStatus.FAIL, "Class D porch ground area must not exceed 3m² measured externally.", f"Area: {area:.2f}m².", "Full planning permission is likely required."))

    if height is None:
        checks.append(_make_check("D-02", "Porch height", RuleStatus.NEEDS_CONFIRMATION, "Class D porch height must not exceed 3.0m above ground level.", "Height not confirmed.", "Confirm porch height."))
    elif height <= 3.0:
        checks.append(_make_check("D-02", "Porch height", RuleStatus.PASS, "Class D porch height must not exceed 3.0m above ground level.", f"Height: {height:.2f}m."))
    else:
        checks.append(_make_check("D-02", "Porch height", RuleStatus.FAIL, "Class D porch height must not exceed 3.0m above ground level.", f"Height: {height:.2f}m.", "Full planning permission is likely required."))

    if highway_boundary_distance is None:
        checks.append(_make_check("D-03", "Porch distance to highway boundary", RuleStatus.NEEDS_CONFIRMATION, "Class D porch must not be within 2m of any boundary with a highway.", "Distance not confirmed.", "Confirm distance to highway boundary."))
    elif highway_boundary_distance >= 2.0:
        checks.append(_make_check("D-03", "Porch distance to highway boundary", RuleStatus.PASS, "Class D porch must not be within 2m of any boundary with a highway.", f"Distance: {highway_boundary_distance:.2f}m."))
    else:
        checks.append(_make_check("D-03", "Porch distance to highway boundary", RuleStatus.FAIL, "Class D porch must not be within 2m of any boundary with a highway.", f"Distance: {highway_boundary_distance:.2f}m.", "Full planning permission is likely required."))
    return checks


# -----------------------------------------------------------------------------
# Main public API
# -----------------------------------------------------------------------------


def run_planning_rule_checks(
    project_types: Optional[List[str]] = None,
    property_type: str = "",
    measurements: Optional[Dict[str, Any]] = None,
    pd_answers: Optional[Dict[str, Any]] = None,
    detected_features: Optional[Dict[str, Any]] = None,
) -> RuleEngineResult:
    """
    Run deterministic planning rule checks.

    Parameters:
        project_types: selected user project types, e.g. ["Ground Floor Rear Extension"]
        property_type: Detached / Semi-Detached / Terraced / Flat / Maisonette etc.
        measurements: values extracted from drawings or user inputs.
            Suggested keys:
                rear_depth_m
                overall_height_m
                eaves_height_m
                boundary_eaves_height_m
                roof_volume_added_m3
                rooflight_projection_m
                porch_area_m2
                porch_height_m
                porch_highway_boundary_distance_m
                curtilage_coverage_percent
                side_width_ratio
        pd_answers: user questionnaire answers and known constraints.
            Suggested keys:
                site_constraints
                within_2m_boundary / within_2m_of_boundary
                forward_of_principal_elevation
                materials_similar
                side_extension_width
                front_roof_plane_highway
                above_existing_roof_height
                eaves_setback_0_2m
                created_by_pd_change_of_use
        detected_features: boolean feature flags extracted by AI/PDF parser.
    """
    project_types = project_types or []
    measurements = measurements or {}
    pd_answers = pd_answers or {}
    detected_features = detected_features or {}

    checks: List[RuleCheck] = []
    checks.extend(check_general_pd_eligibility(property_type, pd_answers))
    checks.extend(check_class_a_extension(project_types, property_type, measurements, pd_answers, detected_features))
    checks.extend(check_class_b_roof_enlargement(project_types, property_type, measurements, pd_answers, detected_features))
    checks.extend(check_class_c_rooflights(project_types, measurements, pd_answers, detected_features))
    checks.extend(check_class_d_porch(project_types, measurements, pd_answers, detected_features))

    relevant_checks = [c for c in checks if c.status != RuleStatus.NOT_APPLICABLE]
    failed = [c for c in relevant_checks if c.status == RuleStatus.FAIL]
    needs = [c for c in relevant_checks if c.status == RuleStatus.NEEDS_CONFIRMATION]
    passed = [c for c in relevant_checks if c.status == RuleStatus.PASS]

    text = _project_text(project_types, detected_features)
    rear_relevant = _contains_any(text, ["rear extension", "single-storey rear extension", "single_storey_rear_extension"])
    class_a_relevant = any(c.code.startswith("A-") and c.status != RuleStatus.NOT_APPLICABLE for c in checks)

    if failed:
        likely_route = RouteStatus.FULL_PLANNING_LIKELY
        confidence = "HIGH" if len(failed) >= 2 else "MEDIUM"
        summary = "One or more deterministic rule checks fail. Full planning is likely unless the failed item is corrected or shown differently on the drawings."
    elif rear_relevant and any(c.code == "A-04" and "prior approval" in c.title.lower() and c.status == RuleStatus.PASS for c in checks):
        likely_route = RouteStatus.PRIOR_APPROVAL_POSSIBLE
        confidence = "MEDIUM" if needs else "HIGH"
        summary = "The rear extension appears to fall within the larger home extension prior approval range, subject to confirmation of all Class A conditions and neighbour consultation."
    elif class_a_relevant and not failed:
        likely_route = RouteStatus.PD_POSSIBLE
        confidence = "MEDIUM" if needs else "HIGH"
        summary = "No deterministic Class A failure has been identified. PD/LDC may be possible, subject to confirming unresolved items."
    elif any(c.code.startswith("B-") and c.status == RuleStatus.PASS for c in checks) and not failed:
        likely_route = RouteStatus.PD_POSSIBLE
        confidence = "MEDIUM" if needs else "HIGH"
        summary = "No deterministic Class B failure has been identified. PD/LDC may be possible, subject to confirming unresolved items."
    elif any(c.code.startswith("D-") and c.status == RuleStatus.PASS for c in checks) and not failed:
        likely_route = RouteStatus.PD_POSSIBLE
        confidence = "MEDIUM" if needs else "HIGH"
        summary = "No deterministic Class D failure has been identified. PD/LDC may be possible, subject to confirming unresolved items."
    else:
        likely_route = RouteStatus.NEEDS_CONFIRMATION
        confidence = "LOW" if needs else "MEDIUM"
        summary = "The selected works do not clearly match a rule set or key information is missing. Further review is required."

    return RuleEngineResult(
        likely_route=likely_route,
        confidence=confidence,
        summary=summary,
        checks=checks,
        failed_checks=failed,
        needs_confirmation=needs,
        passed_checks=passed,
    )


def format_rule_checks_for_prompt(result: RuleEngineResult, max_checks: int = 25) -> str:
    """Format deterministic rule checks so pdf_summary.py can pass them into the AI prompt."""
    lines = [
        "DETERMINISTIC PLANNING RULE ENGINE RESULT",
        f"Likely route: {result.likely_route.value}",
        f"Confidence: {result.confidence}",
        f"Summary: {result.summary}",
        "",
        "Rule checks:",
    ]
    for check in result.checks[:max_checks]:
        if check.status == RuleStatus.NOT_APPLICABLE:
            continue
        lines.append(f"- {check.code} | {check.title} | {check.status.value}")
        lines.append(f"  Rule: {check.rule}")
        if check.evidence:
            lines.append(f"  Evidence: {check.evidence}")
        if check.action:
            lines.append(f"  Action: {check.action}")
    return "\n".join(lines)


def format_rule_checks_for_report(result: RuleEngineResult) -> str:
    """Shorter plain text version for report sections."""
    lines = [
        f"Likely Route: {result.likely_route.value}",
        f"Rule Confidence: {result.confidence}",
        f"Summary: {result.summary}",
    ]
    if result.failed_checks:
        lines.append("Failed checks:")
        for c in result.failed_checks:
            lines.append(f"- {c.title}: {c.evidence or c.rule}")
    if result.needs_confirmation:
        lines.append("Needs confirmation:")
        for c in result.needs_confirmation[:8]:
            lines.append(f"- {c.title}: {c.action or c.rule}")
    return "\n".join(lines)


if __name__ == "__main__":
    # Quick manual test
    result = run_planning_rule_checks(
        project_types=["Ground Floor Rear Extension"],
        property_type="Semi-Detached House",
        measurements={"rear_depth_m": 6.0, "overall_height_m": 4.0, "eaves_height_m": 3.0},
        pd_answers={"within_2m_boundary": "yes", "materials_similar": "yes", "site_constraints": "None"},
        detected_features={"single_storey_rear_extension": True},
    )
    print(format_rule_checks_for_prompt(result))


# -----------------------------------------------------------------------------
# Compatibility API for current ArchLens app/pdf_summary integration
# -----------------------------------------------------------------------------

def facts_from_app_context(
    project_types=None,
    property_type: str = "",
    proposal_summary: str = "",
    pd_context: Optional[Dict[str, Any]] = None,
    scope_items=None,
):
    """Compatibility wrapper used by app.py/pdf_summary.py.

    Returns a conservative dictionary that is passed into run_householder_pd_rules.
    User inputs are used as rule filters only. Planning drawings and AI extraction
    should still be treated as the primary source where they conflict.
    """
    pd_context = pd_context or {}
    scope_items = list(scope_items or [])
    project_types = list(project_types or [])

    pd_answers: Dict[str, Any] = {}
    for key in [
        "site_constraints",
        "within_2m_boundary",
        "within_2m_of_boundary",
        "forward_of_principal_elevation",
        "materials_similar",
        "side_extension_width",
        "front_roof_plane_highway",
        "above_existing_roof_height",
        "above_highest_roof",
        "eaves_setback_0_2m",
        "eaves_setback_200mm",
        "created_by_pd_change_of_use",
        "side_windows_obscure_glazed",
        "porch_ground_area_band",
        "porch_height_band",
        "porch_within_2m_highway",
    ]:
        if key in pd_context:
            pd_answers[key] = pd_context.get(key)

    # Planning history / PD rights condition logic.
    pd_removed = str(pd_context.get("pd_rights_removed", "")).strip().lower()
    implemented = str(pd_context.get("previous_permission_implemented", "")).strip().lower()
    history_note = str(pd_context.get("planning_history_notes", "")).strip()
    if pd_removed == "yes" and implemented == "yes":
        # Treat as a likely PD-rights restriction.
        existing_constraints = str(pd_answers.get("site_constraints", "") or "")
        pd_answers["site_constraints"] = (existing_constraints + ", Article 4 direction / PD rights removed by condition").strip(" ,")
    elif pd_removed == "yes" and implemented == "no":
        # Do not hard fail. This is a planning-history nuance to be verified.
        pd_answers["planning_history_note"] = (
            "Previous permission may have contained a PD removal condition, but the user indicates it was not implemented. "
            "If not implemented, the PD-removal condition may not have taken effect; verify against the council register."
        )
    elif history_note:
        pd_answers["planning_history_note"] = history_note

    measurements: Dict[str, Any] = {}
    if pd_context.get("rear_extension_depth_m"):
        measurements["rear_depth_m"] = pd_context.get("rear_extension_depth_m")
    if pd_context.get("rear_extension_overall_height_m"):
        measurements["overall_height_m"] = pd_context.get("rear_extension_overall_height_m")

    # Convert band answers into conservative measurements where explicit dimensions are not available.
    roof_volume_band = str(pd_context.get("roof_volume_band", "")).lower()
    prop_l = str(property_type or "").lower()
    if "within" in roof_volume_band:
        measurements["roof_volume_added_m3"] = 40.0 if "terrace" in prop_l else 50.0
    elif "over" in roof_volume_band:
        measurements["roof_volume_added_m3"] = 41.0 if "terrace" in prop_l else 51.0

    porch_area = str(pd_context.get("porch_ground_area_band", "")).lower()
    if porch_area == "yes":
        measurements["porch_area_m2"] = 3.0
    elif porch_area == "no":
        measurements["porch_area_m2"] = 3.1

    porch_height = str(pd_context.get("porch_height_band", "")).lower()
    if porch_height == "yes":
        measurements["porch_height_m"] = 3.0
    elif porch_height == "no":
        measurements["porch_height_m"] = 3.1

    detected_features = {
        "loft_extension": any("loft" in p.lower() for p in project_types) or "loft" in proposal_summary.lower(),
        "rear_dormer": "dormer" in proposal_summary.lower(),
        "single_storey_rear_extension": any("ground floor rear" in p.lower() for p in project_types),
        "side_extension": any("side" in p.lower() or "infill" in p.lower() for p in project_types),
        "porch": any("porch" in p.lower() for p in project_types),
    }

    return {
        "project_types": project_types,
        "property_type": property_type,
        "measurements": measurements,
        "pd_answers": pd_answers,
        "detected_features": detected_features,
        "scope_items": scope_items,
    }


def run_householder_pd_rules(facts) -> RuleEngineResult:
    """Compatibility wrapper around run_planning_rule_checks."""
    if isinstance(facts, dict):
        return run_planning_rule_checks(
            project_types=facts.get("project_types") or [],
            property_type=facts.get("property_type") or "",
            measurements=facts.get("measurements") or {},
            pd_answers=facts.get("pd_answers") or {},
            detected_features=facts.get("detected_features") or {},
        )
    return run_planning_rule_checks()


def format_rule_result_for_prompt(result: RuleEngineResult) -> str:
    return format_rule_checks_for_prompt(result)
