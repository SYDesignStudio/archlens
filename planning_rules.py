"""
planning_rules.py

ArchLens AI deterministic planning rule engine.

Purpose
-------
This module checks householder permitted development rules before the AI writes
the narrative report. It is designed to reduce reliance on generative AI for
PASS / FAIL / NEEDS CONFIRMATION decisions.

Source basis
------------
Permitted development rights for householders: Technical Guidance, September 2019
and GPDO Schedule 2, Part 1 Classes A-H.

Important
---------
This module does not replace professional planning judgement. It performs
structured checks against known rule thresholds using supplied/extracted data.
Where data is missing, it returns NEEDS_CONFIRMATION rather than guessing.

Recommended workflow
--------------------
1. Extract facts from drawings/user intake into a ProjectFacts object.
2. Run run_householder_pd_rules(facts).
3. Pass the rule_result_summary into pdf_summary.py.
4. Let AI explain the deterministic result in simple wording, without changing
   PASS / FAIL outcomes unless a human updates the facts.
"""

from __future__ import annotations

from dataclasses import dataclass, field, asdict
from enum import Enum
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple


# -----------------------------------------------------------------------------
# Status model
# -----------------------------------------------------------------------------


class RuleStatus(str, Enum):
    PASS = "PASS"
    FAIL = "FAIL"
    NEEDS_CONFIRMATION = "NEEDS CONFIRMATION"
    NOT_APPLICABLE = "NOT APPLICABLE"


class RouteStatus(str, Enum):
    PD_POSSIBLE = "PD POSSIBLE"
    PRIOR_APPROVAL_POSSIBLE = "PRIOR APPROVAL POSSIBLE"
    FULL_PLANNING_REQUIRED = "FULL PLANNING REQUIRED"
    NEEDS_CONFIRMATION = "NEEDS CONFIRMATION"
    NOT_APPLICABLE = "NOT APPLICABLE"


@dataclass
class RuleCheck:
    class_ref: str
    rule_ref: str
    title: str
    status: RuleStatus
    reason: str
    required: str = ""
    actual: str = ""
    source: str = "Householder Technical Guidance / GPDO Part 1"
    severity: str = "MEDIUM"

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)


@dataclass
class ClassResult:
    class_ref: str
    title: str
    status: RouteStatus
    checks: List[RuleCheck] = field(default_factory=list)
    summary: str = ""

    def to_dict(self) -> Dict[str, Any]:
        return {
            "class_ref": self.class_ref,
            "title": self.title,
            "status": self.status.value,
            "summary": self.summary,
            "checks": [c.to_dict() for c in self.checks],
        }


@dataclass
class RuleEngineResult:
    overall_status: RouteStatus
    likely_route: str
    summary: str
    class_results: List[ClassResult] = field(default_factory=list)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "overall_status": self.overall_status.value,
            "likely_route": self.likely_route,
            "summary": self.summary,
            "class_results": [r.to_dict() for r in self.class_results],
        }


# -----------------------------------------------------------------------------
# Input facts
# -----------------------------------------------------------------------------
# Most fields are Optional because the app may not extract everything from the
# drawings. Unknown information should produce NEEDS CONFIRMATION, not a guess.
# Units:
# - metres for lengths/heights/depths
# - square metres for areas
# - cubic metres for roof volume
# - litres for container capacity
# -----------------------------------------------------------------------------


@dataclass
class ProjectFacts:
    # Basic property/context
    property_type: str = ""  # detached, semi-detached, terraced, end terrace, flat, maisonette, bungalow
    is_single_dwellinghouse: Optional[bool] = None
    created_by_part3_change_of_use_mnp_pa_q: Optional[bool] = None
    is_flat_or_maisonette: Optional[bool] = None
    is_listed_building_or_within_curtilage: Optional[bool] = None
    article_2_3_land: Optional[bool] = None
    article_4_direction: Optional[bool] = None
    pd_rights_removed_by_condition: Optional[bool] = None
    sssi: Optional[bool] = None
    conservation_area: Optional[bool] = None
    world_heritage_site: Optional[bool] = None
    national_park_broads_aonb: Optional[bool] = None

    # Proposal flags
    includes_class_a_extension_or_alteration: bool = False
    includes_single_storey_rear_extension: bool = False
    includes_larger_home_extension: bool = False
    includes_side_extension: bool = False
    includes_rear_and_side_extension: bool = False
    includes_more_than_one_storey_extension: bool = False
    includes_roof_enlargement_class_b: bool = False
    includes_roof_alteration_class_c: bool = False
    includes_porch_class_d: bool = False
    includes_outbuilding_pool_container_class_e: bool = False
    includes_hard_surface_class_f: bool = False
    includes_chimney_flue_svp_class_g: bool = False
    includes_antenna_class_h: bool = False

    # Class A dimensions/facts
    curtilage_coverage_percent_excluding_original_house: Optional[float] = None
    extension_height_m: Optional[float] = None
    extension_eaves_height_m: Optional[float] = None
    existing_house_highest_roof_height_m: Optional[float] = None
    existing_house_eaves_height_m: Optional[float] = None
    projects_beyond_principal_elevation: Optional[bool] = None
    projects_beyond_side_elevation_fronting_highway: Optional[bool] = None
    rear_projection_m: Optional[float] = None
    rear_projection_from_original_wall_m: Optional[float] = None
    total_rear_projection_joined_enlargement_m: Optional[float] = None
    within_2m_boundary: Optional[bool] = None
    side_extension_height_m: Optional[float] = None
    side_extension_width_m: Optional[float] = None
    original_house_width_m: Optional[float] = None
    side_extension_single_storey: Optional[bool] = None
    rear_boundary_distance_m: Optional[float] = None  # for multi-storey extension: boundary opposite rear wall
    includes_verandah_balcony_or_raised_platform: Optional[bool] = None
    includes_microwave_antenna: Optional[bool] = None
    includes_chimney_flue_or_svp: Optional[bool] = None
    alters_roof_under_class_a: Optional[bool] = None
    external_materials_similar: Optional[bool] = None
    upper_floor_side_windows_obscure_and_non_opening_1_7m: Optional[bool] = None
    roof_pitch_matches_original_for_multi_storey: Optional[bool] = None
    includes_external_cladding_article_2_3: Optional[bool] = None

    # Class B roof enlargement
    roof_enlargement_exceeds_highest_roof: Optional[bool] = None
    roof_enlargement_on_principal_roof_slope_fronting_highway: Optional[bool] = None
    added_roof_volume_m3: Optional[float] = None
    includes_roof_balcony_verandah_platform: Optional[bool] = None
    roof_materials_similar: Optional[bool] = None
    original_roof_eaves_maintained_or_reinstated: Optional[bool] = None
    roof_enlargement_eaves_setback_m: Optional[float] = None
    roof_enlargement_is_hip_to_gable: Optional[bool] = None
    roof_enlargement_joins_original_to_rear_or_side_extension_roof: Optional[bool] = None
    roof_enlargement_extends_beyond_outer_wall_face: Optional[bool] = None
    roof_side_windows_obscure_and_non_opening_1_7m: Optional[bool] = None

    # Class C roof alteration / rooflights
    rooflight_projection_m: Optional[float] = None
    roof_alteration_higher_than_original_roof: Optional[bool] = None
    class_c_includes_chimney_flue_svp: Optional[bool] = None
    class_c_includes_solar_equipment: Optional[bool] = None
    class_c_side_windows_obscure_and_non_opening_1_7m: Optional[bool] = None

    # Class D porch
    porch_ground_area_m2: Optional[float] = None
    porch_height_m: Optional[float] = None
    porch_distance_to_highway_boundary_m: Optional[float] = None

    # Class E outbuildings etc
    outbuilding_attached_to_house: Optional[bool] = None
    outbuilding_purpose_incidental: Optional[bool] = None
    outbuilding_forward_of_principal_elevation: Optional[bool] = None
    outbuilding_more_than_one_storey: Optional[bool] = None
    outbuilding_height_m: Optional[float] = None
    outbuilding_eaves_height_m: Optional[float] = None
    outbuilding_within_2m_boundary: Optional[bool] = None
    outbuilding_roof_type: str = ""  # dual-pitched, hipped, flat, mono-pitched, other
    outbuilding_within_curtilage_of_listed_building: Optional[bool] = None
    outbuilding_includes_verandah_balcony_raised_platform: Optional[bool] = None
    outbuilding_raised_platform_height_m: Optional[float] = None
    outbuilding_related_to_dwelling_or_antenna: Optional[bool] = None
    container_capacity_litres: Optional[float] = None
    outbuilding_more_than_20m_from_house: Optional[bool] = None
    outbuilding_area_more_than_20m_from_house_m2: Optional[float] = None
    outbuilding_between_side_wall_and_boundary_on_article_2_3_land: Optional[bool] = None

    # Class F hard surfaces
    hard_surface_forward_of_principal_elevation_and_highway: Optional[bool] = None
    hard_surface_area_m2: Optional[float] = None
    hard_surface_porous_or_drains_to_permeable_area: Optional[bool] = None

    # Class G chimney/flue/SVP
    chimney_flue_svp_height_above_highest_roof_m: Optional[float] = None
    chimney_flue_svp_on_principal_or_side_elevation_fronting_highway_article_2_3: Optional[bool] = None

    # Class H antenna
    number_of_antennas: Optional[int] = None
    antenna_lengths_m: Optional[List[float]] = None
    antenna_on_chimney: Optional[bool] = None
    antenna_installed_on_roof_without_chimney: Optional[bool] = None
    antenna_installed_on_roof_with_chimney: Optional[bool] = None
    antenna_protrudes_above_chimney: Optional[bool] = None
    antenna_cubic_capacity_litres: Optional[float] = None
    antenna_highest_part_higher_than_roof: Optional[bool] = None
    antenna_highest_part_higher_than_chimney_or_0_6m_above_ridge: Optional[bool] = None
    antenna_on_visible_highway_elevation_article_2_3: Optional[bool] = None
    building_height_m: Optional[float] = None
    antenna_sited_to_minimise_visual_impact: Optional[bool] = None


# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------


def _norm(value: str) -> str:
    return (value or "").strip().lower().replace("_", " ")


def is_detached(facts: ProjectFacts) -> bool:
    p = _norm(facts.property_type)
    return "detached" in p and "semi" not in p


def is_terrace(facts: ProjectFacts) -> bool:
    p = _norm(facts.property_type)
    return "terrace" in p or "terraced" in p


def is_non_detached_house(facts: ProjectFacts) -> bool:
    return not is_detached(facts)


def is_flat_or_maisonette(facts: ProjectFacts) -> bool:
    if facts.is_flat_or_maisonette is not None:
        return facts.is_flat_or_maisonette
    p = _norm(facts.property_type)
    return "flat" in p or "maisonette" in p


def _check_bool(
    class_ref: str,
    rule_ref: str,
    title: str,
    value: Optional[bool],
    pass_when: bool,
    fail_reason: str,
    pass_reason: str = "Requirement satisfied.",
    unknown_reason: str = "Not clearly confirmed from the submitted information.",
    required: str = "",
    severity: str = "MEDIUM",
) -> RuleCheck:
    if value is None:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.NEEDS_CONFIRMATION, unknown_reason, required=required, severity=severity)
    if value == pass_when:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.PASS, pass_reason, required=required, actual=str(value), severity="LOW")
    return RuleCheck(class_ref, rule_ref, title, RuleStatus.FAIL, fail_reason, required=required, actual=str(value), severity=severity)


def _check_max(
    class_ref: str,
    rule_ref: str,
    title: str,
    value: Optional[float],
    max_value: float,
    unit: str = "m",
    unknown_reason: str = "Dimension not clearly confirmed from the submitted information.",
    fail_reason: Optional[str] = None,
    pass_reason: Optional[str] = None,
    severity: str = "HIGH",
) -> RuleCheck:
    required = f"≤ {max_value:g}{unit}"
    if value is None:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.NEEDS_CONFIRMATION, unknown_reason, required=required, severity="MEDIUM")
    actual = f"{value:g}{unit}"
    if value <= max_value:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.PASS, pass_reason or f"{actual} is within the {required} limit.", required=required, actual=actual, severity="LOW")
    return RuleCheck(class_ref, rule_ref, title, RuleStatus.FAIL, fail_reason or f"{actual} exceeds the {required} limit.", required=required, actual=actual, severity=severity)


def _check_min(
    class_ref: str,
    rule_ref: str,
    title: str,
    value: Optional[float],
    min_value: float,
    unit: str = "m",
    unknown_reason: str = "Dimension not clearly confirmed from the submitted information.",
    fail_reason: Optional[str] = None,
    pass_reason: Optional[str] = None,
    severity: str = "HIGH",
) -> RuleCheck:
    required = f"≥ {min_value:g}{unit}"
    if value is None:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.NEEDS_CONFIRMATION, unknown_reason, required=required, severity="MEDIUM")
    actual = f"{value:g}{unit}"
    if value >= min_value:
        return RuleCheck(class_ref, rule_ref, title, RuleStatus.PASS, pass_reason or f"{actual} meets the {required} minimum.", required=required, actual=actual, severity="LOW")
    return RuleCheck(class_ref, rule_ref, title, RuleStatus.FAIL, fail_reason or f"{actual} is below the {required} minimum.", required=required, actual=actual, severity=severity)


def _class_status_from_checks(checks: Sequence[RuleCheck], prior_approval_candidate: bool = False) -> RouteStatus:
    active = [c for c in checks if c.status != RuleStatus.NOT_APPLICABLE]
    if not active:
        return RouteStatus.NOT_APPLICABLE
    if any(c.status == RuleStatus.FAIL for c in active):
        return RouteStatus.FULL_PLANNING_REQUIRED
    if any(c.status == RuleStatus.NEEDS_CONFIRMATION for c in active):
        return RouteStatus.NEEDS_CONFIRMATION
    if prior_approval_candidate:
        return RouteStatus.PRIOR_APPROVAL_POSSIBLE
    return RouteStatus.PD_POSSIBLE


def _summary_for_status(status: RouteStatus) -> str:
    if status == RouteStatus.PD_POSSIBLE:
        return "The supplied facts pass the applicable permitted development checks for this Class."
    if status == RouteStatus.PRIOR_APPROVAL_POSSIBLE:
        return "The supplied facts indicate the larger home extension prior approval route may be available, subject to neighbour consultation and all Class A conditions."
    if status == RouteStatus.FULL_PLANNING_REQUIRED:
        return "One or more mandatory permitted development limits appear to fail. Planning permission is likely required unless the facts are corrected."
    if status == RouteStatus.NEEDS_CONFIRMATION:
        return "The route cannot be confirmed because one or more key facts or dimensions are missing."
    return "This Class does not appear to be triggered by the supplied facts."


# -----------------------------------------------------------------------------
# General exclusions
# -----------------------------------------------------------------------------


def check_general_pd_exclusions(facts: ProjectFacts) -> List[RuleCheck]:
    checks = [
        _check_bool(
            "GENERAL",
            "Flats / maisonettes",
            "Householder PD rights",
            not is_flat_or_maisonette(facts),
            True,
            "Flats and maisonettes do not normally benefit from standard householder permitted development rights.",
            "The property is not identified as a flat or maisonette.",
            required="Single dwellinghouse, not a flat/maisonette",
            severity="HIGH",
        ),
        _check_bool(
            "GENERAL",
            "Single dwellinghouse",
            "Property must be a single dwellinghouse",
            facts.is_single_dwellinghouse,
            True,
            "The property is not confirmed as a single dwellinghouse.",
            "The property is confirmed as a single dwellinghouse.",
            required="Single dwellinghouse",
            severity="HIGH",
        ),
        _check_bool(
            "GENERAL",
            "Part 3 change of use exclusion",
            "House not created by Classes M/N/P/PA/Q",
            facts.created_by_part3_change_of_use_mnp_pa_q,
            False,
            "PD rights under Part 1 do not apply where the house was created only by specified Part 3 change of use rights.",
            "No Part 3 M/N/P/PA/Q creation issue identified.",
            required="Not created solely under Classes M/N/P/PA/Q",
            severity="HIGH",
        ),
        _check_bool(
            "GENERAL",
            "Article 4 / condition",
            "PD rights not removed",
            bool(facts.article_4_direction or facts.pd_rights_removed_by_condition),
            False,
            "PD rights may have been removed by Article 4 Direction or planning condition.",
            "No Article 4 Direction or PD removal condition identified from supplied facts.",
            required="No Article 4 / no PD removal condition",
            severity="HIGH",
        ),
    ]
    return checks


# -----------------------------------------------------------------------------
# Class A – enlargement, improvement or alteration
# -----------------------------------------------------------------------------


def check_class_a(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_class_a_extension_or_alteration:
        return ClassResult("Class A", "Enlargement, improvement or alteration", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []

    # A.1(a) Part 3 change of use exclusion
    checks.append(
        _check_bool(
            "Class A",
            "A.1(a)",
            "Not created by specified Part 3 change of use rights",
            facts.created_by_part3_change_of_use_mnp_pa_q,
            False,
            "Class A does not apply where the house was created only by specified Part 3 change of use rights.",
            "No Part 3 M/N/P/PA/Q creation issue identified.",
            required="False",
            severity="HIGH",
        )
    )

    # A.1(b) 50% curtilage coverage
    checks.append(
        _check_max(
            "Class A",
            "A.1(b)",
            "Total ground covered by buildings within curtilage",
            facts.curtilage_coverage_percent_excluding_original_house,
            50,
            "%",
            "Curtilage coverage calculation not confirmed.",
            "The 50% curtilage coverage limit appears to be exceeded.",
        )
    )

    # A.1(c) height not above highest roof
    if facts.extension_height_m is not None and facts.existing_house_highest_roof_height_m is not None:
        checks.append(
            _check_max(
                "Class A",
                "A.1(c)",
                "Height of enlarged part below highest existing roof",
                facts.extension_height_m,
                facts.existing_house_highest_roof_height_m,
                "m",
                fail_reason="The enlarged part would exceed the highest part of the existing roof.",
            )
        )
    else:
        checks.append(RuleCheck("Class A", "A.1(c)", "Height of enlarged part below highest existing roof", RuleStatus.NEEDS_CONFIRMATION, "Extension height and/or existing roof height not confirmed.", required="Extension height ≤ highest existing roof"))

    # A.1(d) eaves not above existing eaves
    if facts.extension_eaves_height_m is not None and facts.existing_house_eaves_height_m is not None:
        checks.append(
            _check_max(
                "Class A",
                "A.1(d)",
                "Eaves height not above existing eaves",
                facts.extension_eaves_height_m,
                facts.existing_house_eaves_height_m,
                "m",
                fail_reason="The extension eaves would exceed the eaves height of the existing house.",
            )
        )
    else:
        checks.append(RuleCheck("Class A", "A.1(d)", "Eaves height not above existing eaves", RuleStatus.NEEDS_CONFIRMATION, "Extension eaves height and/or existing eaves height not confirmed.", required="Extension eaves ≤ existing eaves"))

    # A.1(e)
    checks.append(
        _check_bool(
            "Class A",
            "A.1(e)(i)",
            "Does not project beyond principal elevation",
            facts.projects_beyond_principal_elevation,
            False,
            "Development forward of the principal elevation is not permitted by Class A.",
            "No projection beyond the principal elevation identified.",
            required="False",
            severity="HIGH",
        )
    )
    checks.append(
        _check_bool(
            "Class A",
            "A.1(e)(ii)",
            "Does not project beyond side elevation fronting highway",
            facts.projects_beyond_side_elevation_fronting_highway,
            False,
            "Development forward of a side elevation fronting a highway is not permitted by Class A.",
            "No projection beyond a side elevation fronting a highway identified.",
            required="False",
            severity="HIGH",
        )
    )

    # A.1(f)/(g) single-storey rear extension and larger home extension
    prior_approval_candidate = False
    if facts.includes_single_storey_rear_extension:
        depth = facts.total_rear_projection_joined_enlargement_m or facts.rear_projection_from_original_wall_m or facts.rear_projection_m
        standard_limit = 4.0 if is_detached(facts) else 3.0
        larger_limit = 8.0 if is_detached(facts) else 6.0
        checks.append(
            _check_max(
                "Class A",
                "A.1(f)(ii) / A.1(g)(ii)",
                "Single-storey rear extension overall height",
                facts.extension_height_m,
                4.0,
                "m",
                fail_reason="Single-storey rear extension exceeds 4.0m overall height.",
            )
        )

        if depth is None:
            checks.append(RuleCheck("Class A", "A.1(f)/(g)", "Rear projection from original rear wall", RuleStatus.NEEDS_CONFIRMATION, "Rear projection from the original rear wall is not confirmed.", required=f"≤ {standard_limit:g}m standard PD or ≤ {larger_limit:g}m prior approval"))
        elif depth <= standard_limit:
            checks.append(RuleCheck("Class A", "A.1(f)", "Rear projection from original rear wall", RuleStatus.PASS, f"{depth:g}m is within standard Class A rear extension depth.", required=f"≤ {standard_limit:g}m", actual=f"{depth:g}m", severity="LOW"))
        elif depth <= larger_limit:
            prior_approval_candidate = True
            if facts.article_2_3_land or facts.sssi:
                checks.append(RuleCheck("Class A", "A.1(g)", "Larger home extension eligibility", RuleStatus.FAIL, "The larger home extension prior approval route is not available on Article 2(3) land or SSSI land.", required="Not Article 2(3) land / not SSSI", actual="Constrained land", severity="HIGH"))
            elif facts.article_2_3_land is None or facts.sssi is None:
                checks.append(RuleCheck("Class A", "A.1(g)", "Larger home extension eligibility", RuleStatus.NEEDS_CONFIRMATION, "Article 2(3) land / SSSI status must be confirmed for larger home extension prior approval.", required="Not Article 2(3) land / not SSSI", actual=f"{depth:g}m"))
            else:
                checks.append(RuleCheck("Class A", "A.1(g)", "Rear projection within larger home extension limit", RuleStatus.PASS, f"{depth:g}m is within the larger home extension prior approval depth limit.", required=f"> {standard_limit:g}m and ≤ {larger_limit:g}m", actual=f"{depth:g}m", severity="LOW"))
        else:
            checks.append(RuleCheck("Class A", "A.1(f)/(g)", "Rear projection from original rear wall", RuleStatus.FAIL, f"{depth:g}m exceeds the maximum larger home extension depth limit.", required=f"≤ {larger_limit:g}m", actual=f"{depth:g}m", severity="HIGH"))

    # A.1(h) multi-storey rear
    if facts.includes_more_than_one_storey_extension:
        depth = facts.total_rear_projection_joined_enlargement_m or facts.rear_projection_from_original_wall_m or facts.rear_projection_m
        checks.append(_check_max("Class A", "A.1(h)(i)", "Multi-storey rear projection", depth, 3.0, "m", fail_reason="More-than-one-storey enlargement projects beyond rear wall by more than 3m."))
        checks.append(_check_min("Class A", "A.1(h)(ii)", "Distance to rear boundary opposite rear wall", facts.rear_boundary_distance_m, 7.0, "m", fail_reason="More-than-one-storey enlargement is within 7m of the rear boundary opposite the rear wall."))

    # A.1(i) eaves within 2m boundary
    if facts.within_2m_boundary is True:
        checks.append(_check_max("Class A", "A.1(i)", "Eaves height within 2m of boundary", facts.extension_eaves_height_m, 3.0, "m", fail_reason="Eaves exceed 3.0m within 2m of a boundary."))
    elif facts.within_2m_boundary is None:
        checks.append(RuleCheck("Class A", "A.1(i)", "Eaves height within 2m of boundary", RuleStatus.NEEDS_CONFIRMATION, "Whether the extension is within 2m of a boundary is not confirmed.", required="If within 2m, eaves ≤ 3.0m"))
    else:
        checks.append(RuleCheck("Class A", "A.1(i)", "Eaves height within 2m of boundary", RuleStatus.NOT_APPLICABLE, "Extension is not within 2m of a boundary.", required="If within 2m, eaves ≤ 3.0m", severity="LOW"))

    # A.1(j) side extension
    if facts.includes_side_extension or facts.includes_rear_and_side_extension:
        checks.append(_check_max("Class A", "A.1(j)(i)", "Side extension height", facts.side_extension_height_m or facts.extension_height_m, 4.0, "m", fail_reason="Side extension exceeds 4.0m height."))
        checks.append(
            _check_bool(
                "Class A",
                "A.1(j)(ii)",
                "Side extension single storey only",
                facts.side_extension_single_storey,
                True,
                "Side extension has more than one storey.",
                "Side extension is single storey.",
                required="Single storey",
                severity="HIGH",
            )
        )
        if facts.side_extension_width_m is not None and facts.original_house_width_m:
            half_width = facts.original_house_width_m / 2
            checks.append(_check_max("Class A", "A.1(j)(iii)", "Side extension width", facts.side_extension_width_m, half_width, "m", fail_reason="Side extension width is greater than half the width of the original house."))
        else:
            checks.append(RuleCheck("Class A", "A.1(j)(iii)", "Side extension width", RuleStatus.NEEDS_CONFIRMATION, "Side extension width and/or original house width not confirmed.", required="≤ half width of original house"))

    # A.1(ja) total enlargement
    if facts.total_rear_projection_joined_enlargement_m is not None:
        checks.append(RuleCheck("Class A", "A.1(ja)", "Total enlargement joined to existing enlargement", RuleStatus.PASS, "Total enlargement value supplied and included in relevant checks.", actual=f"{facts.total_rear_projection_joined_enlargement_m:g}m", severity="LOW"))
    else:
        checks.append(RuleCheck("Class A", "A.1(ja)", "Total enlargement joined to existing enlargement", RuleStatus.NEEDS_CONFIRMATION, "Confirm whether the proposal joins any existing enlargement; if so, total enlargement limits apply.", required="Total enlargement must meet A.1(e)-(j)"))

    # A.1(k)
    checks.append(_check_bool("Class A", "A.1(k)(i)", "No verandah, balcony or raised platform", facts.includes_verandah_balcony_or_raised_platform, False, "Class A does not permit verandahs, balconies or raised platforms.", "No verandah, balcony or raised platform identified.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class A", "A.1(k)(ii)", "No microwave antenna under Class A", facts.includes_microwave_antenna, False, "Microwave antennas are not permitted under Class A; consider Class H.", "No microwave antenna included under Class A.", required="False"))
    checks.append(_check_bool("Class A", "A.1(k)(iii)", "No chimney/flue/SVP under Class A", facts.includes_chimney_flue_or_svp, False, "Chimneys, flues and soil vent pipes are not permitted under Class A; consider Class G.", "No chimney/flue/SVP included under Class A.", required="False"))
    checks.append(_check_bool("Class A", "A.1(k)(iv)", "No alteration to roof under Class A", facts.alters_roof_under_class_a, False, "Roof alterations are not permitted under Class A; consider Class B or Class C.", "No Class A roof alteration identified.", required="False"))

    # A.2 Article 2(3) land additional restrictions
    if facts.article_2_3_land is True:
        checks.append(_check_bool("Class A", "A.2(a)", "No exterior cladding on Article 2(3) land", facts.includes_external_cladding_article_2_3, False, "Exterior cladding is not permitted development on Article 2(3) land.", "No exterior cladding identified.", required="False", severity="HIGH"))
        checks.append(_check_bool("Class A", "A.2(b)", "No side extension on Article 2(3) land", facts.includes_side_extension or facts.includes_rear_and_side_extension, False, "Extensions beyond a side wall are not permitted development on Article 2(3) land.", "No side extension on Article 2(3) land.", required="False", severity="HIGH"))
        checks.append(_check_bool("Class A", "A.2(c)", "No multi-storey rear extension on Article 2(3) land", facts.includes_more_than_one_storey_extension, False, "Multi-storey rear extensions are not permitted development on Article 2(3) land.", "No multi-storey rear extension on Article 2(3) land.", required="False", severity="HIGH"))

    # A.3 conditions
    checks.append(_check_bool("Class A", "A.3(a)", "External materials similar", facts.external_materials_similar, True, "External materials are not confirmed as similar in appearance.", "External materials are confirmed as similar.", required="Similar appearance"))
    if facts.includes_more_than_one_storey_extension:
        checks.append(_check_bool("Class A", "A.3(b)", "Upper-floor side windows obscure/non-opening", facts.upper_floor_side_windows_obscure_and_non_opening_1_7m, True, "Upper-floor side windows must be obscure-glazed and non-opening below 1.7m.", "Upper-floor side windows comply.", required="Obscure glazed and non-opening below 1.7m"))
        checks.append(_check_bool("Class A", "A.3(c)", "Roof pitch matches original", facts.roof_pitch_matches_original_for_multi_storey, True, "Roof pitch of multi-storey enlargement must match original so far as practicable.", "Roof pitch condition satisfied.", required="Match original so far as practicable"))

    status = _class_status_from_checks(checks, prior_approval_candidate=prior_approval_candidate)
    return ClassResult("Class A", "Enlargement, improvement or alteration", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class B – roof enlargement
# -----------------------------------------------------------------------------


def check_class_b(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_roof_enlargement_class_b:
        return ClassResult("Class B", "Additions etc. to the roof", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class B", "B.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class B does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class B", "B.1(b)", "Does not exceed highest existing roof", facts.roof_enlargement_exceeds_highest_roof, False, "Roof enlargement exceeds the highest part of the existing roof.", "Roof enlargement does not exceed the highest existing roof.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class B", "B.1(c)", "No enlargement beyond principal roof slope fronting highway", facts.roof_enlargement_on_principal_roof_slope_fronting_highway, False, "Roof enlargement on the principal roof slope fronting a highway is not permitted under Class B.", "No principal highway-facing roof enlargement identified.", required="False", severity="HIGH"))

    volume_limit = 40.0 if is_terrace(facts) else 50.0
    checks.append(_check_max("Class B", "B.1(d)", "Additional roof volume", facts.added_roof_volume_m3, volume_limit, "m³", fail_reason=f"Additional roof volume exceeds the {volume_limit:g}m³ Class B allowance."))

    checks.append(_check_bool("Class B", "B.1(e)(i)", "No roof balcony/verandah/raised platform", facts.includes_roof_balcony_verandah_platform, False, "Class B does not permit balconies, verandahs or raised platforms.", "No roof balcony/verandah/platform identified.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class B", "B.1(e)(ii)", "No chimney/flue/SVP under Class B", facts.includes_chimney_flue_or_svp, False, "Chimneys, flues and SVPs are not permitted under Class B; consider Class G.", "No chimney/flue/SVP included under Class B.", required="False"))
    checks.append(_check_bool("Class B", "B.1(f)", "Not on Article 2(3) land", facts.article_2_3_land, False, "Class B roof enlargements are not permitted development on Article 2(3) land.", "Not identified as Article 2(3) land.", required="False", severity="HIGH"))

    # Conditions B.2
    checks.append(_check_bool("Class B", "B.2(a)", "Roof materials similar", facts.roof_materials_similar, True, "Roof enlargement materials are not confirmed as similar in appearance.", "Roof materials are confirmed as similar.", required="Similar appearance"))

    setback_exempt = bool(facts.roof_enlargement_is_hip_to_gable or facts.roof_enlargement_joins_original_to_rear_or_side_extension_roof)
    if not setback_exempt:
        checks.append(_check_bool("Class B", "B.2(b)(i)(aa)", "Original eaves maintained or reinstated", facts.original_roof_eaves_maintained_or_reinstated, True, "Original roof eaves must be maintained or reinstated.", "Original eaves maintained/reinstated.", required="True"))
        checks.append(_check_min("Class B", "B.2(b)(i)(bb)", "Eaves setback", facts.roof_enlargement_eaves_setback_m, 0.2, "m", fail_reason="Roof enlargement is less than 0.2m from the eaves where the setback condition applies."))
        checks.append(_check_bool("Class B", "B.2(b)(ii)", "Does not extend beyond outside wall face", facts.roof_enlargement_extends_beyond_outer_wall_face, False, "Roof enlargement must not extend beyond the outside face of any external wall of the original house.", "No extension beyond outside wall face identified.", required="False", severity="HIGH"))
    else:
        checks.append(RuleCheck("Class B", "B.2(b)", "Hip-to-gable / roof-joining exception", RuleStatus.NOT_APPLICABLE, "The 0.2m eaves setback / wall-face condition may be exempt or modified for hip-to-gable or roof-joining enlargements.", severity="LOW"))

    checks.append(_check_bool("Class B", "B.2(c)", "Side windows obscure/non-opening", facts.roof_side_windows_obscure_and_non_opening_1_7m, True, "Side windows must be obscure-glazed and non-opening below 1.7m.", "Side windows comply or are not proposed.", required="Obscure glazed and non-opening below 1.7m"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class B", "Additions etc. to the roof", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class C – other roof alterations
# -----------------------------------------------------------------------------


def check_class_c(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_roof_alteration_class_c:
        return ClassResult("Class C", "Other alterations to the roof", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class C", "C.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class C does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_max("Class C", "C.1(b)", "Roof alteration projection", facts.rooflight_projection_m, 0.15, "m", fail_reason="Roof alteration/rooflight protrudes more than 0.15m beyond the roof plane."))
    checks.append(_check_bool("Class C", "C.1(c)", "Not higher than original roof", facts.roof_alteration_higher_than_original_roof, False, "Roof alteration is higher than the highest part of the original roof.", "Roof alteration does not exceed original roof height.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class C", "C.1(d)(i)", "No chimney/flue/SVP under Class C", facts.class_c_includes_chimney_flue_svp, False, "Chimney/flue/SVP works are not permitted under Class C; consider Class G.", "No chimney/flue/SVP under Class C.", required="False"))
    checks.append(_check_bool("Class C", "C.1(d)(ii)", "No solar equipment under Class C", facts.class_c_includes_solar_equipment, False, "Solar equipment is not permitted under Class C; consider Part 14.", "No solar equipment under Class C.", required="False"))
    checks.append(_check_bool("Class C", "C.2", "Side roof windows obscure/non-opening", facts.class_c_side_windows_obscure_and_non_opening_1_7m, True, "Side roof windows must be obscure-glazed and non-opening below 1.7m.", "Side roof windows comply or are not proposed.", required="Obscure glazed and non-opening below 1.7m"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class C", "Other alterations to the roof", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class D – porches
# -----------------------------------------------------------------------------


def check_class_d(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_porch_class_d:
        return ClassResult("Class D", "Porches", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class D", "D.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class D does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_max("Class D", "D.1(b)", "Porch ground area", facts.porch_ground_area_m2, 3.0, "m²", fail_reason="Porch ground area exceeds 3m²."))
    checks.append(_check_max("Class D", "D.1(c)", "Porch height", facts.porch_height_m, 3.0, "m", fail_reason="Porch height exceeds 3m."))
    if facts.porch_distance_to_highway_boundary_m is None:
        checks.append(RuleCheck("Class D", "D.1(d)", "Distance to highway boundary", RuleStatus.NEEDS_CONFIRMATION, "Distance from porch to highway boundary is not confirmed.", required="≥ 2m"))
    elif facts.porch_distance_to_highway_boundary_m < 2:
        checks.append(RuleCheck("Class D", "D.1(d)", "Distance to highway boundary", RuleStatus.FAIL, "Porch is within 2m of a boundary with a highway.", required="≥ 2m", actual=f"{facts.porch_distance_to_highway_boundary_m:g}m", severity="HIGH"))
    else:
        checks.append(RuleCheck("Class D", "D.1(d)", "Distance to highway boundary", RuleStatus.PASS, "Porch is at least 2m from a boundary with a highway.", required="≥ 2m", actual=f"{facts.porch_distance_to_highway_boundary_m:g}m", severity="LOW"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class D", "Porches", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class E – outbuildings, pools, enclosures, containers
# -----------------------------------------------------------------------------


def check_class_e(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_outbuilding_pool_container_class_e:
        return ClassResult("Class E", "Buildings etc. within the curtilage", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class E", "E.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class E does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_max("Class E", "E.1(b)", "Total curtilage building coverage", facts.curtilage_coverage_percent_excluding_original_house, 50.0, "%", fail_reason="Total ground covered by buildings/enclosures/containers exceeds 50% of curtilage excluding original house."))
    checks.append(_check_bool("Class E", "E.1(c)", "Not forward of principal elevation", facts.outbuilding_forward_of_principal_elevation, False, "Class E development forward of the principal elevation is not permitted development.", "Not forward of principal elevation.", required="False", severity="HIGH"))
    checks.append(_check_bool("Class E", "E.1(d)", "Single storey only", facts.outbuilding_more_than_one_storey, False, "Class E buildings must not have more than one storey.", "Building is single storey.", required="False", severity="HIGH"))

    # Height E.1(e)
    if facts.outbuilding_within_2m_boundary is True:
        checks.append(_check_max("Class E", "E.1(e)(ii)", "Height within 2m of boundary", facts.outbuilding_height_m, 2.5, "m", fail_reason="Class E building/container/enclosure within 2m of boundary exceeds 2.5m height."))
    elif _norm(facts.outbuilding_roof_type) in {"dual pitched", "dual-pitched", "hipped"}:
        checks.append(_check_max("Class E", "E.1(e)(i)", "Dual-pitched / hipped roof height", facts.outbuilding_height_m, 4.0, "m", fail_reason="Class E dual-pitched/hipped roof building exceeds 4m height."))
    else:
        checks.append(_check_max("Class E", "E.1(e)(iii)", "Other roof/container/enclosure height", facts.outbuilding_height_m, 3.0, "m", fail_reason="Class E building/container/enclosure exceeds 3m height."))

    checks.append(_check_max("Class E", "E.1(f)", "Eaves height", facts.outbuilding_eaves_height_m, 2.5, "m", fail_reason="Class E building eaves exceed 2.5m."))
    checks.append(_check_bool("Class E", "E.1(g)", "Not within curtilage of listed building", facts.outbuilding_within_curtilage_of_listed_building or facts.is_listed_building_or_within_curtilage, False, "Class E development within the curtilage of a listed building requires planning permission.", "Not within curtilage of a listed building.", required="False", severity="HIGH"))

    # Raised platforms / verandahs
    if facts.outbuilding_includes_verandah_balcony_raised_platform is True:
        checks.append(RuleCheck("Class E", "E.1(h)", "No verandah/balcony/raised platform", RuleStatus.FAIL, "Class E does not permit verandahs, balconies or raised platforms.", required="False", actual="True", severity="HIGH"))
    elif facts.outbuilding_raised_platform_height_m is not None and facts.outbuilding_raised_platform_height_m > 0.3:
        checks.append(RuleCheck("Class E", "E.1(h)", "Raised platform height", RuleStatus.FAIL, "Raised platform exceeds 0.3m and is not permitted under Class E.", required="≤ 0.3m", actual=f"{facts.outbuilding_raised_platform_height_m:g}m", severity="HIGH"))
    elif facts.outbuilding_includes_verandah_balcony_raised_platform is None and facts.outbuilding_raised_platform_height_m is None:
        checks.append(RuleCheck("Class E", "E.1(h)", "No verandah/balcony/raised platform", RuleStatus.NEEDS_CONFIRMATION, "Confirm whether the proposal includes a verandah, balcony or raised platform.", required="False"))
    else:
        checks.append(RuleCheck("Class E", "E.1(h)", "No verandah/balcony/raised platform", RuleStatus.PASS, "No prohibited raised platform/verandah/balcony identified.", required="False", severity="LOW"))

    checks.append(_check_bool("Class E", "E.1(i)", "Incidental purpose only", facts.outbuilding_related_to_dwelling_or_antenna, False, "Class E cannot be used for works related to the dwelling itself or microwave antenna.", "No dwelling-related works/antenna issue identified under Class E.", required="False"))
    checks.append(_check_bool("Class E", "E purpose", "Purpose incidental to enjoyment of dwellinghouse", facts.outbuilding_purpose_incidental, True, "The outbuilding/pool/container is not confirmed as incidental to the enjoyment of the dwellinghouse.", "Purpose is confirmed as incidental.", required="Incidental purpose"))

    checks.append(_check_max("Class E", "E.1(j)", "Container capacity", facts.container_capacity_litres, 3500.0, " litres", fail_reason="Container capacity exceeds 3,500 litres."))

    # E.2 protected land >20m from house
    if facts.national_park_broads_aonb or facts.world_heritage_site:
        if facts.outbuilding_more_than_20m_from_house is True:
            checks.append(_check_max("Class E", "E.2", "Area more than 20m from house on protected land", facts.outbuilding_area_more_than_20m_from_house_m2, 10.0, "m²", fail_reason="On protected land, Class E buildings etc more than 20m from the house exceed 10m²."))
        elif facts.outbuilding_more_than_20m_from_house is None:
            checks.append(RuleCheck("Class E", "E.2", "Area more than 20m from house on protected land", RuleStatus.NEEDS_CONFIRMATION, "Protected land status indicated; confirm whether development is more than 20m from the house.", required="If >20m, total area ≤ 10m²"))

    # E.3 article 2(3) side land
    if facts.article_2_3_land:
        checks.append(_check_bool("Class E", "E.3", "Not between side wall and boundary on Article 2(3) land", facts.outbuilding_between_side_wall_and_boundary_on_article_2_3_land, False, "Class E development between a side wall and boundary is not permitted on Article 2(3) land.", "No side-land Class E issue identified.", required="False", severity="HIGH"))

    # Attached buildings should be Class A not E
    checks.append(_check_bool("Class E", "Class boundary", "Not attached to the house", facts.outbuilding_attached_to_house, False, "Buildings attached to the house are not assessed under Class E; Class A applies.", "Building is not attached to the house.", required="False"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class E", "Buildings etc. within the curtilage", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class F – hard surfaces
# -----------------------------------------------------------------------------


def check_class_f(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_hard_surface_class_f:
        return ClassResult("Class F", "Hard surfaces", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class F", "F.1", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class F does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))

    if facts.hard_surface_forward_of_principal_elevation_and_highway is True:
        if facts.hard_surface_area_m2 is None:
            checks.append(RuleCheck("Class F", "F.2", "Front hard surface area", RuleStatus.NEEDS_CONFIRMATION, "Hard surface area forward of principal elevation and highway is not confirmed.", required="If >5m², porous/drains to permeable area"))
        elif facts.hard_surface_area_m2 <= 5:
            checks.append(RuleCheck("Class F", "F.2", "Front hard surface drainage", RuleStatus.PASS, "Hard surface area is 5m² or less, so the porous/permeable drainage condition is not triggered.", required="≤ 5m² or porous/permeable drainage", actual=f"{facts.hard_surface_area_m2:g}m²", severity="LOW"))
        else:
            checks.append(_check_bool("Class F", "F.2", "Porous or drains to permeable area", facts.hard_surface_porous_or_drains_to_permeable_area, True, "Front hard surface over 5m² must be porous or drain to a permeable area within the curtilage.", "Drainage condition satisfied.", required="True", severity="HIGH"))
    elif facts.hard_surface_forward_of_principal_elevation_and_highway is None:
        checks.append(RuleCheck("Class F", "F.2", "Hard surface location", RuleStatus.NEEDS_CONFIRMATION, "Confirm whether the hard surface lies between the principal elevation and highway.", required="If yes and >5m², porous/permeable drainage required"))
    else:
        checks.append(RuleCheck("Class F", "F.2", "Front hard surface drainage", RuleStatus.NOT_APPLICABLE, "Hard surface is not forward of the principal elevation and highway.", severity="LOW"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class F", "Hard surfaces", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class G – chimneys, flues, soil and vent pipes
# -----------------------------------------------------------------------------


def check_class_g(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_chimney_flue_svp_class_g:
        return ClassResult("Class G", "Chimneys, flues, soil and vent pipes", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class G", "G.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class G does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_max("Class G", "G.1(b)", "Height above highest roof", facts.chimney_flue_svp_height_above_highest_roof_m, 1.0, "m", fail_reason="Chimney/flue/SVP exceeds the highest part of the roof by 1m or more."))
    if facts.article_2_3_land:
        checks.append(_check_bool("Class G", "G.1(c)", "Not on principal/side highway elevation on Article 2(3) land", facts.chimney_flue_svp_on_principal_or_side_elevation_fronting_highway_article_2_3, False, "On Article 2(3) land, chimney/flue/SVP on a principal or side elevation fronting a highway is not permitted development.", "No Article 2(3) highway-facing elevation issue identified.", required="False", severity="HIGH"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class G", "Chimneys, flues, soil and vent pipes", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Class H – microwave antenna
# -----------------------------------------------------------------------------


def check_class_h(facts: ProjectFacts) -> ClassResult:
    if not facts.includes_antenna_class_h:
        return ClassResult("Class H", "Microwave antenna", RouteStatus.NOT_APPLICABLE, [], _summary_for_status(RouteStatus.NOT_APPLICABLE))

    checks: List[RuleCheck] = []
    checks.append(_check_bool("Class H", "H.1(a)", "Not created by specified Part 3 change of use rights", facts.created_by_part3_change_of_use_mnp_pa_q, False, "Class H does not apply where the house was created only by specified Part 3 change of use rights.", "No Part 3 M/N/P/PA/Q creation issue identified.", required="False", severity="HIGH"))
    checks.append(_check_max("Class H", "H.1(b)(i)", "Number of antennas", float(facts.number_of_antennas) if facts.number_of_antennas is not None else None, 2.0, "", fail_reason="More than 2 antennas would be present."))

    lengths = facts.antenna_lengths_m or []
    if not lengths:
        checks.append(RuleCheck("Class H", "H.1(b)(ii)-(iii)", "Antenna length criteria", RuleStatus.NEEDS_CONFIRMATION, "Antenna lengths are not confirmed.", required="Single antenna ≤1m; if two antennas, only one may exceed 0.6m and none may exceed 1m."))
    else:
        too_long = [x for x in lengths if x > 1.0]
        over_06 = [x for x in lengths if x > 0.6]
        if too_long:
            checks.append(RuleCheck("Class H", "H.1(b)(ii)", "Antenna maximum length", RuleStatus.FAIL, "At least one antenna exceeds 1m length.", required="≤ 1m", actual=", ".join(f"{x:g}m" for x in lengths), severity="HIGH"))
        elif len(over_06) > 1:
            checks.append(RuleCheck("Class H", "H.1(b)(iii)", "Two antenna length criteria", RuleStatus.FAIL, "More than one antenna exceeds 0.6m.", required="Only one antenna may exceed 0.6m", actual=", ".join(f"{x:g}m" for x in lengths), severity="HIGH"))
        else:
            checks.append(RuleCheck("Class H", "H.1(b)(ii)-(iii)", "Antenna length criteria", RuleStatus.PASS, "Antenna length criteria satisfied.", required="≤ 1m and only one >0.6m", actual=", ".join(f"{x:g}m" for x in lengths), severity="LOW"))

    if facts.antenna_on_chimney:
        chimney_lengths = lengths or None
        max_len = max(chimney_lengths) if chimney_lengths else None
        checks.append(_check_max("Class H", "H.1(b)(iv)", "Antenna on chimney length", max_len, 0.6, "m", fail_reason="Antenna installed on chimney exceeds 0.6m length."))
        checks.append(_check_bool("Class H", "H.1(b)(v)", "Antenna does not protrude above chimney", facts.antenna_protrudes_above_chimney, False, "Antenna installed on chimney protrudes above chimney.", "Antenna does not protrude above chimney.", required="False", severity="HIGH"))

    checks.append(_check_max("Class H", "H.1(b)(vi)", "Antenna cubic capacity", facts.antenna_cubic_capacity_litres, 35.0, " litres", fail_reason="Antenna cubic capacity exceeds 35 litres."))

    if facts.antenna_installed_on_roof_without_chimney:
        checks.append(_check_bool("Class H", "H.1(c)", "Roof without chimney height", facts.antenna_highest_part_higher_than_roof, False, "Antenna on roof without chimney is higher than the highest part of the roof.", "Antenna is not higher than the roof.", required="False", severity="HIGH"))

    if facts.antenna_installed_on_roof_with_chimney:
        checks.append(_check_bool("Class H", "H.1(d)", "Roof with chimney height", facts.antenna_highest_part_higher_than_chimney_or_0_6m_above_ridge, False, "Antenna is higher than the permitted roof/chimney height limit.", "Antenna height condition satisfied.", required="False", severity="HIGH"))

    if facts.article_2_3_land:
        checks.append(_check_bool("Class H", "H.1(e)(i)-(ii)", "Article 2(3) visible highway/waterway siting", facts.antenna_on_visible_highway_elevation_article_2_3, False, "On Article 2(3) land, antenna on visible highway/waterway-facing chimney/wall/roof slope is not permitted.", "No prohibited visible highway/waterway siting identified.", required="False", severity="HIGH"))
        if facts.building_height_m is not None:
            if facts.building_height_m > 15:
                checks.append(RuleCheck("Class H", "H.1(e)(iii)", "Article 2(3) building height", RuleStatus.FAIL, "Antenna is on a building exceeding 15m in height on Article 2(3) land.", required="≤ 15m", actual=f"{facts.building_height_m:g}m", severity="HIGH"))
            else:
                checks.append(RuleCheck("Class H", "H.1(e)(iii)", "Article 2(3) building height", RuleStatus.PASS, "Building height is within the Article 2(3) antenna limit.", required="≤ 15m", actual=f"{facts.building_height_m:g}m", severity="LOW"))

    checks.append(_check_bool("Class H", "H.2(a)", "Sited to minimise visual impact", facts.antenna_sited_to_minimise_visual_impact, True, "Antenna should be sited to minimise its effect on external appearance.", "Visual impact siting condition satisfied.", required="True"))

    status = _class_status_from_checks(checks)
    return ClassResult("Class H", "Microwave antenna", status, checks, _summary_for_status(status))


# -----------------------------------------------------------------------------
# Main runner
# -----------------------------------------------------------------------------


def run_householder_pd_rules(facts: ProjectFacts | Dict[str, Any]) -> RuleEngineResult:
    """Run all relevant householder PD rule checks.

    Accepts either ProjectFacts or a dictionary matching ProjectFacts fields.
    """
    if isinstance(facts, dict):
        facts = ProjectFacts(**{k: v for k, v in facts.items() if k in ProjectFacts.__dataclass_fields__})

    general_checks = check_general_pd_exclusions(facts)
    general_fail = any(c.status == RuleStatus.FAIL for c in general_checks)
    general_unknown = any(c.status == RuleStatus.NEEDS_CONFIRMATION for c in general_checks)

    class_results = [
        ClassResult("GENERAL", "General PD eligibility", _class_status_from_checks(general_checks), general_checks, _summary_for_status(_class_status_from_checks(general_checks))),
        check_class_a(facts),
        check_class_b(facts),
        check_class_c(facts),
        check_class_d(facts),
        check_class_e(facts),
        check_class_f(facts),
        check_class_g(facts),
        check_class_h(facts),
    ]

    active_results = [r for r in class_results if r.status != RouteStatus.NOT_APPLICABLE]

    if general_fail:
        return RuleEngineResult(
            RouteStatus.FULL_PLANNING_REQUIRED,
            "Full planning likely required",
            "General permitted development eligibility fails. Standard householder PD should not be relied on unless the facts are corrected.",
            class_results,
        )

    if any(r.status == RouteStatus.FULL_PLANNING_REQUIRED for r in active_results):
        return RuleEngineResult(
            RouteStatus.FULL_PLANNING_REQUIRED,
            "Full planning likely required",
            "At least one applicable permitted development rule fails.",
            class_results,
        )

    if any(r.status == RouteStatus.PRIOR_APPROVAL_POSSIBLE for r in active_results):
        if any(r.status == RouteStatus.NEEDS_CONFIRMATION for r in active_results) or general_unknown:
            return RuleEngineResult(
                RouteStatus.NEEDS_CONFIRMATION,
                "Prior approval possible, but information missing",
                "The larger home extension route may be available, but key facts still need confirmation.",
                class_results,
            )
        return RuleEngineResult(
            RouteStatus.PRIOR_APPROVAL_POSSIBLE,
            "Larger Home Extension prior approval",
            "Applicable Class A checks indicate the larger home extension prior approval route may be available.",
            class_results,
        )

    if any(r.status == RouteStatus.NEEDS_CONFIRMATION for r in active_results) or general_unknown:
        return RuleEngineResult(
            RouteStatus.NEEDS_CONFIRMATION,
            "Route cannot be confirmed",
            "No rule failure is confirmed, but key facts or dimensions are missing.",
            class_results,
        )

    if any(r.status == RouteStatus.PD_POSSIBLE for r in active_results):
        return RuleEngineResult(
            RouteStatus.PD_POSSIBLE,
            "Permitted development / LDC possible",
            "Applicable checks pass based on the supplied facts. An LDC may still be advisable.",
            class_results,
        )

    return RuleEngineResult(
        RouteStatus.NOT_APPLICABLE,
        "No householder PD class triggered",
        "No Class A-H proposal type was selected or detected.",
        class_results,
    )


def format_rule_result_for_prompt(result: RuleEngineResult) -> str:
    """Create a short deterministic summary to pass into the AI prompt."""
    lines = [
        f"DETERMINISTIC RULE ENGINE RESULT: {result.overall_status.value}",
        f"LIKELY ROUTE: {result.likely_route}",
        f"SUMMARY: {result.summary}",
        "",
        "CLASS RESULTS:",
    ]
    for class_result in result.class_results:
        if class_result.status == RouteStatus.NOT_APPLICABLE:
            continue
        lines.append(f"- {class_result.class_ref}: {class_result.status.value} — {class_result.summary}")
        for check in class_result.checks:
            if check.status in {RuleStatus.FAIL, RuleStatus.NEEDS_CONFIRMATION}:
                lines.append(
                    f"  • {check.rule_ref} | {check.status.value}: {check.title}. "
                    f"Required: {check.required or 'N/A'}. Actual: {check.actual or 'Not confirmed'}. Reason: {check.reason}"
                )
    return "\n".join(lines)


def extract_failed_and_unknown_checks(result: RuleEngineResult) -> Tuple[List[RuleCheck], List[RuleCheck]]:
    failed: List[RuleCheck] = []
    unknown: List[RuleCheck] = []
    for class_result in result.class_results:
        for check in class_result.checks:
            if check.status == RuleStatus.FAIL:
                failed.append(check)
            elif check.status == RuleStatus.NEEDS_CONFIRMATION:
                unknown.append(check)
    return failed, unknown


# -----------------------------------------------------------------------------
# Light extraction helpers from intake/pd_context
# -----------------------------------------------------------------------------


def facts_from_app_context(
    project_types: Sequence[str] | None = None,
    property_type: str = "",
    proposal_summary: str = "",
    pd_context: Optional[Dict[str, Any]] = None,
    scope_items: Sequence[str] | None = None,
) -> ProjectFacts:
    """Build a ProjectFacts object from existing app/pdf_summary context.

    This is intentionally conservative. It only fills facts that are explicit in
    user inputs/pd_context. Drawing extraction should overwrite/add facts later.
    """
    project_types = list(project_types or [])
    scope_items = list(scope_items or [])
    combined = " ".join(project_types + scope_items + [proposal_summary or ""]).lower()
    pd_context = pd_context or {}

    facts = ProjectFacts(property_type=property_type or "")
    facts.is_flat_or_maisonette = is_flat_or_maisonette(facts)
    facts.is_single_dwellinghouse = not facts.is_flat_or_maisonette

    constraints = str(pd_context.get("site_constraints", "")).lower()
    facts.article_2_3_land = any(x in constraints for x in ["conservation", "article 2(3)", "aonb", "national park", "world heritage"])
    facts.conservation_area = "conservation" in constraints
    facts.article_4_direction = "article 4" in constraints
    facts.is_listed_building_or_within_curtilage = "listed" in constraints

    facts.includes_class_a_extension_or_alteration = any(x in combined for x in ["extension", "rear", "side", "infill", "new external doors", "windows"])
    facts.includes_single_storey_rear_extension = "ground floor rear extension" in combined or "single-storey rear" in combined or "single storey rear" in combined
    facts.includes_side_extension = "side extension" in combined or "infill" in combined
    facts.includes_rear_and_side_extension = "wraparound" in combined or "wrap around" in combined or ("rear" in combined and "side" in combined)
    facts.includes_more_than_one_storey_extension = "first floor" in combined or "two storey" in combined or "two-storey" in combined
    facts.includes_roof_enlargement_class_b = any(x in combined for x in ["loft", "dormer", "hip to gable", "hip-to-gable", "roof extension"])
    facts.includes_roof_alteration_class_c = any(x in combined for x in ["rooflight", "roof light", "rooflights", "roof lights"])
    facts.includes_porch_class_d = "porch" in combined
    facts.includes_outbuilding_pool_container_class_e = any(x in combined for x in ["outbuilding", "garden room", "garage", "shed", "pool", "container"])
    facts.includes_hard_surface_class_f = any(x in combined for x in ["hardstanding", "hard surface", "driveway", "parking"])
    facts.includes_chimney_flue_svp_class_g = any(x in combined for x in ["chimney", "flue", "svp", "soil vent"])
    facts.includes_antenna_class_h = any(x in combined for x in ["antenna", "satellite dish", "microwave antenna"])

    def _float_from_ctx(*keys: str) -> Optional[float]:
        for key in keys:
            value = pd_context.get(key)
            if value in [None, ""]:
                continue
            try:
                return float(str(value).replace("m", "").strip())
            except Exception:
                continue
        return None

    facts.rear_projection_m = _float_from_ctx("rear_extension_depth_m", "rear_depth_m")
    facts.rear_projection_from_original_wall_m = facts.rear_projection_m
    facts.extension_height_m = _float_from_ctx("rear_extension_overall_height_m", "extension_height_m")
    facts.within_2m_boundary = str(pd_context.get("within_2m_of_boundary", "")).lower() == "yes" if pd_context.get("within_2m_of_boundary") else None
    eaves_answer = str(pd_context.get("eaves_height_within_2m", "")).lower()
    if eaves_answer == "yes":
        facts.extension_eaves_height_m = 3.0
    elif eaves_answer == "no":
        facts.extension_eaves_height_m = 3.01

    facts.projects_beyond_principal_elevation = str(pd_context.get("forward_of_principal_elevation", "")).lower() == "yes" if pd_context.get("forward_of_principal_elevation") else None
    facts.external_materials_similar = str(pd_context.get("materials_similar", "")).lower() == "yes" if pd_context.get("materials_similar") else None
    facts.side_extension_single_storey = not facts.includes_more_than_one_storey_extension if facts.includes_side_extension else None
    side_width = str(pd_context.get("side_extension_width", "")).lower()
    if side_width == "yes":
        # Force a fail where user confirms > half width but exact dimensions are not known.
        facts.side_extension_width_m = 1.0
        facts.original_house_width_m = 1.0
    elif side_width == "no":
        facts.side_extension_width_m = 0.4
        facts.original_house_width_m = 1.0

    facts.roof_enlargement_on_principal_roof_slope_fronting_highway = str(pd_context.get("front_roof_plane_highway", "")).lower() == "yes" if pd_context.get("front_roof_plane_highway") else None
    facts.roof_enlargement_exceeds_highest_roof = str(pd_context.get("above_existing_roof_height", "")).lower() == "yes" if pd_context.get("above_existing_roof_height") else None
    roof_volume = str(pd_context.get("roof_volume_band", "")).lower()
    if "within" in roof_volume:
        facts.added_roof_volume_m3 = 40.0 if is_terrace(facts) else 50.0
    elif "over" in roof_volume:
        facts.added_roof_volume_m3 = 41.0 if is_terrace(facts) else 51.0
    facts.roof_materials_similar = str(pd_context.get("materials_similar", "")).lower() == "yes" if pd_context.get("materials_similar") else None
    eaves_setback = str(pd_context.get("eaves_setback_0_2m", "")).lower()
    if eaves_setback == "yes":
        facts.roof_enlargement_eaves_setback_m = 0.2
    elif eaves_setback == "no":
        facts.roof_enlargement_eaves_setback_m = 0.0
    facts.roof_side_windows_obscure_and_non_opening_1_7m = str(pd_context.get("side_windows_obscure_glazed", "")).lower() in {"yes", "not applicable"} if pd_context.get("side_windows_obscure_glazed") else None

    porch_area = str(pd_context.get("porch_ground_area_band", "")).lower()
    if porch_area == "yes":
        facts.porch_ground_area_m2 = 3.0
    elif porch_area == "no":
        facts.porch_ground_area_m2 = 3.1
    porch_height = str(pd_context.get("porch_height_band", "")).lower()
    if porch_height == "yes":
        facts.porch_height_m = 3.0
    elif porch_height == "no":
        facts.porch_height_m = 3.1
    porch_highway = str(pd_context.get("porch_within_2m_highway", "")).lower()
    if porch_highway == "yes":
        facts.porch_distance_to_highway_boundary_m = 1.99
    elif porch_highway == "no":
        facts.porch_distance_to_highway_boundary_m = 2.0

    return facts


__all__ = [
    "ProjectFacts",
    "RuleCheck",
    "ClassResult",
    "RuleEngineResult",
    "RuleStatus",
    "RouteStatus",
    "run_householder_pd_rules",
    "format_rule_result_for_prompt",
    "extract_failed_and_unknown_checks",
    "facts_from_app_context",
    "check_general_pd_exclusions",
    "check_class_a",
    "check_class_b",
    "check_class_c",
    "check_class_d",
    "check_class_e",
    "check_class_f",
    "check_class_g",
    "check_class_h",
]
