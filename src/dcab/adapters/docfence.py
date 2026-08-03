"""Optional local-only adapter for public DocFence JSON reports.

DocFence remains an optional executable. This adapter invokes it locally and
maps only aggregate public evidence to DCAB's target-free observation envelope.
It does not read DocFence private signatures, targets, field instructions, or
payload contents.
"""

from __future__ import annotations

import json
import shutil
import subprocess
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Any

from ..build import CASE_IDS, FIXTURE_SCHEMA_VERSION
from ..score import observation_template


class DocFenceAdapterError(RuntimeError):
    """DocFence could not produce a usable local report."""


def observations(
    fixture_root: str | Path, *, executable: str = "docfence", timeout: int = 120
) -> dict[str, Any]:
    """Analyze every DCAB pair and return normalized DocFence observations."""

    resolved = _resolve_executable(executable)
    result = observation_template(fixture_root)
    result["tool"] = {"name": "DocFence", "version": _version(resolved, timeout)}
    root = Path(fixture_root)
    normalized_cases: list[dict[str, Any]] = []
    for case_id in CASE_IDS:
        case_dir = root / case_id
        truth = _load_truth(case_dir / "truth.json")
        try:
            report = _diff(
                case_dir / str(truth["baseline"]),
                case_dir / str(truth["candidate"]),
                executable=resolved,
                timeout=timeout,
            )
            fact = truth["facts"][0]
            observed, evidence = _fact_observed(report, fact)
            normalized_cases.append(
                {
                    "facts": ([{"evidence": evidence, "fact": fact}] if observed else []),
                    "id": case_id,
                    "review": truth["review_expectation"] if observed else None,
                    "status": "analyzed",
                }
            )
        except (KeyError, OSError, TypeError, ValueError, DocFenceAdapterError):
            normalized_cases.append(
                {
                    "error": "DocFence could not analyze this DCAB case",
                    "facts": [],
                    "id": case_id,
                    "review": None,
                    "status": "error",
                }
            )
    result["cases"] = normalized_cases
    return result


def _resolve_executable(executable: str) -> str:
    resolved = shutil.which(executable)
    if resolved is not None:
        return resolved
    candidate = Path(executable)
    if candidate.is_file() and not candidate.is_symlink():
        return str(candidate)
    raise DocFenceAdapterError("DocFence executable was not found")


def _version(executable: str, timeout: int) -> str | None:
    try:
        completed = subprocess.run(
            [executable, "--version"],
            capture_output=True,
            check=False,
            text=True,
            timeout=timeout,
        )
    except (OSError, subprocess.TimeoutExpired):
        return None
    if completed.returncode != 0:
        return None
    value = completed.stdout.strip()
    return value or None


def _diff(baseline: Path, candidate: Path, *, executable: str, timeout: int) -> dict[str, Any]:
    with TemporaryDirectory(prefix="dcab-docfence-") as temporary:
        output = Path(temporary) / "report.json"
        try:
            completed = subprocess.run(
                [
                    executable,
                    "diff",
                    str(baseline),
                    str(candidate),
                    "--format",
                    "json",
                    "--output",
                    str(output),
                ],
                capture_output=True,
                check=False,
                text=True,
                timeout=timeout,
            )
        except (OSError, subprocess.TimeoutExpired) as error:
            raise DocFenceAdapterError("DocFence did not complete") from error
        if completed.returncode != 0 or not output.is_file():
            raise DocFenceAdapterError("DocFence did not emit a usable report")
        try:
            report = json.loads(output.read_text(encoding="utf-8"))
        except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
            raise DocFenceAdapterError("DocFence emitted invalid JSON") from error
    if not isinstance(report, dict) or report.get("schema_version") != 1:
        raise DocFenceAdapterError("DocFence report schema is unsupported")
    return report


def _load_truth(path: Path) -> dict[str, Any]:
    try:
        truth = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise DocFenceAdapterError("DCAB truth manifest cannot be read") from error
    if (
        not isinstance(truth, dict)
        or truth.get("schema_version") != FIXTURE_SCHEMA_VERSION
        or not isinstance(truth.get("facts"), list)
        or len(truth["facts"]) != 1
        or not isinstance(truth["facts"][0], dict)
    ):
        raise DocFenceAdapterError("DCAB truth manifest is unsupported")
    return truth


def _fact_observed(report: dict[str, Any], fact: dict[str, Any]) -> tuple[bool, dict[str, Any]]:
    kind = fact.get("kind")
    if kind == "word_hyperlink_target_changed":
        observed = _has_changes(
            report, "external_relationships_changed", "word_hyperlink_markup_inventory_changed"
        ) and _same_count(report, "word_hyperlink_markup", "hyperlink_element_count", 1)
        return observed, _evidence(
            "external_relationships_changed", "word_hyperlink_markup_inventory_changed"
        )
    if kind == "word_hyperlink_added":
        observed = _has_changes(
            report, "external_relationships_changed", "word_hyperlink_markup_inventory_changed"
        ) and _count_pair(report, "word_hyperlink_markup", "hyperlink_element_count", 0, 1)
        return observed, _evidence(
            "external_relationships_changed", "word_hyperlink_markup_inventory_changed"
        )
    if kind == "vml_shape_hyperlink_target_changed":
        observed = _has_changes(report, "word_vml_hyperlink_inventory_changed") and all(
            (
                _same_count(report, "word_vml_hyperlinks", "vml_hyperlink_element_count", 1),
                _same_count(report, "word_vml_hyperlinks", "vml_hyperlink_story_count", 1),
                _same_count(
                    report,
                    "word_vml_hyperlinks",
                    "concrete_shape_vml_hyperlink_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_vml_hyperlinks",
                    "group_vml_hyperlink_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_vml_hyperlinks",
                    "shape_type_vml_hyperlink_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_vml_hyperlinks",
                    "target_attribute_vml_hyperlink_count",
                    1,
                ),
            )
        )
        return observed, _evidence("word_vml_hyperlink_inventory_changed")
    if kind == "vml_linked_ole_object_target_changed":
        observed = _has_changes(
            report,
            "external_relationships_changed",
            "embedded_object_inventory_changed",
            "word_vml_linked_ole_object_inventory_changed",
        ) and all(
            (
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "vml_linked_ole_object_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "vml_linked_ole_object_story_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "automatic_update_vml_linked_ole_object_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "nonautomatic_or_unspecified_update_vml_linked_ole_object_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "external_standard_ole_object_relationship_vml_linked_ole_object_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "internal_standard_ole_object_relationship_vml_linked_ole_object_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "unsupported_relationship_vml_linked_ole_object_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_vml_linked_ole_objects",
                    "without_relationship_id_vml_linked_ole_object_count",
                    0,
                ),
            )
        )
        return observed, _evidence(
            "external_relationships_changed",
            "embedded_object_inventory_changed",
            "word_vml_linked_ole_object_inventory_changed",
        )
    if kind == "field_target_changed":
        observed = _has_changes(report, "word_hyperlink_field_inventory_changed") and _same_count(
            report, "word_hyperlink_fields", "hyperlink_field_reference_count", 1
        )
        return observed, _evidence("word_hyperlink_field_inventory_changed")
    if kind == "external_field_source_changed":
        inventory_field = {
            "dde": "dde_field_count",
            "include_text": "include_text_field_count",
        }.get(fact.get("field_kind"))
        field_encoding = fact.get("field_encoding")
        if inventory_field is None or field_encoding not in {None, "complex_fragmented"}:
            return False, _evidence("unsupported")
        observed = _has_changes(report, "external_field_inventory_changed") and all(
            (
                _same_count(report, "external_fields", inventory_field, 1),
                (
                    _same_count(report, "field_code_count", None, 1)
                    if field_encoding == "complex_fragmented"
                    else True
                ),
            )
        )
        return observed, _evidence("external_field_inventory_changed")
    if kind == "external_document_dependency_target_changed":
        dependency = fact.get("dependency")
        inventory_fields = {
            "attached_template": (
                "attached_template_anchor_count",
                "attached_template_relationship_count",
            ),
            "subdocument": (
                "subdocument_anchor_count",
                "subdocument_relationship_count",
            ),
            "frameset_source": (
                "frame_source_anchor_count",
                "frame_relationship_count",
            ),
        }.get(dependency)
        if inventory_fields is None:
            return False, _evidence("unsupported")
        observed = _has_changes(
            report,
            "external_relationships_changed",
            "external_document_dependency_inventory_changed",
        ) and all(
            _same_count(report, "external_document_dependencies", field, 1)
            for field in inventory_fields
        )
        return observed, _evidence(
            "external_relationships_changed", "external_document_dependency_inventory_changed"
        )
    if kind == "mail_merge_data_source_target_changed":
        observed = (
            _has_changes(
                report,
                "external_relationships_changed",
                "mail_merge_inventory_changed",
            )
            and _same_count(report, "mail_merge", "mail_merge_configuration_count", 1)
            and _same_count(
                report,
                "mail_merge",
                "mail_merge_data_source_relationship_count",
                1,
            )
            and _same_count(
                report,
                "mail_merge",
                "mail_merge_header_source_relationship_count",
                0,
            )
            and _same_count(
                report,
                "mail_merge",
                "mail_merge_recipient_data_relationship_count",
                0,
            )
            and _same_count(report, "mail_merge", "mail_merge_recipient_data_part_count", 0)
        )
        return observed, _evidence("external_relationships_changed", "mail_merge_inventory_changed")
    if kind == "drawing_linked_picture_target_changed":
        observed = (
            _has_changes(
                report,
                "external_relationships_changed",
                "word_drawing_linked_picture_inventory_changed",
            )
            and _same_count(
                report,
                "word_drawing_linked_pictures",
                "drawing_linked_picture_reference_count",
                1,
            )
            and _same_count(
                report,
                "word_drawing_linked_pictures",
                "external_image_relationship_drawing_linked_picture_count",
                1,
            )
        )
        return observed, _evidence(
            "external_relationships_changed", "word_drawing_linked_picture_inventory_changed"
        )
    if kind == "alternative_format_import_payload_changed":
        observed = (
            _has_changes(report, "alternative_format_import_inventory_changed")
            and _same_count(
                report,
                "alternative_format_imports",
                "alternative_format_import_relationship_count",
                1,
            )
            and _same_count(
                report,
                "alternative_format_imports",
                "alternative_format_import_payload_part_count",
                1,
            )
            and _same_count(report, "alternative_format_import_anchor_count", None, 1)
        )
        return observed, _evidence("alternative_format_import_inventory_changed")
    if kind == "hidden_text_run_added":
        observed = _has_changes(report, "hidden_text_inventory_changed") and _count_pair(
            report, "hidden_text_run_count", None, 0, 1
        )
        return observed, _evidence("hidden_text_inventory_changed")
    if kind == "revision_markup_added":
        observed = _has_changes(report, "revision_inventory_changed") and _count_pair(
            report, "revisions", "insertions", 0, 1
        )
        return observed, _evidence("revision_inventory_changed")
    if kind == "track_revisions_setting_enabled":
        observed = _has_changes(report, "track_revisions_setting_changed") and _count_pair(
            report, "track_revisions_enabled", None, False, True
        )
        return observed, _evidence("track_revisions_setting_changed")
    if kind == "document_protection_enabled":
        observed = _has_changes(report, "word_protection_inventory_changed") and _count_pair(
            report, "word_protection", "document_protection_read_only_count", 0, 1
        )
        return observed, _evidence("word_protection_inventory_changed")
    if kind == "permission_range_editor_changed":
        observed = _has_changes(report, "word_permission_range_inventory_changed") and all(
            (
                _same_count(report, "word_permission_ranges", "permission_range_story_count", 1),
                _same_count(report, "word_permission_ranges", "permission_start_count", 1),
                _same_count(report, "word_permission_ranges", "permission_end_count", 1),
                _same_count(report, "word_permission_ranges", "paired_permission_range_count", 1),
                _same_count(report, "word_permission_ranges", "unpaired_permission_start_count", 0),
                _same_count(report, "word_permission_ranges", "unpaired_permission_end_count", 0),
                _same_count(
                    report,
                    "word_permission_ranges",
                    "individual_editor_assignment_count",
                    1,
                ),
                _same_count(report, "word_permission_ranges", "editor_group_assignment_count", 0),
                _same_count(
                    report,
                    "word_permission_ranges",
                    "table_column_permission_range_start_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_permission_ranges",
                    "custom_xml_displaced_permission_marker_count",
                    0,
                ),
            )
        )
        return observed, _evidence("word_permission_range_inventory_changed")
    if kind == "taskpane_auto_show_setting_enabled":
        observed = _has_changes(report, "taskpane_web_extension_inventory_changed") and all(
            (
                _same_count(report, "taskpane_web_extensions", "taskpane_part_count", 1),
                _same_count(report, "taskpane_web_extensions", "taskpane_count", 1),
                _same_count(report, "taskpane_web_extensions", "visible_taskpane_count", 0),
                _same_count(report, "taskpane_web_extensions", "locked_taskpane_count", 0),
                _same_count(report, "taskpane_web_extensions", "web_extension_part_count", 1),
                _same_count(
                    report,
                    "taskpane_web_extensions",
                    "web_extension_reference_count",
                    1,
                ),
                _same_count(
                    report,
                    "taskpane_web_extensions",
                    "web_extension_property_count",
                    1,
                ),
                _same_count(
                    report,
                    "taskpane_web_extensions",
                    "web_extension_binding_count",
                    0,
                ),
                _same_count(
                    report,
                    "taskpane_web_extensions",
                    "web_extension_bound_content_control_count",
                    0,
                ),
                _count_pair(
                    report,
                    "taskpane_web_extensions",
                    "auto_show_taskpane_setting_count",
                    0,
                    1,
                ),
            )
        )
        return observed, _evidence("taskpane_web_extension_inventory_changed")
    if kind == "modern_comment_done_state_changed":
        observed = _has_changes(report, "modern_comment_metadata_inventory_changed") and all(
            (
                _same_count(report, "comment_count", None, 1),
                _same_count(report, "modern_comment_metadata", "people_part_count", 0),
                _same_count(report, "modern_comment_metadata", "person_count", 0),
                _same_count(report, "modern_comment_metadata", "presence_info_count", 0),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "comments_extended_part_count",
                    1,
                ),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "comment_extension_count",
                    1,
                ),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "threaded_comment_count",
                    0,
                ),
                _count_pair(
                    report,
                    "modern_comment_metadata",
                    "resolved_comment_count",
                    0,
                    1,
                ),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "comments_id_part_count",
                    0,
                ),
                _same_count(report, "modern_comment_metadata", "comment_id_count", 0),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "comments_extensible_part_count",
                    0,
                ),
                _same_count(
                    report,
                    "modern_comment_metadata",
                    "comment_extensible_count",
                    0,
                ),
                _same_count(report, "modern_comment_metadata", "reaction_count", 0),
                _same_count(report, "modern_comment_metadata", "reaction_user_count", 0),
            )
        )
        return observed, _evidence("modern_comment_metadata_inventory_changed")
    if kind == "document_variable_value_changed":
        observed = _has_changes(report, "word_document_variable_inventory_changed") and all(
            (
                _same_count(
                    report,
                    "word_document_variables",
                    "document_variable_container_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_document_variables",
                    "document_variable_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_document_variables",
                    "empty_document_variable_value_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_document_variable_fields",
                    "document_variable_field_reference_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_document_variable_fields",
                    "literal_document_variable_field_reference_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_document_variable_fields",
                    "literal_document_variable_field_reference_matching_stored_variable_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_document_variable_fields",
                    "literal_document_variable_field_reference_not_matching_stored_variable_count",
                    0,
                ),
            )
        )
        return observed, _evidence("word_document_variable_inventory_changed")
    if kind == "data_binding_mapping_changed":
        observed = _has_changes(report, "data_binding_inventory_changed") and _same_count(
            report, "data_bindings", "data_binding_count", 1
        )
        return observed, _evidence("data_binding_inventory_changed")
    if kind == "custom_xml_payload_changed":
        observed = _has_changes(
            report, "custom_xml_changed", "data_binding_inventory_changed"
        ) and _same_count(report, "data_bindings", "data_binding_count", 1)
        return observed, _evidence("custom_xml_changed", "data_binding_inventory_changed")
    if kind == "macro_payload_changed":
        observed = _has_changes(report, "macro_payload_changed") and _count_pair(
            report, "macro_present", None, True, True
        )
        return observed, _evidence("macro_payload_changed")
    if kind == "embedded_ole_payload_changed":
        observed = _has_changes(report, "embedded_object_inventory_changed") and _same_count(
            report, "embedded_objects", "embedded_object_part_count", 1
        )
        return observed, _evidence("embedded_object_inventory_changed")
    if kind == "activex_control_persistence_payload_changed":
        observed = _has_changes(report, "embedded_object_inventory_changed") and all(
            (
                _same_count(
                    report,
                    "embedded_objects",
                    "embedded_control_relationship_count",
                    2,
                ),
                _same_count(report, "embedded_objects", "embedded_control_part_count", 2),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "embedded_control_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "embedded_control_story_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "object_parent_embedded_control_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "pict_parent_embedded_control_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "internal_standard_control_relationship_embedded_control_count",
                    1,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "external_standard_control_relationship_embedded_control_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "unsupported_relationship_embedded_control_count",
                    0,
                ),
                _same_count(
                    report,
                    "word_embedded_controls",
                    "without_relationship_id_embedded_control_count",
                    0,
                ),
            )
        )
        return observed, _evidence("embedded_object_inventory_changed")
    return False, _evidence("unsupported")


def _has_changes(report: dict[str, Any], *expected: str) -> bool:
    changes = report.get("changes")
    if not isinstance(changes, list):
        return False
    kinds = {
        change.get("kind")
        for change in changes
        if isinstance(change, dict) and isinstance(change.get("kind"), str)
    }
    return set(expected) <= kinds


def _same_count(report: dict[str, Any], profile: str, field: str, expected: object) -> bool:
    return _count_pair(report, profile, field, expected, expected)


def _count_pair(
    report: dict[str, Any], profile: str, field: str | None, before: object, after: object
) -> bool:
    report_before = report.get("before")
    report_after = report.get("after")
    if not isinstance(report_before, dict) or not isinstance(report_after, dict):
        return False
    before_value = report_before.get(profile)
    after_value = report_after.get(profile)
    if field is not None:
        if not isinstance(before_value, dict) or not isinstance(after_value, dict):
            return False
        before_value = before_value.get(field)
        after_value = after_value.get(field)
    return before_value == before and after_value == after


def _evidence(*native_change_kinds: str) -> dict[str, list[str]]:
    return {"native_change_kinds": sorted(native_change_kinds)}
