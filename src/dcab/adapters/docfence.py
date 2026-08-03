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
    if kind == "field_target_changed":
        observed = _has_changes(report, "word_hyperlink_field_inventory_changed") and _same_count(
            report, "word_hyperlink_fields", "hyperlink_field_reference_count", 1
        )
        return observed, _evidence("word_hyperlink_field_inventory_changed")
    if kind == "external_field_source_changed":
        observed = _has_changes(report, "external_field_inventory_changed") and _same_count(
            report, "external_fields", "include_text_field_count", 1
        )
        return observed, _evidence("external_field_inventory_changed")
    if kind == "external_document_dependency_target_changed":
        observed = (
            _has_changes(
                report,
                "external_relationships_changed",
                "external_document_dependency_inventory_changed",
            )
            and _same_count(
                report,
                "external_document_dependencies",
                "attached_template_anchor_count",
                1,
            )
            and _same_count(
                report,
                "external_document_dependencies",
                "attached_template_relationship_count",
                1,
            )
        )
        return observed, _evidence(
            "external_relationships_changed", "external_document_dependency_inventory_changed"
        )
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
