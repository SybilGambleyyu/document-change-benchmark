"""Tool-neutral observation protocol and deterministic DCAB scoring."""

from __future__ import annotations

import json
from collections import Counter
from pathlib import Path
from typing import Any

from .build import CASE_IDS, FIXTURE_SCHEMA_VERSION
from .validate import FixtureValidationError, validate_fixture_tree

OBSERVATION_SCHEMA_VERSION = 1
_CASE_STATUSES = frozenset({"analyzed", "unsupported", "error"})
_REVIEW_DISPOSITIONS = frozenset({"allow", "review", "block"})


class ObservationError(ValueError):
    """An observation report cannot be evaluated against DCAB."""


def observation_template(fixture_root: str | Path) -> dict[str, Any]:
    """Return a valid report that explicitly marks all cases unsupported."""

    _case_truths(fixture_root)
    return {
        "benchmark": {"fixture_schema_versions": [FIXTURE_SCHEMA_VERSION]},
        "cases": [
            {"facts": [], "id": case_id, "review": None, "status": "unsupported"}
            for case_id in CASE_IDS
        ],
        "schema_version": OBSERVATION_SCHEMA_VERSION,
        "tool": {"name": "replace-with-tool-name"},
    }


def load_observations(path: str | Path) -> dict[str, Any]:
    """Load an observation JSON document without scoring it."""

    return _load_json(Path(path), "observation report")


def score_observations(fixture_root: str | Path, observations: dict[str, Any]) -> dict[str, Any]:
    """Score normalized observations against DCAB's deliberately partial oracle.

    The score measures declared-fact recall and reference-policy agreement. An
    observation outside DCAB's narrow public contract remains visible as an
    unrecognized fact; it is never silently classified as a false positive.
    """

    truths = _case_truths(fixture_root)
    tool, reported_cases = _validate_observations(observations, expected_ids=set(truths))
    expected_fact_count = 0
    matched_fact_count = 0
    unrecognized_fact_count = 0
    analyzed_case_count = 0
    unsupported_case_count = 0
    errored_case_count = 0
    not_reported_case_count = 0
    review_reported_count = 0
    review_matched_count = 0
    review_mismatched_count = 0
    complete_case_count = 0
    case_results: list[dict[str, Any]] = []

    for case_id in CASE_IDS:
        truth = truths[case_id]
        expected_facts = truth["facts"]
        expected_fact_count += len(expected_facts)
        report = reported_cases.get(case_id)
        if report is None:
            status = "not_reported"
            reported_facts: list[dict[str, Any]] = []
            reported_review = None
            error = None
            not_reported_case_count += 1
        else:
            status = report["status"]
            reported_facts = [item["fact"] for item in report["facts"]]
            reported_review = report["review"]
            error = report["error"]
            if status == "analyzed":
                analyzed_case_count += 1
            elif status == "unsupported":
                unsupported_case_count += 1
            else:
                errored_case_count += 1

        matched, missing, unrecognized = _match_contract_items(expected_facts, reported_facts)
        matched_fact_count += len(matched)
        unrecognized_fact_count += len(unrecognized)
        review_matched = reported_review == truth["review_expectation"]
        if reported_review is not None:
            review_reported_count += 1
            if review_matched:
                review_matched_count += 1
            else:
                review_mismatched_count += 1
        complete = status == "analyzed" and not missing and not unrecognized and review_matched
        if complete:
            complete_case_count += 1
        case_results.append(
            {
                "complete": complete,
                "error": error,
                "expected_fact_count": len(expected_facts),
                "id": case_id,
                "matched_facts": matched,
                "missing_facts": missing,
                "review": {
                    "expected": truth["review_expectation"],
                    "matched": review_matched if reported_review is not None else False,
                    "reported": reported_review,
                },
                "status": status,
                "unrecognized_facts": unrecognized,
            }
        )

    case_count = len(CASE_IDS)
    return {
        "benchmark": {"case_count": case_count, "fixture_schema_version": FIXTURE_SCHEMA_VERSION},
        "cases": case_results,
        "schema_version": 1,
        "summary": {
            "analysis_coverage": _ratio(analyzed_case_count, case_count),
            "analyzed_case_count": analyzed_case_count,
            "complete_case_count": complete_case_count,
            "errored_case_count": errored_case_count,
            "expected_fact_count": expected_fact_count,
            "fact_recall": _ratio(matched_fact_count, expected_fact_count),
            "matched_fact_count": matched_fact_count,
            "not_reported_case_count": not_reported_case_count,
            "reference_policy_agreement": _ratio(review_matched_count, case_count),
            "review_matched_count": review_matched_count,
            "review_mismatched_count": review_mismatched_count,
            "review_reported_count": review_reported_count,
            "unrecognized_fact_count": unrecognized_fact_count,
            "unsupported_case_count": unsupported_case_count,
        },
        "tool": tool,
    }


def strict_success(score: dict[str, Any]) -> bool:
    """Return whether a score is complete and contains no undeclared facts."""

    summary = score.get("summary")
    benchmark = score.get("benchmark")
    return (
        isinstance(summary, dict)
        and isinstance(benchmark, dict)
        and summary.get("complete_case_count") == benchmark.get("case_count")
        and summary.get("unrecognized_fact_count") == 0
    )


def _case_truths(fixture_root: str | Path) -> dict[str, dict[str, Any]]:
    root = Path(fixture_root)
    try:
        validate_fixture_tree(root)
    except FixtureValidationError as error:
        raise ObservationError(f"fixtures are invalid: {error}") from error
    truths: dict[str, dict[str, Any]] = {}
    for case_id in CASE_IDS:
        truth = _load_json(root / case_id / "truth.json", "truth manifest")
        if truth.get("schema_version") != FIXTURE_SCHEMA_VERSION:
            raise ObservationError(f"{case_id}: fixture schema version is unsupported")
        if truth.get("id") != case_id:
            raise ObservationError(f"{case_id}: truth ID is invalid")
        facts = truth.get("facts")
        if not isinstance(facts, list) or not all(isinstance(fact, dict) for fact in facts):
            raise ObservationError(f"{case_id}: truth facts are invalid")
        if truth.get("review_expectation") not in _REVIEW_DISPOSITIONS:
            raise ObservationError(f"{case_id}: reference review convention is invalid")
        truths[case_id] = truth
    return truths


def _validate_observations(
    observations: dict[str, Any], *, expected_ids: set[str]
) -> tuple[dict[str, Any], dict[str, dict[str, Any]]]:
    if observations.get("schema_version") != OBSERVATION_SCHEMA_VERSION:
        raise ObservationError(f"observation schema_version must be {OBSERVATION_SCHEMA_VERSION}")
    benchmark = observations.get("benchmark")
    if not isinstance(benchmark, dict) or benchmark.get("fixture_schema_versions") != [
        FIXTURE_SCHEMA_VERSION
    ]:
        raise ObservationError("observation fixture_schema_versions is invalid")
    tool = observations.get("tool")
    if not isinstance(tool, dict):
        raise ObservationError("observations.tool must be an object")
    _require_string(tool.get("name"), "observations.tool.name")
    if "version" in tool and tool["version"] is not None and not isinstance(tool["version"], str):
        raise ObservationError("observations.tool.version must be a string when present")

    raw_cases = observations.get("cases")
    if not isinstance(raw_cases, list):
        raise ObservationError("observations.cases must be an array")
    cases: dict[str, dict[str, Any]] = {}
    for index, raw_case in enumerate(raw_cases):
        path = f"observations.cases[{index}]"
        if not isinstance(raw_case, dict):
            raise ObservationError(f"{path} must be an object")
        case_id = _require_string(raw_case.get("id"), f"{path}.id")
        if case_id not in expected_ids:
            raise ObservationError(f"{path} names an unknown DCAB case")
        if case_id in cases:
            raise ObservationError(f"{path} duplicates a DCAB case")
        status = raw_case.get("status")
        if status not in _CASE_STATUSES:
            raise ObservationError(f"{path}.status is invalid")
        raw_facts = raw_case.get("facts", [])
        if not isinstance(raw_facts, list):
            raise ObservationError(f"{path}.facts must be an array")
        facts: list[dict[str, Any]] = []
        for fact_index, item in enumerate(raw_facts):
            item_path = f"{path}.facts[{fact_index}]"
            if not isinstance(item, dict) or not isinstance(item.get("fact"), dict):
                raise ObservationError(f"{item_path}.fact must be an object")
            _require_string(item["fact"].get("kind"), f"{item_path}.fact.kind")
            evidence = item.get("evidence")
            if evidence is not None and not isinstance(evidence, dict):
                raise ObservationError(f"{item_path}.evidence must be an object when present")
            facts.append({"evidence": evidence, "fact": item["fact"]})
        review = raw_case.get("review")
        if review is not None and review not in _REVIEW_DISPOSITIONS:
            raise ObservationError(f"{path}.review is invalid")
        error = raw_case.get("error")
        if status == "analyzed":
            if error is not None:
                raise ObservationError(f"{path}: analyzed cases cannot contain an error")
        elif status == "unsupported":
            if facts or review is not None or error is not None:
                raise ObservationError(
                    f"{path}: unsupported cases cannot report facts, review, or error"
                )
        else:
            _require_string(error, f"{path}.error")
            if facts or review is not None:
                raise ObservationError(f"{path}: errored cases cannot report facts or review")
        cases[case_id] = {
            "error": error,
            "facts": facts,
            "id": case_id,
            "review": review,
            "status": status,
        }
    return {key: value for key, value in tool.items() if key in {"name", "version"}}, cases


def _match_contract_items(
    expected_items: list[dict[str, Any]], reported_items: list[dict[str, Any]]
) -> tuple[list[dict[str, Any]], list[dict[str, Any]], list[dict[str, Any]]]:
    expected_counts = Counter(_canonical_item(item) for item in expected_items)
    reported_counts = Counter(_canonical_item(item) for item in reported_items)
    matched_counts = expected_counts & reported_counts
    missing_counts = expected_counts - reported_counts
    unrecognized_counts = reported_counts - expected_counts
    return (
        _expand_counts(matched_counts),
        _expand_counts(missing_counts),
        _expand_counts(unrecognized_counts),
    )


def _canonical_item(item: dict[str, Any]) -> str:
    return json.dumps(item, sort_keys=True, separators=(",", ":"), ensure_ascii=False)


def _expand_counts(counts: Counter[str]) -> list[dict[str, Any]]:
    values: list[dict[str, Any]] = []
    for encoded in sorted(counts):
        values.extend(json.loads(encoded) for _ in range(counts[encoded]))
    return values


def _ratio(numerator: int, denominator: int) -> float:
    return numerator / denominator if denominator else 0.0


def _require_string(value: object, path: str) -> str:
    if not isinstance(value, str) or not value:
        raise ObservationError(f"{path} must be a nonempty string")
    return value


def _load_json(path: Path, label: str) -> dict[str, Any]:
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError) as error:
        raise ObservationError(f"{label} cannot be read") from error
    if not isinstance(value, dict):
        raise ObservationError(f"{label} must be a JSON object")
    return value
