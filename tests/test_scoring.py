from __future__ import annotations

import json
from pathlib import Path

import pytest

from dcab.build import CASE_IDS
from dcab.score import ObservationError, observation_template, score_observations, strict_success

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_unsupported_template_is_visible_in_score() -> None:
    score = score_observations(FIXTURES, observation_template(FIXTURES))
    assert score["summary"]["fact_recall"] == 0.0
    assert score["summary"]["analysis_coverage"] == 0.0
    assert score["summary"]["unsupported_case_count"] == len(CASE_IDS)
    assert not strict_success(score)


def test_exact_public_facts_and_reference_policy_pass_strict() -> None:
    observations = observation_template(FIXTURES)
    observations["tool"] = {"name": "exact-static-reviewer", "version": "0"}
    observations["cases"] = [
        {
            "facts": [{"evidence": {"native_kind": "test"}, "fact": _truth(case_id)["facts"][0]}],
            "id": case_id,
            "review": _truth(case_id)["review_expectation"],
            "status": "analyzed",
        }
        for case_id in CASE_IDS
    ]
    score = score_observations(FIXTURES, observations)
    assert score["summary"]["fact_recall"] == 1.0
    assert score["summary"]["reference_policy_agreement"] == 1.0
    assert score["summary"]["complete_case_count"] == len(CASE_IDS)
    assert strict_success(score)


def test_unrecognized_fact_is_reviewable_and_prevents_strict_success() -> None:
    observations = _complete_observations()
    observations["cases"][0]["facts"].append({"fact": {"kind": "additional-observation"}})
    score = score_observations(FIXTURES, observations)
    assert score["summary"]["unrecognized_fact_count"] == 1
    assert score["cases"][0]["unrecognized_facts"] == [{"kind": "additional-observation"}]
    assert not strict_success(score)


def test_invalid_unsupported_case_cannot_claim_a_fact() -> None:
    observations = observation_template(FIXTURES)
    observations["cases"][0]["facts"] = [{"fact": _truth(CASE_IDS[0])["facts"][0]}]
    with pytest.raises(ObservationError, match="unsupported cases"):
        score_observations(FIXTURES, observations)


def _complete_observations() -> dict[str, object]:
    observations = observation_template(FIXTURES)
    observations["tool"] = {"name": "exact-static-reviewer"}
    observations["cases"] = [
        {
            "facts": [{"fact": _truth(case_id)["facts"][0]}],
            "id": case_id,
            "review": _truth(case_id)["review_expectation"],
            "status": "analyzed",
        }
        for case_id in CASE_IDS
    ]
    return observations


def _truth(case_id: str) -> dict[str, object]:
    return json.loads((FIXTURES / case_id / "truth.json").read_text(encoding="utf-8"))
