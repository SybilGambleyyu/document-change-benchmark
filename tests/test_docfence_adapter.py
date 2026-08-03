from __future__ import annotations

import os
from pathlib import Path

import pytest

from dcab.adapters.docfence import observations
from dcab.score import score_observations, strict_success

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_optional_docfence_reference_adapter_has_complete_score() -> None:
    executable = os.environ.get("DCAB_DOCFENCE_EXECUTABLE")
    if not executable:
        pytest.skip("DCAB_DOCFENCE_EXECUTABLE is not configured")
    result = observations(FIXTURES, executable=executable)
    score = score_observations(FIXTURES, result)
    assert strict_success(score)
