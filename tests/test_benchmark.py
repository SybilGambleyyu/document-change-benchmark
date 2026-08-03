from __future__ import annotations

import hashlib
import json
import shutil
from pathlib import Path

import pytest
from docx import Document
from docx.opc.package import OpcPackage

from dcab.build import CASE_IDS, CASE_SPECS, build_fixtures
from dcab.resources import bundled_fixture_root
from dcab.validate import FixtureValidationError, validate_fixture_tree

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_checked_in_fixtures_validate() -> None:
    assert validate_fixture_tree(FIXTURES) == {
        "case_count": 12,
        "fact_count": 12,
        "fixture_schema_version": 1,
    }


def test_fixture_generation_is_byte_reproducible(tmp_path: Path) -> None:
    rebuilt = tmp_path / "fixtures"
    assert build_fixtures(rebuilt) == {"case_count": 12, "fixture_schema_version": 1}
    assert _tree_digests(rebuilt) == _tree_digests(FIXTURES)


def test_bundled_fixture_data_matches_repository_fixture_tree() -> None:
    assert _tree_digests(bundled_fixture_root()) == _tree_digests(FIXTURES)


def test_python_docx_opens_every_docx_and_its_opc_reader_opens_all_packages() -> None:
    """Exercise a public independent reader without a Word/client runtime claim."""

    loaded_document_count = 0
    loaded_package_count = 0
    for spec in CASE_SPECS:
        for filename in (spec.baseline_name, spec.candidate_name):
            path = FIXTURES / spec.case_id / filename
            package = OpcPackage.open(path)
            assert package.main_document_part is not None
            loaded_package_count += 1
            if path.suffix == ".docx":
                document = Document(path)
                assert document.element.body is not None
                loaded_document_count += 1
    assert loaded_document_count == 22
    assert loaded_package_count == 24


def test_public_truth_excludes_generated_sensitive_material() -> None:
    forbidden = (
        "example.invalid",
        "rIdHyperlink",
        "rIdVbaProject",
        "rIdOleObject",
        "vbaProject.bin",
        "oleObject1.bin",
        "urn:dcab:fixture",
        "DCAB inert",
    )
    for case_id in CASE_IDS:
        content = (FIXTURES / case_id / "truth.json").read_text(encoding="utf-8")
        assert not any(value in content for value in forbidden)


def test_validator_rejects_a_changed_truth_manifest(tmp_path: Path) -> None:
    copied = tmp_path / "fixtures"
    shutil.copytree(FIXTURES, copied)
    target = copied / CASE_SPECS[0].case_id / "truth.json"
    truth = json.loads(target.read_text(encoding="utf-8"))
    truth["title"] = "tampered"
    target.write_text(json.dumps(truth), encoding="utf-8")
    with pytest.raises(FixtureValidationError, match="not reproducible"):
        validate_fixture_tree(copied)


def test_build_refuses_an_unknown_existing_entry(tmp_path: Path) -> None:
    target = tmp_path / "fixtures"
    target.mkdir()
    (target / "unrelated-file").write_text("do not overwrite", encoding="utf-8")
    with pytest.raises(ValueError, match="unknown entries"):
        build_fixtures(target, force=True)


def _tree_digests(root: Path) -> dict[str, str]:
    return {
        str(path.relative_to(root)): hashlib.sha256(path.read_bytes()).hexdigest()
        for path in sorted(root.rglob("*"))
        if path.is_file()
    }
