from __future__ import annotations

import hashlib
import json
import shutil
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

import pytest
from docx import Document
from docx.opc.package import OpcPackage

from dcab.build import CASE_IDS, CASE_SPECS, build_fixtures
from dcab.resources import bundled_fixture_root
from dcab.validate import FixtureValidationError, validate_fixture_tree

FIXTURES = Path(__file__).resolve().parents[1] / "fixtures"


def test_checked_in_fixtures_validate() -> None:
    assert validate_fixture_tree(FIXTURES) == {
        "case_count": 23,
        "fact_count": 23,
        "fixture_schema_version": 1,
    }


def test_fixture_generation_is_byte_reproducible(tmp_path: Path) -> None:
    rebuilt = tmp_path / "fixtures"
    assert build_fixtures(rebuilt) == {"case_count": 23, "fixture_schema_version": 1}
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
    assert loaded_document_count == 44
    assert loaded_package_count == 46


def test_mail_merge_source_pair_has_a_fixed_anchor_and_one_target_boundary() -> None:
    """The data-source marker stays fixed while only its external target changes."""

    case = FIXTURES / "external.mail_merge_data_source_target_retargeted"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/_rels/settings.xml.rels"
        ]
        assert baseline.read("word/document.xml") == candidate.read("word/document.xml")
        settings = baseline.read("word/settings.xml")
        assert settings == candidate.read("word/settings.xml")
        settings_root = ET.fromstring(settings)
        baseline_relationships = ET.fromstring(baseline.read("word/_rels/settings.xml.rels"))
        candidate_relationships = ET.fromstring(candidate.read("word/_rels/settings.xml.rels"))

    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    relationship_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    source = settings_root.find(f"{{{word_namespace}}}mailMerge/{{{word_namespace}}}dataSource")
    assert source is not None
    assert source.get(f"{{{relationship_namespace}}}id") == "rIdMailMergeSource"
    baseline_relationship = baseline_relationships.find(
        f"{{{package_relationship_namespace}}}Relationship[@Id='rIdMailMergeSource']"
    )
    candidate_relationship = candidate_relationships.find(
        f"{{{package_relationship_namespace}}}Relationship[@Id='rIdMailMergeSource']"
    )
    assert baseline_relationship is not None
    assert candidate_relationship is not None
    assert baseline_relationship.get("Type") == f"{relationship_namespace}/mailMergeSource"
    assert candidate_relationship.get("Type") == f"{relationship_namespace}/mailMergeSource"
    assert baseline_relationship.get("TargetMode") == "External"
    assert candidate_relationship.get("TargetMode") == "External"
    assert baseline_relationship.get("Target") != candidate_relationship.get("Target")


def test_dde_field_pair_has_a_fixed_shape_and_one_source_argument_boundary() -> None:
    """The result, application, and item stay fixed while the source argument changes."""

    case = FIXTURES / "external.dde_field_source_retargeted"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/document.xml"
        ]
        baseline_root = ET.fromstring(baseline.read("word/document.xml"))
        candidate_root = ET.fromstring(candidate.read("word/document.xml"))

    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    instruction_attribute = f"{{{word_namespace}}}instr"
    fields_before = list(baseline_root.iter(f"{{{word_namespace}}}fldSimple"))
    fields_after = list(candidate_root.iter(f"{{{word_namespace}}}fldSimple"))
    assert len(fields_before) == len(fields_after) == 1
    assert "DDE DCAB" in fields_before[0].get(instruction_attribute, "")
    assert "DDE DCAB" in fields_after[0].get(instruction_attribute, "")
    assert "approved-source.xlsx" in fields_before[0].get(instruction_attribute, "")
    assert "candidate-source.xlsx" in fields_after[0].get(instruction_attribute, "")
    assert list(fields_before[0].itertext()) == list(fields_after[0].itertext())


def test_document_variable_pair_has_a_fixed_field_and_one_value_boundary() -> None:
    """The field reference stays fixed while one persisted document-variable value changes."""

    case = FIXTURES / "binding.document_variable_value_changed"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/settings.xml"
        ]
        assert baseline.read("word/document.xml") == candidate.read("word/document.xml")
        baseline_settings = ET.fromstring(baseline.read("word/settings.xml"))
        candidate_settings = ET.fromstring(candidate.read("word/settings.xml"))
        document_root = ET.fromstring(baseline.read("word/document.xml"))

    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    doc_vars_before = baseline_settings.find(f"{{{word_namespace}}}docVars")
    doc_vars_after = candidate_settings.find(f"{{{word_namespace}}}docVars")
    assert doc_vars_before is not None
    assert doc_vars_after is not None
    assert len(doc_vars_before) == len(doc_vars_after) == 1
    variable_before = doc_vars_before[0]
    variable_after = doc_vars_after[0]
    assert variable_before.tag == variable_after.tag == f"{{{word_namespace}}}docVar"
    assert variable_before.get(f"{{{word_namespace}}}name") == "DCABReviewState"
    assert variable_after.get(f"{{{word_namespace}}}name") == "DCABReviewState"
    assert variable_before.get(f"{{{word_namespace}}}val") == "approved-state"
    assert variable_after.get(f"{{{word_namespace}}}val") == "candidate-state"
    fields = list(document_root.iter(f"{{{word_namespace}}}fldSimple"))
    assert len(fields) == 1
    assert fields[0].get(f"{{{word_namespace}}}instr") == " DOCVARIABLE DCABReviewState "


def test_vml_shape_hyperlink_pair_has_one_direct_target_boundary() -> None:
    """One legacy VML rectangle keeps its shape while its direct href changes."""

    case = FIXTURES / "interaction.vml_shape_hyperlink_target_retargeted"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/document.xml"
        ]
        assert baseline.read("word/_rels/document.xml.rels") == candidate.read(
            "word/_rels/document.xml.rels"
        )
        baseline_document = baseline.read("word/document.xml")
        candidate_document = candidate.read("word/document.xml")
        baseline_root = ET.fromstring(baseline_document)
        candidate_root = ET.fromstring(candidate_document)

    assert (
        candidate_document.replace(
            b"https://candidate.example.invalid/dcab-vml-shape",
            b"https://approved.example.invalid/dcab-vml-shape",
            1,
        )
        == baseline_document
    )
    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    vml_namespace = "urn:schemas-microsoft-com:vml"
    pict_tag = f"{{{word_namespace}}}pict"
    run_tag = f"{{{word_namespace}}}r"
    text_tag = f"{{{word_namespace}}}t"
    baseline_pictures = list(baseline_root.iter(pict_tag))
    candidate_pictures = list(candidate_root.iter(pict_tag))
    assert len(baseline_pictures) == len(candidate_pictures) == 1
    for root, picture in zip(
        (baseline_root, candidate_root),
        (baseline_pictures[0], candidate_pictures[0]),
        strict=True,
    ):
        assert picture.attrib == {}
        assert len(picture) == 1
        picture_runs = [run for run in root.iter(run_tag) if picture in run]
        assert len(picture_runs) == 1
        assert picture_runs[0].attrib == {}
        assert list(picture_runs[0]) == [picture]
    baseline_shape = baseline_pictures[0][0]
    candidate_shape = candidate_pictures[0][0]
    assert baseline_shape.tag == candidate_shape.tag == f"{{{vml_namespace}}}rect"
    assert baseline_shape.attrib == {
        "id": "DCABVmlLinkShape",
        "style": "width:1pt;height:1pt",
        "filled": "f",
        "stroked": "f",
        "href": "https://approved.example.invalid/dcab-vml-shape",
        "target": "_blank",
    }
    assert candidate_shape.attrib == {
        "id": "DCABVmlLinkShape",
        "style": "width:1pt;height:1pt",
        "filled": "f",
        "stroked": "f",
        "href": "https://candidate.example.invalid/dcab-vml-shape",
        "target": "_blank",
    }
    assert [node.text for node in baseline_root.iter(text_tag)] == [
        node.text for node in candidate_root.iter(text_tag)
    ]


def test_permission_range_pair_has_a_fixed_boundary_and_one_editor_change() -> None:
    """Only a synthetic individual editor assignment changes inside one marker pair."""

    case = FIXTURES / "review.permission_range_editor_changed"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/document.xml"
        ]
        baseline_root = ET.fromstring(baseline.read("word/document.xml"))
        candidate_root = ET.fromstring(candidate.read("word/document.xml"))

    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    permission_start_tag = f"{{{word_namespace}}}permStart"
    permission_end_tag = f"{{{word_namespace}}}permEnd"
    run_tag = f"{{{word_namespace}}}r"
    text_tag = f"{{{word_namespace}}}t"
    baseline_start = list(baseline_root.iter(permission_start_tag))
    candidate_start = list(candidate_root.iter(permission_start_tag))
    baseline_end = list(baseline_root.iter(permission_end_tag))
    candidate_end = list(candidate_root.iter(permission_end_tag))
    assert len(baseline_start) == len(candidate_start) == 1
    assert len(baseline_end) == len(candidate_end) == 1
    assert baseline_start[0].get(f"{{{word_namespace}}}id") == "0"
    assert candidate_start[0].get(f"{{{word_namespace}}}id") == "0"
    assert baseline_end[0].get(f"{{{word_namespace}}}id") == "0"
    assert candidate_end[0].get(f"{{{word_namespace}}}id") == "0"
    assert baseline_start[0].get(f"{{{word_namespace}}}ed") == "DCAB_EDITOR_BASELINE"
    assert candidate_start[0].get(f"{{{word_namespace}}}ed") == "DCAB_EDITOR_CANDIDATE"
    for root in (baseline_root, candidate_root):
        paragraphs = list(root.iter(f"{{{word_namespace}}}p"))
        matching = [
            paragraph
            for paragraph in paragraphs
            if any(child.tag == permission_start_tag for child in paragraph)
        ]
        assert len(matching) == 1
        children = list(matching[0])
        start_index = next(
            index for index, child in enumerate(children) if child.tag == permission_start_tag
        )
        assert children[start_index + 1].tag == run_tag
        assert children[start_index + 1][0].tag == text_tag
        assert children[start_index + 1][0].text == "DCAB editable-range carrier"
        assert children[start_index + 2].tag == permission_end_tag
    assert [node.text for node in baseline_root.iter(text_tag)] == [
        node.text for node in candidate_root.iter(text_tag)
    ]


def test_taskpane_auto_show_pair_has_fixed_internal_parts_and_one_value_boundary() -> None:
    """Only the stored auto-show property value changes in a fixed add-in topology."""

    case = FIXTURES / "interaction.taskpane_auto_show_setting_enabled"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/webextensions/webextension1.xml"
        ]
        assert baseline.read("word/document.xml") == candidate.read("word/document.xml")
        baseline_taskpanes = baseline.read("word/webextensions/taskpanes.xml")
        candidate_taskpanes = candidate.read("word/webextensions/taskpanes.xml")
        baseline_relationships = baseline.read("word/webextensions/_rels/taskpanes.xml.rels")
        candidate_relationships = candidate.read("word/webextensions/_rels/taskpanes.xml.rels")
        baseline_extension = baseline.read("word/webextensions/webextension1.xml")
        candidate_extension = candidate.read("word/webextensions/webextension1.xml")

    assert baseline_taskpanes == candidate_taskpanes
    assert baseline_relationships == candidate_relationships
    assert candidate_extension.replace(b'value="true"', b'value="false"', 1) == baseline_extension

    taskpane_namespace = "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
    web_extension_namespace = (
        "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
    )
    relationship_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    taskpanes = ET.fromstring(baseline_taskpanes)
    assert taskpanes.tag == f"{{{taskpane_namespace}}}taskpanes"
    assert len(taskpanes) == 1
    taskpane = taskpanes[0]
    assert taskpane.tag == f"{{{taskpane_namespace}}}taskpane"
    assert taskpane.attrib == {
        "dockstate": "right",
        "visibility": "0",
        "width": "350",
        "row": "0",
        "locked": "false",
    }
    assert len(taskpane) == 1
    reference = taskpane[0]
    assert reference.tag == f"{{{taskpane_namespace}}}webextensionref"
    assert reference.attrib == {f"{{{relationship_namespace}}}id": "rIdTaskpaneWebExtension"}

    relationships = ET.fromstring(baseline_relationships)
    assert relationships.tag == f"{{{package_relationship_namespace}}}Relationships"
    assert len(relationships) == 1
    relationship = relationships[0]
    assert relationship.attrib == {
        "Id": "rIdTaskpaneWebExtension",
        "Type": "http://schemas.microsoft.com/office/2011/relationships/webextension",
        "Target": "webextension1.xml",
    }

    baseline_root = ET.fromstring(baseline_extension)
    candidate_root = ET.fromstring(candidate_extension)
    expected_children = [
        "reference",
        "alternateReferences",
        "properties",
        "bindings",
        "snapshot",
    ]
    assert baseline_root.tag == candidate_root.tag == f"{{{web_extension_namespace}}}webextension"
    assert (
        baseline_root.attrib
        == candidate_root.attrib
        == {"id": "{3F08C2A1-681F-451E-95B6-001122334455}"}
    )
    assert [child.tag for child in baseline_root] == [
        f"{{{web_extension_namespace}}}{name}" for name in expected_children
    ]
    assert [child.tag for child in candidate_root] == [
        f"{{{web_extension_namespace}}}{name}" for name in expected_children
    ]
    assert (
        baseline_root[0].attrib
        == candidate_root[0].attrib
        == {
            "id": "{F928A11C-9164-4F8A-8D92-556677889900}",
            "version": "1.0.0.0",
            "store": "EXCatalog",
            "storeType": "EXCatalog",
        }
    )
    assert len(baseline_root[0]) == len(candidate_root[0]) == 0
    assert len(baseline_root[1]) == len(candidate_root[1]) == 0
    assert len(baseline_root[3]) == len(candidate_root[3]) == 0
    assert len(baseline_root[4]) == len(candidate_root[4]) == 0
    assert len(baseline_root[2]) == len(candidate_root[2]) == 1
    baseline_property = baseline_root[2][0]
    candidate_property = candidate_root[2][0]
    assert (
        baseline_property.tag == candidate_property.tag == f"{{{web_extension_namespace}}}property"
    )
    assert baseline_property.attrib == {
        "name": "Office.AutoShowTaskpaneWithDocument",
        "value": "false",
    }
    assert candidate_property.attrib == {
        "name": "Office.AutoShowTaskpaneWithDocument",
        "value": "true",
    }


def test_modern_comment_done_pair_has_fixed_anchor_and_one_state_boundary() -> None:
    """The classic comment stays fixed while one commentsExtended done value changes."""

    case = FIXTURES / "review.modern_comment_done_state_changed"
    with (
        zipfile.ZipFile(case / "baseline.docx") as baseline,
        zipfile.ZipFile(case / "candidate.docx") as candidate,
    ):
        members = sorted(baseline.namelist())
        assert members == sorted(candidate.namelist())
        assert [name for name in members if baseline.read(name) != candidate.read(name)] == [
            "word/commentsExtended.xml"
        ]
        baseline_document = baseline.read("word/document.xml")
        candidate_document = candidate.read("word/document.xml")
        baseline_comments = baseline.read("word/comments.xml")
        candidate_comments = candidate.read("word/comments.xml")
        baseline_relationships = baseline.read("word/_rels/document.xml.rels")
        candidate_relationships = candidate.read("word/_rels/document.xml.rels")
        baseline_content_types = baseline.read("[Content_Types].xml")
        candidate_content_types = candidate.read("[Content_Types].xml")
        baseline_extended = baseline.read("word/commentsExtended.xml")
        candidate_extended = candidate.read("word/commentsExtended.xml")

    assert baseline_document == candidate_document
    assert baseline_comments == candidate_comments
    assert baseline_relationships == candidate_relationships
    assert baseline_content_types == candidate_content_types
    assert candidate_extended.replace(b'w15:done="1"', b'w15:done="0"', 1) == baseline_extended

    word_namespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    word_2010_namespace = "http://schemas.microsoft.com/office/word/2010/wordml"
    word_2012_namespace = "http://schemas.microsoft.com/office/word/2012/wordml"
    relationship_namespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_relationship_namespace = "http://schemas.openxmlformats.org/package/2006/relationships"
    content_types_namespace = "http://schemas.openxmlformats.org/package/2006/content-types"
    document_root = ET.fromstring(baseline_document)
    comments_root = ET.fromstring(baseline_comments)
    baseline_extended_root = ET.fromstring(baseline_extended)
    candidate_extended_root = ET.fromstring(candidate_extended)

    start_tag = f"{{{word_namespace}}}commentRangeStart"
    end_tag = f"{{{word_namespace}}}commentRangeEnd"
    reference_tag = f"{{{word_namespace}}}commentReference"
    run_tag = f"{{{word_namespace}}}r"
    text_tag = f"{{{word_namespace}}}t"
    starts = list(document_root.iter(start_tag))
    ends = list(document_root.iter(end_tag))
    references = list(document_root.iter(reference_tag))
    assert len(starts) == len(ends) == len(references) == 1
    assert starts[0].attrib == {f"{{{word_namespace}}}id": "0"}
    assert ends[0].attrib == {f"{{{word_namespace}}}id": "0"}
    assert references[0].attrib == {f"{{{word_namespace}}}id": "0"}
    paragraphs = [
        paragraph
        for paragraph in document_root.iter(f"{{{word_namespace}}}p")
        if starts[0] in paragraph and ends[0] in paragraph
    ]
    assert len(paragraphs) == 1
    paragraph_children = list(paragraphs[0])
    start_index = paragraph_children.index(starts[0])
    assert paragraph_children[start_index + 1].tag == run_tag
    assert paragraph_children[start_index + 1][0].tag == text_tag
    assert paragraph_children[start_index + 1][0].text == "DCAB comment anchor carrier"
    assert paragraph_children[start_index + 2] is ends[0]
    assert paragraph_children[start_index + 3].tag == run_tag
    assert list(paragraph_children[start_index + 3]) == [references[0]]

    assert comments_root.tag == f"{{{word_namespace}}}comments"
    assert len(comments_root) == 1
    comment = comments_root[0]
    assert comment.tag == f"{{{word_namespace}}}comment"
    assert comment.attrib == {
        f"{{{word_namespace}}}id": "0",
        f"{{{word_namespace}}}author": "DCAB-FIXTURE-COMMENT-AUTHOR",
        f"{{{word_namespace}}}initials": "DCF",
        f"{{{word_namespace}}}date": "2026-08-03T00:00:00Z",
    }
    assert len(comment) == 1
    comment_paragraph = comment[0]
    assert comment_paragraph.tag == f"{{{word_namespace}}}p"
    assert comment_paragraph.attrib == {
        f"{{{word_2010_namespace}}}paraId": "0A0B0C0D",
        f"{{{word_2010_namespace}}}textId": "77777777",
    }
    assert len(comment_paragraph) == 2
    assert comment_paragraph[0].tag == run_tag
    assert comment_paragraph[0][0].tag == f"{{{word_namespace}}}annotationRef"
    assert comment_paragraph[1].tag == run_tag
    assert comment_paragraph[1][0].tag == text_tag
    assert comment_paragraph[1][0].text == "DCAB fixed review comment"

    relationships = ET.fromstring(baseline_relationships)
    assert relationships.tag == f"{{{package_relationship_namespace}}}Relationships"
    relationship_by_id = {relationship.get("Id"): relationship for relationship in relationships}
    assert relationship_by_id["rIdComments"].attrib == {
        "Id": "rIdComments",
        "Type": f"{relationship_namespace}/comments",
        "Target": "comments.xml",
    }
    assert relationship_by_id["rIdCommentsExtended"].attrib == {
        "Id": "rIdCommentsExtended",
        "Type": "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
        "Target": "commentsExtended.xml",
    }

    content_types = ET.fromstring(baseline_content_types)
    content_type_overrides = {
        child.get("PartName"): child.get("ContentType")
        for child in content_types
        if child.tag == f"{{{content_types_namespace}}}Override"
    }
    assert content_type_overrides["/word/comments.xml"] == (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
    )
    assert content_type_overrides["/word/commentsExtended.xml"] == (
        "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
    )

    assert (
        baseline_extended_root.tag
        == candidate_extended_root.tag
        == (f"{{{word_2012_namespace}}}commentsEx")
    )
    assert len(baseline_extended_root) == len(candidate_extended_root) == 1
    assert (
        baseline_extended_root[0].tag
        == candidate_extended_root[0].tag
        == (f"{{{word_2012_namespace}}}commentEx")
    )
    assert baseline_extended_root[0].attrib == {
        f"{{{word_2012_namespace}}}paraId": "0A0B0C0D",
        f"{{{word_2012_namespace}}}done": "0",
    }
    assert candidate_extended_root[0].attrib == {
        f"{{{word_2012_namespace}}}paraId": "0A0B0C0D",
        f"{{{word_2012_namespace}}}done": "1",
    }
    assert [node.text for node in document_root.iter(text_tag)] == [
        node.text for node in ET.fromstring(candidate_document).iter(text_tag)
    ]


def test_public_truth_excludes_generated_sensitive_material() -> None:
    forbidden = (
        "example.invalid",
        "rIdHyperlink",
        "rIdAttachedTemplate",
        "rIdMailMergeSource",
        "rIdSubDocument",
        "rIdLinkedPicture",
        "rIdAltChunk",
        "rIdVbaProject",
        "rIdOleObject",
        "vbaProject.bin",
        "oleObject1.bin",
        "urn:dcab:fixture",
        "DCAB inert",
        "DCAB synthetic alternate-content",
        "approved-source.xlsx",
        "candidate-source.xlsx",
        "Sheet1!R1C1",
        "DDE DCAB",
        "DCABReviewState",
        "approved-state",
        "candidate-state",
        "DCAB_EDITOR_BASELINE",
        "DCAB_EDITOR_CANDIDATE",
        "DCAB editable-range carrier",
        "rIdTaskpaneWebExtensions",
        "rIdTaskpaneWebExtension",
        "webextension1.xml",
        "rIdComments",
        "rIdCommentsExtended",
        "comments.xml",
        "commentsExtended.xml",
        "{3F08C2A1-681F-451E-95B6-001122334455}",
        "{F928A11C-9164-4F8A-8D92-556677889900}",
        "EXCatalog",
        "Office.AutoShowTaskpaneWithDocument",
        "DCABVmlLinkShape",
        "_blank",
        "0A0B0C0D",
        "77777777",
        "DCAB-FIXTURE-COMMENT-AUTHOR",
        "DCF",
        "2026-08-03T00:00:00Z",
        "DCAB comment anchor carrier",
        "DCAB fixed review comment",
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
