"""Independent structural checks for generated DCAB fixture trees.

Validation reads ZIP members and XML only. It does not invoke Word, resolve a
relationship target, update a field, deserialize an opaque payload, or execute
stored code.
"""

from __future__ import annotations

import io
import json
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .build import (
    _ACTIVE_X_BINARY_CONTENT_TYPE,
    _ACTIVE_X_BINARY_RELATIONSHIP,
    _ACTIVE_X_CLASS_ID,
    _ACTIVE_X_CONTENT_TYPE,
    _ACTIVE_X_CONTROL_NAME,
    _ACTIVE_X_NS,
    _ALT_CHUNK_CONTENT_TYPE,
    _ALT_CHUNK_RELATIONSHIP,
    _ATTACHED_TEMPLATE_RELATIONSHIP,
    _COMMENTS_CONTENT_TYPE,
    _COMMENTS_EXTENDED_CONTENT_TYPE,
    _COMMENTS_EXTENDED_RELATIONSHIP,
    _COMMENTS_RELATIONSHIP,
    _CONTENT_TYPES_NS,
    _CONTROL_RELATIONSHIP,
    _CUSTOM_XML_PROPERTIES_CONTENT_TYPE,
    _CUSTOM_XML_PROPERTIES_NS,
    _CUSTOM_XML_PROPERTIES_RELATIONSHIP,
    _CUSTOM_XML_RELATIONSHIP,
    _DOCM_MAIN_CONTENT_TYPE,
    _DOCUMENT_VARIABLE_NAME,
    _DOCX_MAIN_CONTENT_TYPE,
    _DRAWING_NS,
    _FRAME_LAYOUT,
    _FRAME_NAME,
    _FRAME_RELATIONSHIP,
    _FRAME_SIZE,
    _HYPERLINK_RELATIONSHIP,
    _IMAGE_RELATIONSHIP,
    _INCLUDE_TEXT_FIELD_RESULT,
    _LINKED_OLE_OBJECT_ID,
    _LINKED_OLE_PROG_ID,
    _LINKED_OLE_SHAPE_ID,
    _LINKED_OLE_UPDATE_MODE,
    _MAIL_MERGE_SOURCE_RELATIONSHIP,
    _MODERN_COMMENT_ANCHOR_TEXT,
    _MODERN_COMMENT_AUTHOR,
    _MODERN_COMMENT_DATE,
    _MODERN_COMMENT_ID,
    _MODERN_COMMENT_INITIALS,
    _MODERN_COMMENT_PARAGRAPH_ID,
    _MODERN_COMMENT_TEXT,
    _MODERN_COMMENT_TEXT_ID,
    _OFFICE_DOCUMENT_RELATIONSHIP,
    _OFFICE_VML_NS,
    _OLE_CONTENT_TYPE,
    _OLE_OBJECT_RELATIONSHIP,
    _PACKAGE_REL_NS,
    _PERMISSION_RANGE_MARKER_ID,
    _PERMISSION_RANGE_TEXT,
    _PICTURE_NS,
    _REL_NS,
    _SETTINGS_CONTENT_TYPE,
    _SETTINGS_RELATIONSHIP,
    _STYLES_CONTENT_TYPE,
    _STYLES_RELATIONSHIP,
    _SUBDOCUMENT_RELATIONSHIP,
    _TASKPANE_AUTO_SHOW_PROPERTY_NAME,
    _TASKPANE_WEB_EXTENSION_CONTENT_TYPE,
    _TASKPANE_WEB_EXTENSION_ID,
    _TASKPANE_WEB_EXTENSION_NS,
    _TASKPANE_WEB_EXTENSION_REFERENCE_ID,
    _TASKPANE_WEB_EXTENSION_REFERENCE_STORE,
    _TASKPANE_WEB_EXTENSION_REFERENCE_STORE_TYPE,
    _TASKPANE_WEB_EXTENSION_REFERENCE_VERSION,
    _TASKPANE_WEB_EXTENSION_RELATIONSHIP,
    _TASKPANE_WEB_EXTENSION_TASKPANES_CONTENT_TYPE,
    _TASKPANE_WEB_EXTENSION_TASKPANES_NS,
    _TASKPANE_WEB_EXTENSION_TASKPANES_RELATIONSHIP,
    _VBA_PROJECT_CONTENT_TYPE,
    _VBA_PROJECT_RELATIONSHIP,
    _VISIBLE_TEXT,
    _VML_NS,
    _VML_SHAPE_ID,
    _VML_SHAPE_TARGET_FRAME,
    _WEB_SETTINGS_CONTENT_TYPE,
    _WEB_SETTINGS_RELATIONSHIP,
    _WORD_2010_WORDML_NS,
    _WORD_2012_WORDML_NS,
    _WORD_NS,
    _WORDPROCESSING_DRAWING_NS,
    CASE_IDS,
    CASE_SPECS,
    FIXTURE_SCHEMA_VERSION,
    CaseSpec,
    DocumentVariant,
    _complex_include_text_instruction_chunks,
    _dde_field_instruction,
    _document_variable_field_instruction,
    case_files,
    truth_manifest,
)

_XML_NS = "http://www.w3.org/XML/1998/namespace"


class FixtureValidationError(ValueError):
    """A fixture tree no longer matches DCAB's source contract."""


def validate_fixture_tree(fixture_root: str | Path) -> dict[str, int]:
    """Validate a complete deterministic fixture tree.

    The verifier establishes ZIP integrity, package skeleton shape, exact
    generated evidence, pair member boundaries, stable Word text nodes, and a
    target-free public oracle. It deliberately does not make a client-rendering
    or runtime-behavior claim.
    """

    root = Path(fixture_root)
    if not root.is_dir() or root.is_symlink():
        raise FixtureValidationError("fixture root must be a non-symlink directory")
    if {child.name for child in root.iterdir()} != {"manifest.jsonl", *CASE_IDS}:
        raise FixtureValidationError("fixture root does not match the DCAB catalogue")

    fact_count = 0
    for spec in CASE_SPECS:
        _validate_case(root / spec.case_id, spec)
        fact_count += len(truth_manifest(spec)["facts"])
    _validate_manifest(root / "manifest.jsonl")
    return {
        "case_count": len(CASE_SPECS),
        "fact_count": fact_count,
        "fixture_schema_version": FIXTURE_SCHEMA_VERSION,
    }


def _validate_case(case_dir: Path, spec: CaseSpec) -> None:
    if not case_dir.is_dir() or case_dir.is_symlink():
        _invalid(spec, "case directory is invalid")
    expected_files = case_files(spec)
    if {child.name for child in case_dir.iterdir()} != set(expected_files):
        _invalid(spec, "case files do not match source")

    actual_files: dict[str, bytes] = {}
    for name, generated in expected_files.items():
        path = case_dir / name
        if not path.is_file() or path.is_symlink():
            _invalid(spec, "case file is invalid")
        try:
            data = path.read_bytes()
        except OSError as error:
            raise FixtureValidationError(f"{spec.case_id}: case file cannot be read") from error
        if data != generated:
            _invalid(spec, "case file is not reproducible from source")
        actual_files[name] = data

    baseline = _read_archive(actual_files[spec.baseline_name], spec, "baseline")
    candidate = _read_archive(actual_files[spec.candidate_name], spec, "candidate")
    _validate_package(baseline, spec.baseline, spec, "baseline")
    _validate_package(candidate, spec.candidate, spec, "candidate")
    if set(baseline) != set(candidate):
        _invalid(spec, "pair package members differ")
    changed_members = tuple(name for name in sorted(baseline) if baseline[name] != candidate[name])
    if changed_members != spec.changed_members:
        _invalid(spec, "pair member boundary differs from source")
    if _stored_texts(baseline, spec) != _stored_texts(candidate, spec):
        _invalid(spec, "stored Word text is not stable")
    if _VISIBLE_TEXT not in _stored_texts(baseline, spec):
        _invalid(spec, "expected fixed text is absent")

    try:
        truth = json.loads(actual_files["truth.json"].decode("utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise FixtureValidationError(f"{spec.case_id}: truth manifest cannot be read") from error
    if truth != truth_manifest(spec):
        _invalid(spec, "public truth does not match source")
    _validate_public_truth(truth, spec)


def _validate_package(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    expected_members = {
        "[Content_Types].xml",
        "_rels/.rels",
        "word/document.xml",
        "word/_rels/document.xml.rels",
        "word/settings.xml",
        "word/styles.xml",
    }
    if variant.custom_xml_payload is not None:
        expected_members |= {
            "customXml/item1.xml",
            "customXml/_rels/item1.xml.rels",
            "customXml/itemProps1.xml",
        }
    if variant.macro_payload is not None:
        expected_members.add("word/vbaProject.bin")
    if variant.embedded_payload is not None:
        expected_members.add("word/embeddings/oleObject1.bin")
    if variant.activex_persistence_payload is not None:
        expected_members |= {
            "word/activeX/activeX1.xml",
            "word/activeX/_rels/activeX1.xml.rels",
            "word/activeX/activeX1.bin",
        }
    if variant.modern_comment_done is not None:
        expected_members |= {"word/comments.xml", "word/commentsExtended.xml"}
    if variant.alternative_format_import_payload is not None:
        expected_members.add("word/afchunk1.html")
    if variant.taskpane_auto_show is not None:
        expected_members |= {
            "word/webextensions/taskpanes.xml",
            "word/webextensions/_rels/taskpanes.xml.rels",
            "word/webextensions/webextension1.xml",
        }
    if variant.frameset_source_target is not None:
        expected_members |= {"word/webSettings.xml", "word/_rels/webSettings.xml.rels"}
    if (
        variant.attached_template_target is not None
        or variant.mail_merge_data_source_target is not None
    ):
        expected_members.add("word/_rels/settings.xml.rels")
    if set(members) != expected_members:
        _invalid(spec, f"{side} package members are invalid")

    _validate_content_types(members, variant, spec, side)
    _validate_root_relationships(members, spec, side)
    _validate_document_relationships(members, variant, spec, side)
    _validate_document_xml(members, variant, spec, side)
    _validate_active_x_control(members, variant, spec, side)
    _validate_modern_comment(members, variant, spec, side)
    _validate_taskpane_web_extension(members, variant, spec, side)
    _validate_alternative_format_import(members, variant, spec, side)
    _validate_settings(members, variant, spec, side)
    _validate_settings_relationships(members, variant, spec, side)
    _validate_web_settings(members, variant, spec, side)
    _validate_styles(members, spec, side)
    _validate_custom_xml(members, variant, spec, side)
    if (
        variant.macro_payload is not None
        and members["word/vbaProject.bin"] != variant.macro_payload
    ):
        _invalid(spec, f"{side} macro payload is invalid")
    if (
        variant.embedded_payload is not None
        and members["word/embeddings/oleObject1.bin"] != variant.embedded_payload
    ):
        _invalid(spec, f"{side} embedded OLE payload is invalid")
    if (
        variant.activex_persistence_payload is not None
        and members["word/activeX/activeX1.bin"] != variant.activex_persistence_payload
    ):
        _invalid(spec, f"{side} ActiveX persistence payload is invalid")


def _validate_content_types(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    root = _parse_xml(members["[Content_Types].xml"], spec)
    if root.tag != f"{{{_CONTENT_TYPES_NS}}}Types":
        _invalid(spec, f"{side} content types root is invalid")
    overrides: dict[str, str] = {}
    defaults: dict[str, str] = {}
    for child in root:
        if child.tag == f"{{{_CONTENT_TYPES_NS}}}Override":
            part_name = child.get("PartName")
            content_type = child.get("ContentType")
            if not part_name or not content_type or part_name in overrides:
                _invalid(spec, f"{side} content type override is invalid")
            overrides[part_name] = content_type
        elif child.tag == f"{{{_CONTENT_TYPES_NS}}}Default":
            extension = child.get("Extension")
            content_type = child.get("ContentType")
            if not extension or not content_type or extension in defaults:
                _invalid(spec, f"{side} content type default is invalid")
            defaults[extension] = content_type
        else:
            _invalid(spec, f"{side} content type entry is invalid")
    expected = {
        "/word/document.xml": (
            _DOCM_MAIN_CONTENT_TYPE if variant.extension == "docm" else _DOCX_MAIN_CONTENT_TYPE
        ),
        "/word/settings.xml": _SETTINGS_CONTENT_TYPE,
        "/word/styles.xml": _STYLES_CONTENT_TYPE,
    }
    if variant.custom_xml_payload is not None:
        expected["/customXml/itemProps1.xml"] = _CUSTOM_XML_PROPERTIES_CONTENT_TYPE
    if variant.alternative_format_import_payload is not None:
        expected["/word/afchunk1.html"] = _ALT_CHUNK_CONTENT_TYPE
    if variant.macro_payload is not None:
        expected["/word/vbaProject.bin"] = _VBA_PROJECT_CONTENT_TYPE
    if variant.embedded_payload is not None:
        expected["/word/embeddings/oleObject1.bin"] = _OLE_CONTENT_TYPE
    if variant.activex_persistence_payload is not None:
        expected.update(
            {
                "/word/activeX/activeX1.xml": _ACTIVE_X_CONTENT_TYPE,
                "/word/activeX/activeX1.bin": _ACTIVE_X_BINARY_CONTENT_TYPE,
            }
        )
    if variant.frameset_source_target is not None:
        expected["/word/webSettings.xml"] = _WEB_SETTINGS_CONTENT_TYPE
    if variant.modern_comment_done is not None:
        expected.update(
            {
                "/word/comments.xml": _COMMENTS_CONTENT_TYPE,
                "/word/commentsExtended.xml": _COMMENTS_EXTENDED_CONTENT_TYPE,
            }
        )
    if variant.taskpane_auto_show is not None:
        expected.update(
            {
                "/word/webextensions/taskpanes.xml": _TASKPANE_WEB_EXTENSION_TASKPANES_CONTENT_TYPE,
                "/word/webextensions/webextension1.xml": _TASKPANE_WEB_EXTENSION_CONTENT_TYPE,
            }
        )
    if overrides != expected:
        _invalid(spec, f"{side} content type overrides are invalid")
    if defaults != {
        "rels": "application/vnd.openxmlformats-package.relationships+xml",
        "xml": "application/xml",
    }:
        _invalid(spec, f"{side} content type defaults are invalid")


def _validate_root_relationships(members: dict[str, bytes], spec: CaseSpec, side: str) -> None:
    root = _parse_xml(members["_rels/.rels"], spec)
    if root.tag != f"{{{_PACKAGE_REL_NS}}}Relationships" or len(root) != 1:
        _invalid(spec, f"{side} package office-document relationship is invalid")
    relationship = root[0]
    expected = {
        "Id": "rIdOfficeDocument",
        "Target": "word/document.xml",
        "Type": _OFFICE_DOCUMENT_RELATIONSHIP,
    }
    if (
        relationship.tag != f"{{{_PACKAGE_REL_NS}}}Relationship"
        or {key: relationship.get(key) for key in expected} != expected
        or relationship.get("TargetMode") is not None
    ):
        _invalid(spec, f"{side} package office-document relationship is invalid")


def _validate_document_relationships(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    actual = _relationship_map(members["word/_rels/document.xml.rels"], spec)
    expected = {
        "rIdStyles": (_STYLES_RELATIONSHIP, "styles.xml", "internal"),
        "rIdSettings": (_SETTINGS_RELATIONSHIP, "settings.xml", "internal"),
    }
    if variant.direct_hyperlink_target is not None:
        expected["rIdHyperlink"] = (
            _HYPERLINK_RELATIONSHIP,
            variant.direct_hyperlink_target,
            "external",
        )
    if variant.drawing_linked_picture_target is not None:
        expected["rIdLinkedPicture"] = (
            _IMAGE_RELATIONSHIP,
            variant.drawing_linked_picture_target,
            "external",
        )
    if variant.subdocument_target is not None:
        expected["rIdSubDocument"] = (
            _SUBDOCUMENT_RELATIONSHIP,
            variant.subdocument_target,
            "external",
        )
    if variant.frameset_source_target is not None:
        expected["rIdWebSettings"] = (
            _WEB_SETTINGS_RELATIONSHIP,
            "webSettings.xml",
            "internal",
        )
    if variant.vml_linked_ole_target is not None:
        expected["rIdLinkedOleObject"] = (
            _OLE_OBJECT_RELATIONSHIP,
            variant.vml_linked_ole_target,
            "external",
        )
    if variant.alternative_format_import_payload is not None:
        expected["rIdAltChunk"] = (_ALT_CHUNK_RELATIONSHIP, "afchunk1.html", "internal")
    if variant.custom_xml_payload is not None:
        expected["rIdCustomXml"] = (_CUSTOM_XML_RELATIONSHIP, "../customXml/item1.xml", "internal")
    if variant.macro_payload is not None:
        expected["rIdVbaProject"] = (_VBA_PROJECT_RELATIONSHIP, "vbaProject.bin", "internal")
    if variant.embedded_payload is not None:
        expected["rIdOleObject"] = (
            _OLE_OBJECT_RELATIONSHIP,
            "embeddings/oleObject1.bin",
            "internal",
        )
    if variant.activex_persistence_payload is not None:
        expected["rIdActiveXControl"] = (
            _CONTROL_RELATIONSHIP,
            "activeX/activeX1.xml",
            "internal",
        )
    if variant.modern_comment_done is not None:
        expected.update(
            {
                "rIdComments": (_COMMENTS_RELATIONSHIP, "comments.xml", "internal"),
                "rIdCommentsExtended": (
                    _COMMENTS_EXTENDED_RELATIONSHIP,
                    "commentsExtended.xml",
                    "internal",
                ),
            }
        )
    if variant.taskpane_auto_show is not None:
        expected["rIdTaskpaneWebExtensions"] = (
            _TASKPANE_WEB_EXTENSION_TASKPANES_RELATIONSHIP,
            "webextensions/taskpanes.xml",
            "internal",
        )
    if actual != expected:
        _invalid(spec, f"{side} document relationships are invalid")


def _validate_document_xml(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    root = _parse_xml(members["word/document.xml"], spec)
    if root.tag != _word_tag("document"):
        _invalid(spec, f"{side} document root is invalid")
    body = root.find(_word_tag("body"))
    if body is None or body.find(_word_tag("sectPr")) is None:
        _invalid(spec, f"{side} document body is invalid")

    subdocuments = [element for element in body if element.tag == _word_tag("subDoc")]
    if len(subdocuments) != int(variant.subdocument_target is not None):
        _invalid(spec, f"{side} subdocument markup is invalid")
    if subdocuments and subdocuments[0].get(f"{{{_REL_NS}}}id") != "rIdSubDocument":
        _invalid(spec, f"{side} subdocument relationship is invalid")

    alt_chunks = [element for element in body if element.tag == _word_tag("altChunk")]
    if len(alt_chunks) != int(variant.alternative_format_import_payload is not None):
        _invalid(spec, f"{side} alternative-format import markup is invalid")
    if alt_chunks and alt_chunks[0].get(f"{{{_REL_NS}}}id") != "rIdAltChunk":
        _invalid(spec, f"{side} alternative-format import relationship is invalid")

    hyperlinks = list(root.iter(_word_tag("hyperlink")))
    expected_hyperlink_count = 1 if variant.direct_hyperlink_target is not None else 0
    if len(hyperlinks) != expected_hyperlink_count:
        _invalid(spec, f"{side} direct hyperlink markup is invalid")
    if hyperlinks and hyperlinks[0].get(f"{{{_REL_NS}}}id") != "rIdHyperlink":
        _invalid(spec, f"{side} direct hyperlink relationship is invalid")

    vml_pictures = list(root.iter(_word_tag("pict")))
    expected_vml_shape_count = int(variant.vml_shape_hyperlink_target is not None)
    if len(vml_pictures) != expected_vml_shape_count:
        _invalid(spec, f"{side} VML shape hyperlink markup is invalid")
    if vml_pictures:
        picture = vml_pictures[0]
        picture_runs = [run for run in root.iter(_word_tag("r")) if picture in run]
        if (
            picture.attrib
            or len(picture) != 1
            or (picture.text or "").strip()
            or (picture.tail or "").strip()
            or len(picture_runs) != 1
            or picture_runs[0].attrib
            or len(picture_runs[0]) != 1
            or (picture_runs[0].text or "").strip()
            or (picture_runs[0].tail or "").strip()
        ):
            _invalid(spec, f"{side} VML shape hyperlink markup is invalid")
        shape = picture[0]
        expected_shape_attributes = {
            "id": _VML_SHAPE_ID,
            "style": "width:1pt;height:1pt",
            "filled": "f",
            "stroked": "f",
            "href": variant.vml_shape_hyperlink_target,
            "target": _VML_SHAPE_TARGET_FRAME,
        }
        if (
            shape.tag != f"{{{_VML_NS}}}rect"
            or shape.attrib != expected_shape_attributes
            or list(shape)
            or (shape.text or "").strip()
            or (shape.tail or "").strip()
        ):
            _invalid(spec, f"{side} VML shape hyperlink markup is invalid")

    linked_picture_blips = list(root.iter(_drawing_tag("blip")))
    expected_linked_picture_count = int(variant.drawing_linked_picture_target is not None)
    if len(linked_picture_blips) != expected_linked_picture_count:
        _invalid(spec, f"{side} linked-picture markup is invalid")
    if linked_picture_blips:
        linked_picture = linked_picture_blips[0]
        if (
            linked_picture.get(f"{{{_REL_NS}}}link") != "rIdLinkedPicture"
            or linked_picture.get(f"{{{_REL_NS}}}embed") is not None
        ):
            _invalid(spec, f"{side} linked-picture relationship is invalid")
        graphic_data = list(root.iter(_drawing_tag("graphicData")))
        inlines = list(root.iter(_wordprocessing_drawing_tag("inline")))
        pictures = list(root.iter(_picture_tag("pic")))
        if (
            len(graphic_data) != 1
            or graphic_data[0].get("uri") != _PICTURE_NS
            or len(inlines) != 1
            or len(pictures) != 1
        ):
            _invalid(spec, f"{side} linked-picture DrawingML shape is invalid")

    fields = list(root.iter(_word_tag("fldSimple")))
    expected_field_instructions: list[str] = []
    if variant.hyperlink_field_target is not None:
        expected_field_instructions.append(f' HYPERLINK "{variant.hyperlink_field_target}" ')
    if variant.include_text_target is not None:
        expected_field_instructions.append(f' INCLUDETEXT "{variant.include_text_target}" ')
    if variant.dde_source_file is not None:
        expected_field_instructions.append(_dde_field_instruction(variant.dde_source_file))
    if variant.document_variable_value is not None:
        expected_field_instructions.append(_document_variable_field_instruction())
    if len(fields) != len(expected_field_instructions):
        _invalid(spec, f"{side} field markup is invalid")
    for field, expected_instruction in zip(fields, expected_field_instructions, strict=True):
        instruction = field.get(_word_tag("instr"))
        if instruction != expected_instruction:
            _invalid(spec, f"{side} field instruction is invalid")

    _validate_complex_include_text_field(root, variant, spec, side)

    permission_starts = list(root.iter(_word_tag("permStart")))
    permission_ends = list(root.iter(_word_tag("permEnd")))
    expected_permission_range_count = int(variant.permission_range_editor is not None)
    if (
        len(permission_starts) != expected_permission_range_count
        or len(permission_ends) != expected_permission_range_count
    ):
        _invalid(spec, f"{side} editable-range permission markup is invalid")
    if permission_starts:
        permission_start = permission_starts[0]
        permission_end = permission_ends[0]
        expected_start_attributes = {
            _word_tag("id"): _PERMISSION_RANGE_MARKER_ID,
            _word_tag("ed"): variant.permission_range_editor,
        }
        expected_end_attributes = {_word_tag("id"): _PERMISSION_RANGE_MARKER_ID}
        if (
            permission_start.attrib != expected_start_attributes
            or permission_end.attrib != expected_end_attributes
            or list(permission_start)
            or list(permission_end)
            or (permission_start.text or "").strip()
            or (permission_end.text or "").strip()
            or (permission_start.tail or "").strip()
            or (permission_end.tail or "").strip()
        ):
            _invalid(spec, f"{side} editable-range permission markup is invalid")
        permission_paragraphs = [
            paragraph
            for paragraph in root.iter(_word_tag("p"))
            if permission_start in paragraph or permission_end in paragraph
        ]
        if len(permission_paragraphs) != 1:
            _invalid(spec, f"{side} editable-range permission markup is invalid")
        paragraph_children = list(permission_paragraphs[0])
        start_index = paragraph_children.index(permission_start)
        end_index = paragraph_children.index(permission_end)
        if end_index != start_index + 2:
            _invalid(spec, f"{side} editable-range permission markup is invalid")
        carrier = paragraph_children[start_index + 1]
        if (
            carrier.tag != _word_tag("r")
            or carrier.attrib
            or len(carrier) != 1
            or carrier[0].tag != _word_tag("t")
            or carrier[0].attrib
            or carrier[0].text != _PERMISSION_RANGE_TEXT
            or list(carrier[0])
            or (carrier[0].tail or "").strip()
        ):
            _invalid(spec, f"{side} editable-range permission markup is invalid")

    comment_starts = list(root.iter(_word_tag("commentRangeStart")))
    comment_ends = list(root.iter(_word_tag("commentRangeEnd")))
    comment_references = list(root.iter(_word_tag("commentReference")))
    expected_comment_anchor_count = int(variant.modern_comment_done is not None)
    if (
        len(comment_starts) != expected_comment_anchor_count
        or len(comment_ends) != expected_comment_anchor_count
        or len(comment_references) != expected_comment_anchor_count
    ):
        _invalid(spec, f"{side} modern-comment anchor markup is invalid")
    if comment_starts:
        comment_start = comment_starts[0]
        comment_end = comment_ends[0]
        comment_reference = comment_references[0]
        expected_anchor_attributes = {_word_tag("id"): _MODERN_COMMENT_ID}
        if not (
            _is_empty_element(
                comment_start, _word_tag("commentRangeStart"), expected_anchor_attributes
            )
            and _is_empty_element(
                comment_end, _word_tag("commentRangeEnd"), expected_anchor_attributes
            )
            and _is_empty_element(
                comment_reference, _word_tag("commentReference"), expected_anchor_attributes
            )
        ):
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")
        reference_runs = [run for run in root.iter(_word_tag("r")) if comment_reference in run]
        if len(reference_runs) != 1:
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")
        reference_run = reference_runs[0]
        if (
            reference_run.attrib
            or len(reference_run) != 1
            or (reference_run.text or "").strip()
            or (reference_run.tail or "").strip()
        ):
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")
        comment_paragraphs = [
            paragraph
            for paragraph in root.iter(_word_tag("p"))
            if comment_start in paragraph or comment_end in paragraph or reference_run in paragraph
        ]
        if len(comment_paragraphs) != 1:
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")
        paragraph_children = list(comment_paragraphs[0])
        start_index = paragraph_children.index(comment_start)
        end_index = paragraph_children.index(comment_end)
        reference_index = paragraph_children.index(reference_run)
        if end_index != start_index + 2 or reference_index != end_index + 1:
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")
        carrier = paragraph_children[start_index + 1]
        if not _is_comment_text_run(carrier, _MODERN_COMMENT_ANCHOR_TEXT):
            _invalid(spec, f"{side} modern-comment anchor markup is invalid")

    hidden_count = sum(1 for run in root.iter(_word_tag("r")) if _run_has_vanish(run))
    if hidden_count != int(variant.hidden_text):
        _invalid(spec, f"{side} hidden-text markup is invalid")
    insertions = list(root.iter(_word_tag("ins")))
    if len(insertions) != int(variant.insertion_markup):
        _invalid(spec, f"{side} revision markup is invalid")
    if insertions and (
        insertions[0].get(_word_tag("id")) != "1"
        or insertions[0].get(_word_tag("author")) != "DCAB"
    ):
        _invalid(spec, f"{side} revision markup is invalid")

    bindings = list(root.iter(_word_tag("dataBinding")))
    if len(bindings) != int(variant.data_binding_xpath is not None):
        _invalid(spec, f"{side} data-binding markup is invalid")
    if bindings:
        binding = bindings[0]
        if (
            binding.get(_word_tag("xpath")) != variant.data_binding_xpath
            or binding.get(_word_tag("storeItemID")) is None
            or binding.get(_word_tag("prefixMappings")) is None
        ):
            _invalid(spec, f"{side} data-binding markup is invalid")

    ole_markers = [
        element
        for element in root.iter()
        if element.tag == f"{{{_OFFICE_VML_NS}}}OLEObject"
        and element.get("Type", "").casefold() == "embed"
    ]
    if len(ole_markers) != int(variant.embedded_payload is not None):
        _invalid(spec, f"{side} embedded OLE marker is invalid")
    if ole_markers and ole_markers[0].get(f"{{{_REL_NS}}}id") != "rIdOleObject":
        _invalid(spec, f"{side} embedded OLE relationship is invalid")

    active_x_controls = list(root.iter(_word_tag("control")))
    if len(active_x_controls) != int(variant.activex_persistence_payload is not None):
        _invalid(spec, f"{side} ActiveX embedded-control markup is invalid")
    if active_x_controls:
        active_x_control = active_x_controls[0]
        active_x_containers = [
            element for element in root.iter(_word_tag("object")) if active_x_control in element
        ]
        active_x_runs = [
            element
            for element in root.iter(_word_tag("r"))
            if active_x_containers and active_x_containers[0] in element
        ]
        expected_control_attributes = {
            f"{{{_REL_NS}}}id": "rIdActiveXControl",
            _word_tag("name"): _ACTIVE_X_CONTROL_NAME,
        }
        if (
            len(active_x_containers) != 1
            or len(active_x_runs) != 1
            or active_x_runs[0].attrib
            or len(active_x_runs[0]) != 1
            or active_x_runs[0][0] is not active_x_containers[0]
            or (active_x_runs[0].text or "").strip()
            or (active_x_runs[0].tail or "").strip()
        ):
            _invalid(spec, f"{side} ActiveX embedded-control markup is invalid")
        active_x_container = active_x_containers[0]
        if (
            active_x_container.attrib
            or len(active_x_container) != 1
            or active_x_container[0] is not active_x_control
            or (active_x_container.text or "").strip()
            or (active_x_container.tail or "").strip()
            or active_x_control.attrib != expected_control_attributes
            or list(active_x_control)
            or (active_x_control.text or "").strip()
            or (active_x_control.tail or "").strip()
        ):
            _invalid(spec, f"{side} ActiveX embedded-control markup is invalid")

    linked_ole_markers = [
        element
        for element in root.iter()
        if element.tag == f"{{{_OFFICE_VML_NS}}}OLEObject"
        and element.get("Type", "").casefold() == "link"
    ]
    if len(linked_ole_markers) != int(variant.vml_linked_ole_target is not None):
        _invalid(spec, f"{side} VML linked-OLE marker is invalid")
    if linked_ole_markers:
        linked_ole_marker = linked_ole_markers[0]
        linked_ole_containers = [
            element for element in root.iter(_word_tag("object")) if linked_ole_marker in element
        ]
        linked_ole_runs = [
            element
            for element in root.iter(_word_tag("r"))
            if linked_ole_containers and linked_ole_containers[0] in element
        ]
        expected_marker_attributes = {
            "Type": "Link",
            "ProgID": _LINKED_OLE_PROG_ID,
            "ShapeID": _LINKED_OLE_SHAPE_ID,
            "DrawAspect": "Content",
            "ObjectID": _LINKED_OLE_OBJECT_ID,
            f"{{{_REL_NS}}}id": "rIdLinkedOleObject",
            "UpdateMode": _LINKED_OLE_UPDATE_MODE,
        }
        expected_shape_attributes = {
            "id": _LINKED_OLE_SHAPE_ID,
            "style": "width:1pt;height:1pt",
            f"{{{_OFFICE_VML_NS}}}ole": "",
        }
        if (
            len(linked_ole_containers) != 1
            or len(linked_ole_runs) != 1
            or linked_ole_runs[0].attrib
            or list(linked_ole_runs[0]) != linked_ole_containers
            or (linked_ole_runs[0].text or "").strip()
            or (linked_ole_runs[0].tail or "").strip()
        ):
            _invalid(spec, f"{side} VML linked-OLE marker is invalid")
        linked_ole_container = linked_ole_containers[0]
        if (
            linked_ole_container.attrib
            or len(linked_ole_container) != 2
            or (linked_ole_container.text or "").strip()
            or (linked_ole_container.tail or "").strip()
        ):
            _invalid(spec, f"{side} VML linked-OLE marker is invalid")
        shape, marker = linked_ole_container
        if (
            shape.tag != f"{{{_VML_NS}}}shape"
            or shape.attrib != expected_shape_attributes
            or list(shape)
            or (shape.text or "").strip()
            or (shape.tail or "").strip()
            or marker is not linked_ole_marker
            or marker.attrib != expected_marker_attributes
            or list(marker)
            or (marker.text or "").strip()
            or (marker.tail or "").strip()
        ):
            _invalid(spec, f"{side} VML linked-OLE marker is invalid")


def _validate_active_x_control(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    """Validate the fixed ActiveX persistence topology without loading its binary."""

    if variant.activex_persistence_payload is None:
        return
    root = _parse_xml(members["word/activeX/activeX1.xml"], spec)
    expected_attributes = {
        f"{{{_ACTIVE_X_NS}}}classid": _ACTIVE_X_CLASS_ID,
        f"{{{_ACTIVE_X_NS}}}persistence": "persistStorage",
        f"{{{_REL_NS}}}id": "rIdActiveXBinary",
    }
    if (
        root.tag != f"{{{_ACTIVE_X_NS}}}ocx"
        or root.attrib != expected_attributes
        or list(root)
        or (root.text or "").strip()
        or (root.tail or "").strip()
    ):
        _invalid(spec, f"{side} ActiveX persistence part is invalid")
    relationships = _relationship_map(members["word/activeX/_rels/activeX1.xml.rels"], spec)
    if relationships != {
        "rIdActiveXBinary": (
            _ACTIVE_X_BINARY_RELATIONSHIP,
            "activeX1.bin",
            "internal",
        )
    }:
        _invalid(spec, f"{side} ActiveX persistence relationship is invalid")


def _validate_complex_include_text_field(
    root: ET.Element, variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    """Validate one complete, deliberately fragmented complex INCLUDETEXT field."""

    field_characters = list(root.iter(_word_tag("fldChar")))
    instruction_texts = list(root.iter(_word_tag("instrText")))
    target = variant.complex_include_text_target
    if target is None:
        if field_characters or instruction_texts:
            _invalid(spec, f"{side} unexpected complex-field markup is present")
        return
    if len(field_characters) != 3 or len(instruction_texts) != 3:
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")

    field_runs = [
        run
        for run in root.iter(_word_tag("r"))
        if any(character in run for character in field_characters)
        or any(instruction in run for instruction in instruction_texts)
    ]
    if len(field_runs) != 6:
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")
    field_paragraphs = [
        paragraph
        for paragraph in root.iter(_word_tag("p"))
        if all(run in paragraph for run in field_runs)
    ]
    if len(field_paragraphs) != 1:
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")
    paragraph_children = list(field_paragraphs[0])
    begin_index = paragraph_children.index(field_runs[0])
    if begin_index + 6 >= len(paragraph_children):
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")
    field_sequence = paragraph_children[begin_index : begin_index + 7]
    instruction_chunks = _complex_include_text_instruction_chunks(target)
    if not (
        _is_field_character_run(field_sequence[0], "begin")
        and all(
            _is_instruction_text_run(run, chunk)
            for run, chunk in zip(field_sequence[1:4], instruction_chunks, strict=True)
        )
        and _is_field_character_run(field_sequence[4], "separate")
        and _is_comment_text_run(field_sequence[5], _INCLUDE_TEXT_FIELD_RESULT)
        and _is_field_character_run(field_sequence[6], "end")
    ):
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")
    if tuple(field_runs) != tuple(field_sequence[index] for index in (0, 1, 2, 3, 4, 6)):
        _invalid(spec, f"{side} complex INCLUDETEXT field markup is invalid")


def _validate_modern_comment(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    """Validate the fixed classic-comment and Office 2013 metadata pairing."""

    if variant.modern_comment_done is None:
        return

    comments = _parse_xml(members["word/comments.xml"], spec)
    if (
        comments.tag != _word_tag("comments")
        or comments.attrib
        or len(comments) != 1
        or (comments.text or "").strip()
        or (comments.tail or "").strip()
    ):
        _invalid(spec, f"{side} classic comment root is invalid")
    comment = comments[0]
    expected_comment_attributes = {
        _word_tag("id"): _MODERN_COMMENT_ID,
        _word_tag("author"): _MODERN_COMMENT_AUTHOR,
        _word_tag("initials"): _MODERN_COMMENT_INITIALS,
        _word_tag("date"): _MODERN_COMMENT_DATE,
    }
    if (
        comment.tag != _word_tag("comment")
        or comment.attrib != expected_comment_attributes
        or len(comment) != 1
        or (comment.text or "").strip()
        or (comment.tail or "").strip()
    ):
        _invalid(spec, f"{side} classic comment is invalid")
    paragraph = comment[0]
    expected_paragraph_attributes = {
        f"{{{_WORD_2010_WORDML_NS}}}paraId": _MODERN_COMMENT_PARAGRAPH_ID,
        f"{{{_WORD_2010_WORDML_NS}}}textId": _MODERN_COMMENT_TEXT_ID,
    }
    if (
        paragraph.tag != _word_tag("p")
        or paragraph.attrib != expected_paragraph_attributes
        or len(paragraph) != 2
        or (paragraph.text or "").strip()
        or (paragraph.tail or "").strip()
    ):
        _invalid(spec, f"{side} classic comment paragraph is invalid")
    annotation_run, text_run = paragraph
    if (
        annotation_run.tag != _word_tag("r")
        or annotation_run.attrib
        or len(annotation_run) != 1
        or not _is_empty_element(annotation_run[0], _word_tag("annotationRef"), {})
        or (annotation_run.text or "").strip()
        or (annotation_run.tail or "").strip()
        or not _is_comment_text_run(text_run, _MODERN_COMMENT_TEXT)
    ):
        _invalid(spec, f"{side} classic comment content is invalid")

    comments_extended = _parse_xml(members["word/commentsExtended.xml"], spec)
    expected_done = "1" if variant.modern_comment_done else "0"
    if (
        comments_extended.tag != f"{{{_WORD_2012_WORDML_NS}}}commentsEx"
        or comments_extended.attrib
        or len(comments_extended) != 1
        or (comments_extended.text or "").strip()
        or (comments_extended.tail or "").strip()
        or not _is_empty_element(
            comments_extended[0],
            f"{{{_WORD_2012_WORDML_NS}}}commentEx",
            {
                f"{{{_WORD_2012_WORDML_NS}}}paraId": _MODERN_COMMENT_PARAGRAPH_ID,
                f"{{{_WORD_2012_WORDML_NS}}}done": expected_done,
            },
        )
    ):
        _invalid(spec, f"{side} commentsExtended metadata is invalid")


def _validate_taskpane_web_extension(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    """Validate the fixed, internal-only task-pane web-extension topology."""

    if variant.taskpane_auto_show is None:
        return

    taskpanes = _parse_xml(members["word/webextensions/taskpanes.xml"], spec)
    if (
        taskpanes.tag != f"{{{_TASKPANE_WEB_EXTENSION_TASKPANES_NS}}}taskpanes"
        or taskpanes.attrib
        or len(taskpanes) != 1
        or (taskpanes.text or "").strip()
    ):
        _invalid(spec, f"{side} task-pane web-extension root is invalid")
    taskpane = taskpanes[0]
    expected_taskpane_attributes = {
        "dockstate": "right",
        "visibility": "0",
        "width": "350",
        "row": "0",
        "locked": "false",
    }
    if (
        taskpane.tag != f"{{{_TASKPANE_WEB_EXTENSION_TASKPANES_NS}}}taskpane"
        or taskpane.attrib != expected_taskpane_attributes
        or len(taskpane) != 1
        or (taskpane.text or "").strip()
        or (taskpane.tail or "").strip()
    ):
        _invalid(spec, f"{side} task-pane web-extension markup is invalid")
    reference = taskpane[0]
    if not _is_empty_element(
        reference,
        f"{{{_TASKPANE_WEB_EXTENSION_TASKPANES_NS}}}webextensionref",
        {f"{{{_REL_NS}}}id": "rIdTaskpaneWebExtension"},
    ):
        _invalid(spec, f"{side} task-pane web-extension reference is invalid")

    if _relationship_map(members["word/webextensions/_rels/taskpanes.xml.rels"], spec) != {
        "rIdTaskpaneWebExtension": (
            _TASKPANE_WEB_EXTENSION_RELATIONSHIP,
            "webextension1.xml",
            "internal",
        )
    }:
        _invalid(spec, f"{side} task-pane web-extension relationships are invalid")

    extension = _parse_xml(members["word/webextensions/webextension1.xml"], spec)
    expected_tags = (
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}reference",
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}alternateReferences",
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}properties",
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}bindings",
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}snapshot",
    )
    if (
        extension.tag != f"{{{_TASKPANE_WEB_EXTENSION_NS}}}webextension"
        or extension.attrib != {"id": _TASKPANE_WEB_EXTENSION_ID}
        or tuple(child.tag for child in extension) != expected_tags
        or (extension.text or "").strip()
    ):
        _invalid(spec, f"{side} web-extension root is invalid")
    extension_reference, alternate_references, properties, bindings, snapshot = extension
    if not _is_empty_element(
        extension_reference,
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}reference",
        {
            "id": _TASKPANE_WEB_EXTENSION_REFERENCE_ID,
            "version": _TASKPANE_WEB_EXTENSION_REFERENCE_VERSION,
            "store": _TASKPANE_WEB_EXTENSION_REFERENCE_STORE,
            "storeType": _TASKPANE_WEB_EXTENSION_REFERENCE_STORE_TYPE,
        },
    ):
        _invalid(spec, f"{side} web-extension reference is invalid")
    if not _is_empty_element(
        alternate_references,
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}alternateReferences",
        {},
    ):
        _invalid(spec, f"{side} web-extension alternate references are invalid")
    if (
        properties.tag != f"{{{_TASKPANE_WEB_EXTENSION_NS}}}properties"
        or properties.attrib
        or len(properties) != 1
        or (properties.text or "").strip()
        or (properties.tail or "").strip()
    ):
        _invalid(spec, f"{side} web-extension properties are invalid")
    expected_auto_show_value = "true" if variant.taskpane_auto_show else "false"
    if not _is_empty_element(
        properties[0],
        f"{{{_TASKPANE_WEB_EXTENSION_NS}}}property",
        {
            "name": _TASKPANE_AUTO_SHOW_PROPERTY_NAME,
            "value": expected_auto_show_value,
        },
    ):
        _invalid(spec, f"{side} web-extension auto-show setting is invalid")
    if not _is_empty_element(bindings, f"{{{_TASKPANE_WEB_EXTENSION_NS}}}bindings", {}):
        _invalid(spec, f"{side} web-extension bindings are invalid")
    if not _is_empty_element(snapshot, f"{{{_TASKPANE_WEB_EXTENSION_NS}}}snapshot", {}):
        _invalid(spec, f"{side} web-extension snapshot is invalid")


def _validate_alternative_format_import(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    if variant.alternative_format_import_payload is None:
        return
    if members["word/afchunk1.html"] != variant.alternative_format_import_payload:
        _invalid(spec, f"{side} alternative-format import payload is invalid")


def _validate_settings(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    root = _parse_xml(members["word/settings.xml"], spec)
    if root.tag != _word_tag("settings"):
        _invalid(spec, f"{side} settings root is invalid")
    tracks = list(root.iter(_word_tag("trackRevisions")))
    protections = list(root.iter(_word_tag("documentProtection")))
    attached_templates = list(root.iter(_word_tag("attachedTemplate")))
    mail_merges = list(root.iter(_word_tag("mailMerge")))
    document_variable_containers = [
        element for element in root if element.tag == _word_tag("docVars")
    ]
    if len(tracks) != int(variant.track_revisions):
        _invalid(spec, f"{side} Track Changes setting is invalid")
    if len(protections) != int(variant.document_protection):
        _invalid(spec, f"{side} document protection setting is invalid")
    if len(attached_templates) != int(variant.attached_template_target is not None):
        _invalid(spec, f"{side} attached template setting is invalid")
    if (
        attached_templates
        and attached_templates[0].get(f"{{{_REL_NS}}}id") != "rIdAttachedTemplate"
    ):
        _invalid(spec, f"{side} attached template relationship is invalid")
    if len(mail_merges) != int(variant.mail_merge_data_source_target is not None):
        _invalid(spec, f"{side} mail-merge setting is invalid")
    if len(document_variable_containers) != int(variant.document_variable_value is not None):
        _invalid(spec, f"{side} document-variable setting is invalid")
    if document_variable_containers:
        container = document_variable_containers[0]
        if len(container) != 1 or (container.text or "").strip():
            _invalid(spec, f"{side} document-variable setting is invalid")
        document_variable = container[0]
        expected_attributes = {
            _word_tag("name"): _DOCUMENT_VARIABLE_NAME,
            _word_tag("val"): variant.document_variable_value,
        }
        if (
            document_variable.tag != _word_tag("docVar")
            or document_variable.attrib != expected_attributes
            or list(document_variable)
            or (document_variable.text or "").strip()
            or (document_variable.tail or "").strip()
        ):
            _invalid(spec, f"{side} document-variable setting is invalid")
    if mail_merges:
        mail_merge = mail_merges[0]
        main_document_types = [
            element for element in mail_merge if element.tag == _word_tag("mainDocumentType")
        ]
        data_sources = [element for element in mail_merge if element.tag == _word_tag("dataSource")]
        if (
            len(mail_merge) != 2
            or len(main_document_types) != 1
            or main_document_types[0].get(_word_tag("val")) != "formLetters"
            or len(data_sources) != 1
            or data_sources[0].get(f"{{{_REL_NS}}}id") != "rIdMailMergeSource"
        ):
            _invalid(spec, f"{side} mail-merge data-source anchor is invalid")
    if protections:
        protection = protections[0]
        if (
            protection.get(_word_tag("edit")) != "readOnly"
            or protection.get(_word_tag("enforcement")) != "1"
            or any(
                protection.get(_word_tag(attribute)) is not None
                for attribute in ("hash", "hashValue", "salt", "saltValue")
            )
        ):
            _invalid(spec, f"{side} document protection setting is invalid")


def _validate_settings_relationships(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    has_relationships = "word/_rels/settings.xml.rels" in members
    if variant.attached_template_target is None and variant.mail_merge_data_source_target is None:
        if has_relationships:
            _invalid(spec, f"{side} unexpected settings relationships are present")
        return
    if not has_relationships:
        _invalid(spec, f"{side} settings relationships are absent")
    relationships = _relationship_map(members["word/_rels/settings.xml.rels"], spec)
    expected = {}
    if variant.attached_template_target is not None:
        expected["rIdAttachedTemplate"] = (
            _ATTACHED_TEMPLATE_RELATIONSHIP,
            variant.attached_template_target,
            "external",
        )
    if variant.mail_merge_data_source_target is not None:
        expected["rIdMailMergeSource"] = (
            _MAIL_MERGE_SOURCE_RELATIONSHIP,
            variant.mail_merge_data_source_target,
            "external",
        )
    if relationships != expected:
        _invalid(spec, f"{side} settings relationships are invalid")


def _validate_web_settings(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    """Validate the complete, source-backed root frameset topology."""

    if variant.frameset_source_target is None:
        return
    web_settings = _parse_xml(members["word/webSettings.xml"], spec)
    if (
        web_settings.tag != _word_tag("webSettings")
        or web_settings.attrib
        or len(web_settings) != 1
        or (web_settings.text or "").strip()
        or (web_settings.tail or "").strip()
    ):
        _invalid(spec, f"{side} Web Settings root is invalid")
    frameset = web_settings[0]
    if (
        frameset.tag != _word_tag("frameset")
        or frameset.attrib
        or len(frameset) != 2
        or (frameset.text or "").strip()
        or (frameset.tail or "").strip()
    ):
        _invalid(spec, f"{side} frameset markup is invalid")
    layout, frame = frameset
    if not _is_empty_element(
        layout,
        _word_tag("frameLayout"),
        {_word_tag("val"): _FRAME_LAYOUT},
    ):
        _invalid(spec, f"{side} frameset layout is invalid")
    if (
        frame.tag != _word_tag("frame")
        or frame.attrib
        or len(frame) != 3
        or (frame.text or "").strip()
        or (frame.tail or "").strip()
    ):
        _invalid(spec, f"{side} frameset frame is invalid")
    size, name, source = frame
    if not (
        _is_empty_element(size, _word_tag("sz"), {_word_tag("val"): _FRAME_SIZE})
        and _is_empty_element(name, _word_tag("name"), {_word_tag("val"): _FRAME_NAME})
        and _is_empty_element(
            source,
            _word_tag("sourceFileName"),
            {f"{{{_REL_NS}}}id": "rIdFrameSource"},
        )
    ):
        _invalid(spec, f"{side} frameset source anchor is invalid")
    relationships = _relationship_map(members["word/_rels/webSettings.xml.rels"], spec)
    expected = {
        "rIdFrameSource": (
            _FRAME_RELATIONSHIP,
            variant.frameset_source_target,
            "external",
        )
    }
    if relationships != expected:
        _invalid(spec, f"{side} frameset source relationship is invalid")


def _validate_styles(members: dict[str, bytes], spec: CaseSpec, side: str) -> None:
    root = _parse_xml(members["word/styles.xml"], spec)
    if root.tag != _word_tag("styles"):
        _invalid(spec, f"{side} styles root is invalid")


def _validate_custom_xml(
    members: dict[str, bytes], variant: DocumentVariant, spec: CaseSpec, side: str
) -> None:
    if variant.custom_xml_payload is None:
        return
    if members["customXml/item1.xml"] != variant.custom_xml_payload:
        _invalid(spec, f"{side} custom XML payload is invalid")
    _parse_xml(members["customXml/item1.xml"], spec)
    properties = _parse_xml(members["customXml/itemProps1.xml"], spec)
    if properties.tag != f"{{{_CUSTOM_XML_PROPERTIES_NS}}}datastoreItem":
        _invalid(spec, f"{side} custom XML properties are invalid")
    relationships = _relationship_map(members["customXml/_rels/item1.xml.rels"], spec)
    expected = {
        "rIdItemProperties": (_CUSTOM_XML_PROPERTIES_RELATIONSHIP, "itemProps1.xml", "internal")
    }
    if relationships != expected:
        _invalid(spec, f"{side} custom XML relationships are invalid")


def _stored_texts(members: dict[str, bytes], spec: CaseSpec) -> tuple[str, ...]:
    root = _parse_xml(members["word/document.xml"], spec)
    return tuple(element.text or "" for element in root.iter(_word_tag("t")))


def _relationship_map(data: bytes, spec: CaseSpec) -> dict[str, tuple[str, str, str]]:
    root = _parse_xml(data, spec)
    if root.tag != f"{{{_PACKAGE_REL_NS}}}Relationships":
        _invalid(spec, "relationships root is invalid")
    records: dict[str, tuple[str, str, str]] = {}
    for child in root:
        if child.tag != f"{{{_PACKAGE_REL_NS}}}Relationship":
            _invalid(spec, "relationship is invalid")
        relationship_id = child.get("Id")
        relationship_type = child.get("Type")
        target = child.get("Target")
        target_mode = child.get("TargetMode", "Internal").casefold()
        if (
            not relationship_id
            or not relationship_type
            or target is None
            or relationship_id in records
            or target_mode not in {"internal", "external"}
        ):
            _invalid(spec, "relationship is incomplete")
        records[relationship_id] = (relationship_type, target, target_mode)
    return records


def _read_archive(data: bytes, spec: CaseSpec, side: str) -> dict[str, bytes]:
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            if archive.testzip() is not None:
                _invalid(spec, f"{side} archive failed integrity check")
            infos = archive.infolist()
            names = [info.filename for info in infos]
            if len(names) != len(set(names)) or any(
                not name or name.startswith("/") or "\\" in name or ".." in Path(name).parts
                for name in names
            ):
                _invalid(spec, f"{side} archive member names are invalid")
            return {info.filename: archive.read(info) for info in infos}
    except (OSError, RuntimeError, zipfile.BadZipFile):
        raise FixtureValidationError(f"{spec.case_id}: {side} package cannot be read") from None


def _validate_manifest(path: Path) -> None:
    if not path.is_file() or path.is_symlink():
        raise FixtureValidationError("manifest is invalid")
    try:
        rows = [json.loads(line) for line in path.read_text(encoding="utf-8").splitlines()]
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        raise FixtureValidationError("manifest cannot be read") from None
    expected = [
        {
            "id": spec.case_id,
            "schema_version": FIXTURE_SCHEMA_VERSION,
            "truth": f"{spec.case_id}/truth.json",
        }
        for spec in CASE_SPECS
    ]
    if rows != expected:
        raise FixtureValidationError("manifest does not match the DCAB catalogue")


def _validate_public_truth(truth: dict[str, Any], spec: CaseSpec) -> None:
    encoded = json.dumps(truth, sort_keys=True)
    forbidden = (
        "example.invalid",
        "rIdHyperlink",
        "rIdAttachedTemplate",
        "rIdMailMergeSource",
        "rIdSubDocument",
        "rIdWebSettings",
        "rIdFrameSource",
        "rIdLinkedPicture",
        "rIdAltChunk",
        "rIdVbaProject",
        "rIdOleObject",
        "rIdActiveXControl",
        "rIdActiveXBinary",
        "vbaProject.bin",
        "oleObject1.bin",
        "activeX1.xml",
        "activeX1.bin",
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
        "rIdTaskpaneWebExtensions",
        "rIdTaskpaneWebExtension",
        "webextension1.xml",
        "webSettings.xml",
        "rIdComments",
        "rIdCommentsExtended",
        "comments.xml",
        "commentsExtended.xml",
        _TASKPANE_WEB_EXTENSION_ID,
        _TASKPANE_WEB_EXTENSION_REFERENCE_ID,
        _TASKPANE_WEB_EXTENSION_REFERENCE_STORE,
        _TASKPANE_AUTO_SHOW_PROPERTY_NAME,
        _VML_SHAPE_ID,
        _VML_SHAPE_TARGET_FRAME,
        _FRAME_LAYOUT,
        _FRAME_NAME,
        _FRAME_SIZE,
        _MODERN_COMMENT_PARAGRAPH_ID,
        _MODERN_COMMENT_TEXT_ID,
        _MODERN_COMMENT_AUTHOR,
        _MODERN_COMMENT_INITIALS,
        _MODERN_COMMENT_DATE,
        _MODERN_COMMENT_ANCHOR_TEXT,
        _MODERN_COMMENT_TEXT,
        _ACTIVE_X_CONTROL_NAME,
        _ACTIVE_X_CLASS_ID,
    )
    if any(value in encoded for value in forbidden):
        _invalid(spec, "public truth contains private fixture material")


def _parse_xml(data: bytes, spec: CaseSpec) -> ET.Element:
    if b"<!DOCTYPE" in data.upper() or b"<!ENTITY" in data.upper():
        _invalid(spec, "XML declarations are unsupported")
    try:
        return ET.fromstring(data)
    except (ET.ParseError, UnicodeError, ValueError):
        _invalid(spec, "XML cannot be parsed")
    raise AssertionError("unreachable")


def _is_empty_element(element: ET.Element, tag: str, attributes: dict[str, str]) -> bool:
    return (
        element.tag == tag
        and element.attrib == attributes
        and not list(element)
        and not (element.text or "").strip()
        and not (element.tail or "").strip()
    )


def _is_comment_text_run(run: ET.Element, text: str) -> bool:
    """Return whether ``run`` has exactly one unstyled Word text child."""

    return (
        run.tag == _word_tag("r")
        and not run.attrib
        and len(run) == 1
        and run[0].tag == _word_tag("t")
        and not run[0].attrib
        and run[0].text == text
        and not list(run[0])
        and not (run[0].tail or "").strip()
        and not (run.text or "").strip()
        and not (run.tail or "").strip()
    )


def _is_field_character_run(run: ET.Element, field_type: str) -> bool:
    """Return whether ``run`` has one unadorned complex-field marker."""

    return (
        run.tag == _word_tag("r")
        and not run.attrib
        and len(run) == 1
        and _is_empty_element(
            run[0],
            _word_tag("fldChar"),
            {_word_tag("fldCharType"): field_type},
        )
        and not (run.text or "").strip()
        and not (run.tail or "").strip()
    )


def _is_instruction_text_run(run: ET.Element, text: str) -> bool:
    """Return whether ``run`` has one preserved-whitespace field-code fragment."""

    return (
        run.tag == _word_tag("r")
        and not run.attrib
        and len(run) == 1
        and run[0].tag == _word_tag("instrText")
        and run[0].attrib == {f"{{{_XML_NS}}}space": "preserve"}
        and run[0].text == text
        and not list(run[0])
        and not (run[0].tail or "").strip()
        and not (run.text or "").strip()
        and not (run.tail or "").strip()
    )


def _run_has_vanish(run: ET.Element) -> bool:
    properties = run.find(_word_tag("rPr"))
    return properties is not None and properties.find(_word_tag("vanish")) is not None


def _word_tag(local_name: str) -> str:
    return f"{{{_WORD_NS}}}{local_name}"


def _drawing_tag(local_name: str) -> str:
    return f"{{{_DRAWING_NS}}}{local_name}"


def _wordprocessing_drawing_tag(local_name: str) -> str:
    return f"{{{_WORDPROCESSING_DRAWING_NS}}}{local_name}"


def _picture_tag(local_name: str) -> str:
    return f"{{{_PICTURE_NS}}}{local_name}"


def _invalid(spec: CaseSpec, message: str) -> None:
    raise FixtureValidationError(f"{spec.case_id}: {message}")
