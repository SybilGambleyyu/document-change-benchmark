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
    _ALT_CHUNK_CONTENT_TYPE,
    _ALT_CHUNK_RELATIONSHIP,
    _ATTACHED_TEMPLATE_RELATIONSHIP,
    _CONTENT_TYPES_NS,
    _CUSTOM_XML_PROPERTIES_CONTENT_TYPE,
    _CUSTOM_XML_PROPERTIES_NS,
    _CUSTOM_XML_PROPERTIES_RELATIONSHIP,
    _CUSTOM_XML_RELATIONSHIP,
    _DOCM_MAIN_CONTENT_TYPE,
    _DOCX_MAIN_CONTENT_TYPE,
    _DRAWING_NS,
    _HYPERLINK_RELATIONSHIP,
    _IMAGE_RELATIONSHIP,
    _OFFICE_DOCUMENT_RELATIONSHIP,
    _OFFICE_VML_NS,
    _OLE_CONTENT_TYPE,
    _OLE_OBJECT_RELATIONSHIP,
    _PACKAGE_REL_NS,
    _PICTURE_NS,
    _REL_NS,
    _SETTINGS_CONTENT_TYPE,
    _SETTINGS_RELATIONSHIP,
    _STYLES_CONTENT_TYPE,
    _STYLES_RELATIONSHIP,
    _SUBDOCUMENT_RELATIONSHIP,
    _VBA_PROJECT_CONTENT_TYPE,
    _VBA_PROJECT_RELATIONSHIP,
    _VISIBLE_TEXT,
    _WORD_NS,
    _WORDPROCESSING_DRAWING_NS,
    CASE_IDS,
    CASE_SPECS,
    FIXTURE_SCHEMA_VERSION,
    CaseSpec,
    DocumentVariant,
    case_files,
    truth_manifest,
)


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
    if variant.alternative_format_import_payload is not None:
        expected_members.add("word/afchunk1.html")
    if variant.attached_template_target is not None:
        expected_members.add("word/_rels/settings.xml.rels")
    if set(members) != expected_members:
        _invalid(spec, f"{side} package members are invalid")

    _validate_content_types(members, variant, spec, side)
    _validate_root_relationships(members, spec, side)
    _validate_document_relationships(members, variant, spec, side)
    _validate_document_xml(members, variant, spec, side)
    _validate_alternative_format_import(members, variant, spec, side)
    _validate_settings(members, variant, spec, side)
    _validate_settings_relationships(members, variant, spec, side)
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
    expected_fields: list[tuple[str, str]] = []
    if variant.hyperlink_field_target is not None:
        expected_fields.append(("HYPERLINK", variant.hyperlink_field_target))
    if variant.include_text_target is not None:
        expected_fields.append(("INCLUDETEXT", variant.include_text_target))
    if len(fields) != len(expected_fields):
        _invalid(spec, f"{side} field markup is invalid")
    for field, (kind, target) in zip(fields, expected_fields, strict=True):
        instruction = field.get(_word_tag("instr"))
        expected_instruction = f' {kind} "{target}" '
        if instruction != expected_instruction:
            _invalid(spec, f"{side} field instruction is invalid")

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
    if variant.attached_template_target is None:
        if has_relationships:
            _invalid(spec, f"{side} unexpected settings relationships are present")
        return
    if not has_relationships:
        _invalid(spec, f"{side} attached template relationships are absent")
    relationships = _relationship_map(members["word/_rels/settings.xml.rels"], spec)
    expected = {
        "rIdAttachedTemplate": (
            _ATTACHED_TEMPLATE_RELATIONSHIP,
            variant.attached_template_target,
            "external",
        )
    }
    if relationships != expected:
        _invalid(spec, f"{side} attached template relationships are invalid")


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
