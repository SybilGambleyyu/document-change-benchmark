"""Deterministic source for DCAB's synthetic WordprocessingML fixture pairs.

DCAB builds compact OPC packages rather than redistributing real documents.
Stored URI-like values use the reserved ``example.invalid`` domain, and VBA,
OLE, and ActiveX persistence payloads are inert marker bytes. Building a
fixture never opens a Word client, resolves a relationship, updates a field,
parses an opaque payload, or executes stored code.
"""

from __future__ import annotations

import html
import io
import json
import stat
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Any

FIXTURE_SCHEMA_VERSION = 1

_CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
_PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_WORD_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_WORD_2010_WORDML_NS = "http://schemas.microsoft.com/office/word/2010/wordml"
_WORD_2012_WORDML_NS = "http://schemas.microsoft.com/office/word/2012/wordml"
_DRAWING_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
_WORDPROCESSING_DRAWING_NS = (
    "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
)
_PICTURE_NS = "http://schemas.openxmlformats.org/drawingml/2006/picture"
_OFFICE_VML_NS = "urn:schemas-microsoft-com:office:office"
_VML_NS = "urn:schemas-microsoft-com:vml"
_ACTIVE_X_NS = "http://schemas.microsoft.com/office/2006/activeX"
_CUSTOM_XML_PROPERTIES_NS = "http://schemas.openxmlformats.org/officeDocument/2006/customXml"
_TASKPANE_WEB_EXTENSION_TASKPANES_NS = (
    "http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11"
)
_TASKPANE_WEB_EXTENSION_NS = (
    "http://schemas.microsoft.com/office/webextensions/webextension/2010/11"
)

_DOCX_MAIN_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
)
_DOCM_MAIN_CONTENT_TYPE = "application/vnd.ms-word.document.macroEnabled.main+xml"
_STYLES_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"
_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"
)
_WEB_SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.webSettings+xml"
)
_COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
_COMMENTS_EXTENDED_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"
)
_CUSTOM_XML_PROPERTIES_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"
)
_ALT_CHUNK_CONTENT_TYPE = "text/html"
_VBA_PROJECT_CONTENT_TYPE = "application/vnd.ms-office.vbaProject"
_OLE_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.oleObject"
_ACTIVE_X_CONTENT_TYPE = "application/vnd.ms-office.activeX+xml"
_ACTIVE_X_BINARY_CONTENT_TYPE = "application/vnd.ms-office.activeX"
_TASKPANE_WEB_EXTENSION_TASKPANES_CONTENT_TYPE = (
    "application/vnd.ms-office.webextensiontaskpanes+xml"
)
_TASKPANE_WEB_EXTENSION_CONTENT_TYPE = "application/vnd.ms-office.webextension+xml"

_OFFICE_DOCUMENT_RELATIONSHIP = f"{_REL_NS}/officeDocument"
_HYPERLINK_RELATIONSHIP = f"{_REL_NS}/hyperlink"
_STYLES_RELATIONSHIP = f"{_REL_NS}/styles"
_SETTINGS_RELATIONSHIP = f"{_REL_NS}/settings"
_WEB_SETTINGS_RELATIONSHIP = f"{_REL_NS}/webSettings"
_COMMENTS_RELATIONSHIP = f"{_REL_NS}/comments"
_COMMENTS_EXTENDED_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended"
)
_ATTACHED_TEMPLATE_RELATIONSHIP = f"{_REL_NS}/attachedTemplate"
_MAIL_MERGE_SOURCE_RELATIONSHIP = f"{_REL_NS}/mailMergeSource"
_SUBDOCUMENT_RELATIONSHIP = f"{_REL_NS}/subDocument"
_FRAME_RELATIONSHIP = f"{_REL_NS}/frame"
_IMAGE_RELATIONSHIP = f"{_REL_NS}/image"
_ALT_CHUNK_RELATIONSHIP = f"{_REL_NS}/afChunk"
_CUSTOM_XML_RELATIONSHIP = f"{_REL_NS}/customXml"
_CUSTOM_XML_PROPERTIES_RELATIONSHIP = f"{_REL_NS}/customXmlProps"
_VBA_PROJECT_RELATIONSHIP = "http://schemas.microsoft.com/office/2006/relationships/vbaProject"
_OLE_OBJECT_RELATIONSHIP = f"{_REL_NS}/oleObject"
_CONTROL_RELATIONSHIP = f"{_REL_NS}/control"
_ACTIVE_X_BINARY_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2006/relationships/activeXControlBinary"
)
_TASKPANE_WEB_EXTENSION_TASKPANES_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2011/relationships/webextensiontaskpanes"
)
_TASKPANE_WEB_EXTENSION_RELATIONSHIP = (
    "http://schemas.microsoft.com/office/2011/relationships/webextension"
)

_VISIBLE_TEXT = "DCAB synthetic static-review fixture"
_HYPERLINK_DISPLAY_TEXT = "DCAB hyperlink display text"
_HYPERLINK_FIELD_RESULT = "DCAB hyperlink field result"
_INCLUDE_TEXT_FIELD_RESULT = "DCAB include-text field result"
_DDE_FIELD_RESULT = "DCAB DDE field result"
_DOCUMENT_VARIABLE_FIELD_RESULT = "DCAB document-variable field result"
_PERMISSION_RANGE_TEXT = "DCAB editable-range carrier"
_HIDDEN_TEXT = "DCAB hidden-text carrier"
_REVISION_TEXT = "DCAB revision carrier"
_BINDING_TEXT = "DCAB bound-content carrier"
_CUSTOM_XML_STORE_ID = "{3B1F8916-697D-4C4E-A4A1-55F64D9F2A80}"

_DIRECT_LINK_APPROVED = "https://approved.example.invalid/dcab-hyperlink"
_DIRECT_LINK_CANDIDATE = "https://candidate.example.invalid/dcab-hyperlink"
_VML_SHAPE_HYPERLINK_APPROVED = "https://approved.example.invalid/dcab-vml-shape"
_VML_SHAPE_HYPERLINK_CANDIDATE = "https://candidate.example.invalid/dcab-vml-shape"
_VML_SHAPE_ID = "DCABVmlLinkShape"
_VML_SHAPE_TARGET_FRAME = "_blank"
_FIELD_LINK_APPROVED = "https://approved.example.invalid/dcab-field-hyperlink"
_FIELD_LINK_CANDIDATE = "https://candidate.example.invalid/dcab-field-hyperlink"
_INCLUDE_TEXT_APPROVED = "https://approved.example.invalid/dcab-source.docx"
_INCLUDE_TEXT_CANDIDATE = "https://candidate.example.invalid/dcab-source.docx"
_DDE_APPLICATION = "DCAB"
_DDE_ITEM = "Sheet1!R1C1"
_DDE_SOURCE_APPROVED = "C:\\DCAB\\approved-source.xlsx"
_DDE_SOURCE_CANDIDATE = "C:\\DCAB\\candidate-source.xlsx"
_DOCUMENT_VARIABLE_NAME = "DCABReviewState"
_DOCUMENT_VARIABLE_APPROVED = "approved-state"
_DOCUMENT_VARIABLE_CANDIDATE = "candidate-state"
_PERMISSION_RANGE_MARKER_ID = "0"
_PERMISSION_RANGE_EDITOR_APPROVED = "DCAB_EDITOR_BASELINE"
_PERMISSION_RANGE_EDITOR_CANDIDATE = "DCAB_EDITOR_CANDIDATE"
_TASKPANE_WEB_EXTENSION_ID = "{3F08C2A1-681F-451E-95B6-001122334455}"
_TASKPANE_WEB_EXTENSION_REFERENCE_ID = "{F928A11C-9164-4F8A-8D92-556677889900}"
_TASKPANE_WEB_EXTENSION_REFERENCE_VERSION = "1.0.0.0"
_TASKPANE_WEB_EXTENSION_REFERENCE_STORE = "EXCatalog"
_TASKPANE_WEB_EXTENSION_REFERENCE_STORE_TYPE = "EXCatalog"
_TASKPANE_AUTO_SHOW_PROPERTY_NAME = "Office.AutoShowTaskpaneWithDocument"
_MODERN_COMMENT_ID = "0"
_MODERN_COMMENT_PARAGRAPH_ID = "0A0B0C0D"
_MODERN_COMMENT_TEXT_ID = "77777777"
_MODERN_COMMENT_AUTHOR = "DCAB-FIXTURE-COMMENT-AUTHOR"
_MODERN_COMMENT_INITIALS = "DCF"
_MODERN_COMMENT_DATE = "2026-08-03T00:00:00Z"
_MODERN_COMMENT_ANCHOR_TEXT = "DCAB comment anchor carrier"
_MODERN_COMMENT_TEXT = "DCAB fixed review comment"
_ATTACHED_TEMPLATE_APPROVED = "https://approved.example.invalid/dcab-template.dotx"
_ATTACHED_TEMPLATE_CANDIDATE = "https://candidate.example.invalid/dcab-template.dotx"
_MAIL_MERGE_SOURCE_APPROVED = "https://approved.example.invalid/dcab-mail-merge.csv"
_MAIL_MERGE_SOURCE_CANDIDATE = "https://candidate.example.invalid/dcab-mail-merge.csv"
_SUBDOCUMENT_APPROVED = "https://approved.example.invalid/dcab-subdocument.docx"
_SUBDOCUMENT_CANDIDATE = "https://candidate.example.invalid/dcab-subdocument.docx"
_FRAME_SOURCE_APPROVED = "https://approved.example.invalid/dcab-frame.docx"
_FRAME_SOURCE_CANDIDATE = "https://candidate.example.invalid/dcab-frame.docx"
_FRAME_LAYOUT = "rows"
_FRAME_NAME = "DCAB frame carrier"
_FRAME_SIZE = "216"
_LINKED_OLE_APPROVED = "https://approved.example.invalid/dcab-linked-ole.xlsx"
_LINKED_OLE_CANDIDATE = "https://candidate.example.invalid/dcab-linked-ole.xlsx"
_LINKED_OLE_SHAPE_ID = "DCABLinkedOleShape"
_LINKED_OLE_PROG_ID = "DCAB.Synthetic"
_LINKED_OLE_OBJECT_ID = "DCABLinkedOleObject"
_LINKED_OLE_UPDATE_MODE = "OnCall"
_LINKED_PICTURE_APPROVED = "https://approved.example.invalid/dcab-linked-picture.png"
_LINKED_PICTURE_CANDIDATE = "https://candidate.example.invalid/dcab-linked-picture.png"
_ALT_CHUNK_APPROVED = b"<html><body>DCAB synthetic alternate-content marker: approved</body></html>"
_ALT_CHUNK_CANDIDATE = (
    b"<html><body>DCAB synthetic alternate-content marker: candidate</body></html>"
)
_BINDING_XPATH_APPROVED = "/dcab:fixture/dcab:approved"
_BINDING_XPATH_CANDIDATE = "/dcab:fixture/dcab:candidate"
_CUSTOM_XML_APPROVED = (
    b'<?xml version="1.0" encoding="UTF-8"?><dcab:fixture xmlns:dcab="urn:dcab:fixture">'
    b"<dcab:value>approved</dcab:value></dcab:fixture>"
)
_CUSTOM_XML_CANDIDATE = (
    b'<?xml version="1.0" encoding="UTF-8"?><dcab:fixture xmlns:dcab="urn:dcab:fixture">'
    b"<dcab:value>candidate</dcab:value></dcab:fixture>"
)
_OPAQUE_MACRO_APPROVED = b"DCAB inert synthetic VBA marker payload: approved\n"
_OPAQUE_MACRO_CANDIDATE = b"DCAB inert synthetic VBA marker payload: candidate\n"
_OPAQUE_OLE_APPROVED = b"DCAB inert synthetic OLE marker payload: approved\n"
_OPAQUE_OLE_CANDIDATE = b"DCAB inert synthetic OLE marker payload: candidate\n"
_ACTIVE_X_CONTROL_NAME = "DCABActiveXControl"
_ACTIVE_X_CLASS_ID = "{11111111-2222-3333-4444-555555555555}"
_OPAQUE_ACTIVE_X_APPROVED = b"DCAB inert synthetic ActiveX marker payload: approved\n"
_OPAQUE_ACTIVE_X_CANDIDATE = b"DCAB inert synthetic ActiveX marker payload: candidate\n"


class FixtureBuildError(ValueError):
    """The fixture destination or requested package state is unsafe or invalid."""


@dataclass(frozen=True)
class DocumentVariant:
    """The package state required for one side of one DCAB pair."""

    extension: str = "docx"
    direct_hyperlink_target: str | None = None
    vml_shape_hyperlink_target: str | None = None
    hyperlink_field_target: str | None = None
    include_text_target: str | None = None
    complex_include_text_target: str | None = None
    dde_source_file: str | None = None
    document_variable_value: str | None = None
    permission_range_editor: str | None = None
    taskpane_auto_show: bool | None = None
    modern_comment_done: bool | None = None
    attached_template_target: str | None = None
    mail_merge_data_source_target: str | None = None
    subdocument_target: str | None = None
    frameset_source_target: str | None = None
    vml_linked_ole_target: str | None = None
    drawing_linked_picture_target: str | None = None
    alternative_format_import_payload: bytes | None = None
    hidden_text: bool = False
    insertion_markup: bool = False
    track_revisions: bool = False
    document_protection: bool = False
    data_binding_xpath: str | None = None
    custom_xml_payload: bytes | None = None
    macro_payload: bytes | None = None
    embedded_payload: bytes | None = None
    activex_persistence_payload: bytes | None = None


@dataclass(frozen=True)
class CaseSpec:
    """Source, public assertion, and reference convention for one fixture pair."""

    case_id: str
    title: str
    description: str
    fact: dict[str, str]
    review_expectation: str
    baseline: DocumentVariant
    candidate: DocumentVariant
    changed_members: tuple[str, ...]

    @property
    def baseline_name(self) -> str:
        return f"baseline.{self.baseline.extension}"

    @property
    def candidate_name(self) -> str:
        return f"candidate.{self.candidate.extension}"


CASE_SPECS: tuple[CaseSpec, ...] = (
    CaseSpec(
        case_id="interaction.word_hyperlink_target_retargeted",
        title="Direct Word hyperlink target retargeted",
        description=(
            "A relationship-backed w:hyperlink retains its display text and stored markup "
            "shape while its external relationship target changes."
        ),
        fact={
            "binding": "external",
            "kind": "word_hyperlink_target_changed",
            "relationship_category": "hyperlink",
            "source": "word_hyperlink",
        },
        review_expectation="block",
        baseline=DocumentVariant(direct_hyperlink_target=_DIRECT_LINK_APPROVED),
        candidate=DocumentVariant(direct_hyperlink_target=_DIRECT_LINK_CANDIDATE),
        changed_members=("word/_rels/document.xml.rels",),
    ),
    CaseSpec(
        case_id="interaction.word_hyperlink_added",
        title="Direct Word hyperlink added",
        description=(
            "A normal run is wrapped in external relationship-backed w:hyperlink markup "
            "without changing its stored display text."
        ),
        fact={
            "binding": "external",
            "kind": "word_hyperlink_added",
            "relationship_category": "hyperlink",
            "source": "word_hyperlink",
        },
        review_expectation="block",
        baseline=DocumentVariant(),
        candidate=DocumentVariant(direct_hyperlink_target=_DIRECT_LINK_CANDIDATE),
        changed_members=("word/_rels/document.xml.rels", "word/document.xml"),
    ),
    CaseSpec(
        case_id="interaction.vml_shape_hyperlink_target_retargeted",
        title="VML shape hyperlink target retargeted",
        description=(
            "A legacy VML rectangle retains its fixed shape and target frame while its "
            "direct hyperlink target changes."
        ),
        fact={
            "kind": "vml_shape_hyperlink_target_changed",
            "source": "word_vml",
        },
        review_expectation="block",
        baseline=DocumentVariant(vml_shape_hyperlink_target=_VML_SHAPE_HYPERLINK_APPROVED),
        candidate=DocumentVariant(vml_shape_hyperlink_target=_VML_SHAPE_HYPERLINK_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="interaction.word_hyperlink_field_target_retargeted",
        title="HYPERLINK field target retargeted",
        description=(
            "A complete simple HYPERLINK field retains its stored result text while its "
            "private field instruction destination changes."
        ),
        fact={
            "field_kind": "hyperlink",
            "kind": "field_target_changed",
            "source": "word_field",
        },
        review_expectation="block",
        baseline=DocumentVariant(hyperlink_field_target=_FIELD_LINK_APPROVED),
        candidate=DocumentVariant(hyperlink_field_target=_FIELD_LINK_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="interaction.taskpane_auto_show_setting_enabled",
        title="Task-pane auto-show setting enabled",
        description=(
            "An internally linked Office web-extension task pane retains its parts, "
            "reference, and property shape while its stored auto-show setting changes "
            "from false to true."
        ),
        fact={
            "kind": "taskpane_auto_show_setting_enabled",
            "source": "office_web_extension",
        },
        review_expectation="review",
        baseline=DocumentVariant(taskpane_auto_show=False),
        candidate=DocumentVariant(taskpane_auto_show=True),
        changed_members=("word/webextensions/webextension1.xml",),
    ),
    CaseSpec(
        case_id="review.modern_comment_done_state_changed",
        title="Office 2013 extended-comment done state changed",
        description=(
            "An anchored classic comment and its Office 2013 commentsExtended metadata "
            "retain their fixed shape while the stored done state changes from false to true."
        ),
        fact={
            "kind": "modern_comment_done_state_changed",
            "source": "word_comments_extended",
        },
        review_expectation="review",
        baseline=DocumentVariant(modern_comment_done=False),
        candidate=DocumentVariant(modern_comment_done=True),
        changed_members=("word/commentsExtended.xml",),
    ),
    CaseSpec(
        case_id="external.include_text_field_target_retargeted",
        title="INCLUDETEXT field source retargeted",
        description=(
            "A complete simple INCLUDETEXT field retains its stored result text while its "
            "private external source instruction changes."
        ),
        fact={
            "field_kind": "include_text",
            "kind": "external_field_source_changed",
            "source": "word_field",
        },
        review_expectation="block",
        baseline=DocumentVariant(include_text_target=_INCLUDE_TEXT_APPROVED),
        candidate=DocumentVariant(include_text_target=_INCLUDE_TEXT_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="external.complex_include_text_field_target_retargeted",
        title="Fragmented complex INCLUDETEXT source retargeted",
        description=(
            "A complete complex INCLUDETEXT field retains its begin/separate/end markers, "
            "fragmented instruction shape, and stored result text while its private external "
            "source instruction changes."
        ),
        fact={
            "field_encoding": "complex_fragmented",
            "field_kind": "include_text",
            "kind": "external_field_source_changed",
            "source": "word_field",
        },
        review_expectation="block",
        baseline=DocumentVariant(complex_include_text_target=_INCLUDE_TEXT_APPROVED),
        candidate=DocumentVariant(complex_include_text_target=_INCLUDE_TEXT_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="external.dde_field_source_retargeted",
        title="DDE field source retargeted",
        description=(
            "A complete simple DDE field retains its stored result, application token, and "
            "item token while its private source-file argument changes."
        ),
        fact={
            "field_kind": "dde",
            "kind": "external_field_source_changed",
            "source": "word_field",
        },
        review_expectation="block",
        baseline=DocumentVariant(dde_source_file=_DDE_SOURCE_APPROVED),
        candidate=DocumentVariant(dde_source_file=_DDE_SOURCE_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="binding.document_variable_value_changed",
        title="Document-variable value changed",
        description=(
            "A fixed DOCVARIABLE field reference retains its stored result text and private "
            "variable name while the value of its settings-stored document variable changes."
        ),
        fact={
            "binding": "document_variable",
            "kind": "document_variable_value_changed",
            "source": "word_settings",
        },
        review_expectation="review",
        baseline=DocumentVariant(document_variable_value=_DOCUMENT_VARIABLE_APPROVED),
        candidate=DocumentVariant(document_variable_value=_DOCUMENT_VARIABLE_CANDIDATE),
        changed_members=("word/settings.xml",),
    ),
    CaseSpec(
        case_id="external.attached_template_target_retargeted",
        title="Attached template target retargeted",
        description=(
            "A w:attachedTemplate settings anchor retains its relationship ID while the "
            "external attached-template relationship target changes."
        ),
        fact={
            "binding": "external",
            "dependency": "attached_template",
            "kind": "external_document_dependency_target_changed",
            "source": "word_settings",
        },
        review_expectation="block",
        baseline=DocumentVariant(attached_template_target=_ATTACHED_TEMPLATE_APPROVED),
        candidate=DocumentVariant(attached_template_target=_ATTACHED_TEMPLATE_CANDIDATE),
        changed_members=("word/_rels/settings.xml.rels",),
    ),
    CaseSpec(
        case_id="external.mail_merge_data_source_target_retargeted",
        title="Mail-merge data-source target retargeted",
        description=(
            "A w:mailMerge w:dataSource settings anchor retains its relationship ID while the "
            "external mail-merge data-source relationship target changes."
        ),
        fact={
            "binding": "external",
            "kind": "mail_merge_data_source_target_changed",
            "relationship_category": "mail_merge_source",
            "source": "word_settings",
        },
        review_expectation="block",
        baseline=DocumentVariant(mail_merge_data_source_target=_MAIL_MERGE_SOURCE_APPROVED),
        candidate=DocumentVariant(mail_merge_data_source_target=_MAIL_MERGE_SOURCE_CANDIDATE),
        changed_members=("word/_rels/settings.xml.rels",),
    ),
    CaseSpec(
        case_id="external.subdocument_target_retargeted",
        title="Master-document subdocument target retargeted",
        description=(
            "A w:subDoc master-document anchor retains its relationship ID while the "
            "external subdocument relationship target changes."
        ),
        fact={
            "binding": "external",
            "dependency": "subdocument",
            "kind": "external_document_dependency_target_changed",
            "source": "word_document",
        },
        review_expectation="block",
        baseline=DocumentVariant(subdocument_target=_SUBDOCUMENT_APPROVED),
        candidate=DocumentVariant(subdocument_target=_SUBDOCUMENT_CANDIDATE),
        changed_members=("word/_rels/document.xml.rels",),
    ),
    CaseSpec(
        case_id="external.frameset_source_target_retargeted",
        title="Frameset source target retargeted",
        description=(
            "A complete Web Settings frameset retains its layout, frame anchor, and "
            "relationship ID while its required external frame-source relationship target changes."
        ),
        fact={
            "binding": "external",
            "dependency": "frameset_source",
            "kind": "external_document_dependency_target_changed",
            "source": "word_web_settings",
        },
        review_expectation="block",
        baseline=DocumentVariant(frameset_source_target=_FRAME_SOURCE_APPROVED),
        candidate=DocumentVariant(frameset_source_target=_FRAME_SOURCE_CANDIDATE),
        changed_members=("word/_rels/webSettings.xml.rels",),
    ),
    CaseSpec(
        case_id="external.vml_linked_ole_object_target_retargeted",
        title="VML linked-OLE object target retargeted",
        description=(
            "A VML o:OLEObject Type=Link anchor retains its placeholder, link metadata, and "
            "relationship ID while the external standard oleObject relationship target changes."
        ),
        fact={
            "binding": "external",
            "kind": "vml_linked_ole_object_target_changed",
            "source": "word_vml",
        },
        review_expectation="block",
        baseline=DocumentVariant(vml_linked_ole_target=_LINKED_OLE_APPROVED),
        candidate=DocumentVariant(vml_linked_ole_target=_LINKED_OLE_CANDIDATE),
        changed_members=("word/_rels/document.xml.rels",),
    ),
    CaseSpec(
        case_id="external.drawing_linked_picture_target_retargeted",
        title="DrawingML linked-picture target retargeted",
        description=(
            "An a:blip r:link marker retains its stored DrawingML shape while the "
            "external image relationship target changes."
        ),
        fact={
            "binding": "external",
            "kind": "drawing_linked_picture_target_changed",
            "relationship_category": "image",
            "source": "word_drawing",
        },
        review_expectation="block",
        baseline=DocumentVariant(drawing_linked_picture_target=_LINKED_PICTURE_APPROVED),
        candidate=DocumentVariant(drawing_linked_picture_target=_LINKED_PICTURE_CANDIDATE),
        changed_members=("word/_rels/document.xml.rels",),
    ),
    CaseSpec(
        case_id="import.alternative_format_html_payload_changed",
        title="Alternative-format HTML import payload changed",
        description=(
            "A w:altChunk anchor and internal afChunk relationship retain their shape while "
            "the synthetic HTML import payload changes."
        ),
        fact={
            "kind": "alternative_format_import_payload_changed",
            "payload_kind": "html",
            "source": "word_alt_chunk",
        },
        review_expectation="block",
        baseline=DocumentVariant(alternative_format_import_payload=_ALT_CHUNK_APPROVED),
        candidate=DocumentVariant(alternative_format_import_payload=_ALT_CHUNK_CANDIDATE),
        changed_members=("word/afchunk1.html",),
    ),
    CaseSpec(
        case_id="review.hidden_text_run_added",
        title="Direct hidden-text run added",
        description=(
            "A run keeps the same stored text while direct w:vanish markup is introduced."
        ),
        fact={"kind": "hidden_text_run_added"},
        review_expectation="review",
        baseline=DocumentVariant(),
        candidate=DocumentVariant(hidden_text=True),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="review.tracked_insertion_markup_added",
        title="Tracked insertion markup added",
        description=(
            "An unchanged run is wrapped in stored w:ins review markup with fixed synthetic "
            "revision metadata."
        ),
        fact={"kind": "revision_markup_added", "revision_kind": "insertion"},
        review_expectation="review",
        baseline=DocumentVariant(),
        candidate=DocumentVariant(insertion_markup=True),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="review.track_revisions_setting_enabled",
        title="Track Changes setting enabled",
        description="The stored w:trackRevisions document setting is enabled.",
        fact={"kind": "track_revisions_setting_enabled"},
        review_expectation="review",
        baseline=DocumentVariant(),
        candidate=DocumentVariant(track_revisions=True),
        changed_members=("word/settings.xml",),
    ),
    CaseSpec(
        case_id="review.document_protection_enabled",
        title="Read-only document protection enabled",
        description=(
            "A password-free w:documentProtection declaration enables read-only editing "
            "restrictions."
        ),
        fact={"kind": "document_protection_enabled", "protection_mode": "read_only"},
        review_expectation="review",
        baseline=DocumentVariant(),
        candidate=DocumentVariant(document_protection=True),
        changed_members=("word/settings.xml",),
    ),
    CaseSpec(
        case_id="review.permission_range_editor_changed",
        title="Editable-range editor assignment changed",
        description=(
            "One paired w:permStart/w:permEnd boundary retains its marker ID and covered "
            "stored text while its synthetic individual editor assignment changes."
        ),
        fact={
            "kind": "permission_range_editor_changed",
            "source": "word_permission_markup",
        },
        review_expectation="review",
        baseline=DocumentVariant(permission_range_editor=_PERMISSION_RANGE_EDITOR_APPROVED),
        candidate=DocumentVariant(permission_range_editor=_PERMISSION_RANGE_EDITOR_CANDIDATE),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="binding.data_binding_xpath_retargeted",
        title="Content-control data binding retargeted",
        description=(
            "A content control keeps its stored display text and custom XML part while its "
            "private w:dataBinding XPath changes."
        ),
        fact={"kind": "data_binding_mapping_changed", "mapping": "custom_xml"},
        review_expectation="review",
        baseline=DocumentVariant(
            data_binding_xpath=_BINDING_XPATH_APPROVED,
            custom_xml_payload=_CUSTOM_XML_APPROVED,
        ),
        candidate=DocumentVariant(
            data_binding_xpath=_BINDING_XPATH_CANDIDATE,
            custom_xml_payload=_CUSTOM_XML_APPROVED,
        ),
        changed_members=("word/document.xml",),
    ),
    CaseSpec(
        case_id="binding.custom_xml_payload_changed",
        title="Bound custom XML payload changed",
        description=(
            "A content control's mapping stays fixed while the referenced synthetic custom XML "
            "payload changes."
        ),
        fact={"kind": "custom_xml_payload_changed"},
        review_expectation="review",
        baseline=DocumentVariant(
            data_binding_xpath=_BINDING_XPATH_APPROVED,
            custom_xml_payload=_CUSTOM_XML_APPROVED,
        ),
        candidate=DocumentVariant(
            data_binding_xpath=_BINDING_XPATH_APPROVED,
            custom_xml_payload=_CUSTOM_XML_CANDIDATE,
        ),
        changed_members=("customXml/item1.xml",),
    ),
    CaseSpec(
        case_id="macro.vba_project_payload_changed",
        title="VBA project payload changed",
        description=(
            "A macro-enabled package retains its macro relationship and contains only inert "
            "marker bytes whose payload changes."
        ),
        fact={"kind": "macro_payload_changed"},
        review_expectation="block",
        baseline=DocumentVariant(extension="docm", macro_payload=_OPAQUE_MACRO_APPROVED),
        candidate=DocumentVariant(extension="docm", macro_payload=_OPAQUE_MACRO_CANDIDATE),
        changed_members=("word/vbaProject.bin",),
    ),
    CaseSpec(
        case_id="embedded.ole_payload_changed",
        title="Embedded OLE payload changed",
        description=(
            "A fixed internal OLE relationship and VML embedded-object marker retain their "
            "shape while opaque synthetic payload bytes change."
        ),
        fact={"kind": "embedded_ole_payload_changed"},
        review_expectation="block",
        baseline=DocumentVariant(embedded_payload=_OPAQUE_OLE_APPROVED),
        candidate=DocumentVariant(embedded_payload=_OPAQUE_OLE_CANDIDATE),
        changed_members=("word/embeddings/oleObject1.bin",),
    ),
    CaseSpec(
        case_id="embedded.activex_control_persistence_payload_changed",
        title="ActiveX control persistence payload changed",
        description=(
            "A fixed w:control anchor, ActiveX persistence part, and internal binary "
            "relationship retain their topology while inert synthetic persistence bytes change."
        ),
        fact={
            "kind": "activex_control_persistence_payload_changed",
            "source": "word_embedded_control",
        },
        review_expectation="block",
        baseline=DocumentVariant(activex_persistence_payload=_OPAQUE_ACTIVE_X_APPROVED),
        candidate=DocumentVariant(activex_persistence_payload=_OPAQUE_ACTIVE_X_CANDIDATE),
        changed_members=("word/activeX/activeX1.bin",),
    ),
)

CASE_IDS = tuple(spec.case_id for spec in CASE_SPECS)


def truth_manifest(spec: CaseSpec) -> dict[str, Any]:
    """Return the public, target-free truth contract for ``spec``."""

    return {
        "baseline": spec.baseline_name,
        "candidate": spec.candidate_name,
        "coverage_expectations": [],
        "description": spec.description,
        "facts": [spec.fact],
        "id": spec.case_id,
        "invariants": {
            "changed_member_count": len(spec.changed_members),
            "package_member_set_stable": True,
            "stored_text_stable": True,
        },
        "review_expectation": spec.review_expectation,
        "schema_version": FIXTURE_SCHEMA_VERSION,
        "title": spec.title,
    }


def case_files(spec: CaseSpec) -> dict[str, bytes]:
    """Render every file for one case without touching the filesystem."""

    truth = json.dumps(truth_manifest(spec), indent=2, sort_keys=True).encode("utf-8") + b"\n"
    return {
        spec.baseline_name: render_package(spec.baseline),
        spec.candidate_name: render_package(spec.candidate),
        "truth.json": truth,
    }


def render_package(variant: DocumentVariant) -> bytes:
    """Create a deterministic, compact, nonexecuting WordprocessingML package."""

    _validate_variant(variant)
    members: dict[str, bytes] = {
        "[Content_Types].xml": _content_types(variant),
        "_rels/.rels": _root_relationships(),
        "word/document.xml": _document_xml(variant),
        "word/_rels/document.xml.rels": _document_relationships(variant),
        "word/settings.xml": _settings_xml(variant),
        "word/styles.xml": _styles_xml(),
    }
    if variant.frameset_source_target is not None:
        members.update(
            {
                "word/webSettings.xml": _web_settings_xml(),
                "word/_rels/webSettings.xml.rels": _web_settings_relationships(variant),
            }
        )
    if (
        variant.attached_template_target is not None
        or variant.mail_merge_data_source_target is not None
    ):
        members["word/_rels/settings.xml.rels"] = _settings_relationships(variant)
    if variant.alternative_format_import_payload is not None:
        members["word/afchunk1.html"] = variant.alternative_format_import_payload
    if variant.custom_xml_payload is not None:
        members.update(
            {
                "customXml/item1.xml": variant.custom_xml_payload,
                "customXml/_rels/item1.xml.rels": _custom_xml_relationships(),
                "customXml/itemProps1.xml": _custom_xml_properties(),
            }
        )
    if variant.macro_payload is not None:
        members["word/vbaProject.bin"] = variant.macro_payload
    if variant.embedded_payload is not None:
        members["word/embeddings/oleObject1.bin"] = variant.embedded_payload
    if variant.activex_persistence_payload is not None:
        members.update(
            {
                "word/activeX/activeX1.xml": _active_x_control_xml(),
                "word/activeX/_rels/activeX1.xml.rels": _active_x_control_relationships(),
                "word/activeX/activeX1.bin": variant.activex_persistence_payload,
            }
        )
    if variant.modern_comment_done is not None:
        members.update(
            {
                "word/comments.xml": _comments_xml(),
                "word/commentsExtended.xml": _comments_extended_xml(variant.modern_comment_done),
            }
        )
    if variant.taskpane_auto_show is not None:
        members.update(
            {
                "word/webextensions/taskpanes.xml": _taskpane_web_extension_taskpanes(),
                "word/webextensions/_rels/taskpanes.xml.rels": (
                    _taskpane_web_extension_relationships()
                ),
                "word/webextensions/webextension1.xml": _taskpane_web_extension(
                    variant.taskpane_auto_show
                ),
            }
        )
    return _zip_members(members)


def build_fixtures(fixture_root: str | Path, *, force: bool = False) -> dict[str, int]:
    """Write a complete deterministic tree, refusing unknown existing output."""

    root = Path(fixture_root)
    if root.exists() and (root.is_symlink() or not root.is_dir()):
        raise FixtureBuildError("fixture destination must be a non-symlink directory")
    root.mkdir(parents=True, exist_ok=True)
    expected_root_names = {"manifest.jsonl", *CASE_IDS}
    existing_root_names = {child.name for child in root.iterdir()}
    if existing_root_names - expected_root_names:
        raise FixtureBuildError("fixture destination contains unknown entries")
    if existing_root_names and not force:
        raise FixtureBuildError(
            "fixture destination is not empty; use --force to replace DCAB files"
        )

    for spec in CASE_SPECS:
        _write_case(root / spec.case_id, case_files(spec), force=force)
    manifest = b"".join(
        json.dumps(
            {
                "id": spec.case_id,
                "schema_version": FIXTURE_SCHEMA_VERSION,
                "truth": f"{spec.case_id}/truth.json",
            },
            sort_keys=True,
        ).encode("utf-8")
        + b"\n"
        for spec in CASE_SPECS
    )
    _write_regular_file(root / "manifest.jsonl", manifest, force=force)
    return {"case_count": len(CASE_SPECS), "fixture_schema_version": FIXTURE_SCHEMA_VERSION}


def _validate_variant(variant: DocumentVariant) -> None:
    if variant.extension not in {"docx", "docm"}:
        raise FixtureBuildError("fixture package extension is unsupported")
    if (variant.extension == "docm") != (variant.macro_payload is not None):
        raise FixtureBuildError("macro-enabled packages must have exactly one inert macro payload")
    if variant.data_binding_xpath is not None and variant.custom_xml_payload is None:
        raise FixtureBuildError("data-binding fixtures require a custom XML payload")
    if variant.data_binding_xpath is None and variant.custom_xml_payload is not None:
        raise FixtureBuildError("custom XML fixtures require a data-binding declaration")
    if variant.direct_hyperlink_target is not None and not variant.direct_hyperlink_target:
        raise FixtureBuildError("hyperlink target cannot be empty")
    if variant.vml_shape_hyperlink_target is not None and not variant.vml_shape_hyperlink_target:
        raise FixtureBuildError("VML shape hyperlink target cannot be empty")
    if variant.hyperlink_field_target is not None and not variant.hyperlink_field_target:
        raise FixtureBuildError("HYPERLINK target cannot be empty")
    if variant.include_text_target is not None and not variant.include_text_target:
        raise FixtureBuildError("INCLUDETEXT target cannot be empty")
    if variant.complex_include_text_target is not None and not variant.complex_include_text_target:
        raise FixtureBuildError("complex INCLUDETEXT target cannot be empty")
    if variant.include_text_target is not None and variant.complex_include_text_target is not None:
        raise FixtureBuildError("simple and complex INCLUDETEXT targets are mutually exclusive")
    if variant.dde_source_file is not None and not variant.dde_source_file:
        raise FixtureBuildError("DDE source file cannot be empty")
    if variant.document_variable_value is not None and not variant.document_variable_value:
        raise FixtureBuildError("document-variable value cannot be empty")
    if variant.permission_range_editor is not None and not variant.permission_range_editor:
        raise FixtureBuildError("permission-range editor cannot be empty")
    if variant.taskpane_auto_show is not None and not isinstance(variant.taskpane_auto_show, bool):
        raise FixtureBuildError("task-pane auto-show setting must be boolean")
    if variant.modern_comment_done is not None and not isinstance(
        variant.modern_comment_done, bool
    ):
        raise FixtureBuildError("modern-comment done setting must be boolean")
    if variant.attached_template_target is not None and not variant.attached_template_target:
        raise FixtureBuildError("attached template target cannot be empty")
    if (
        variant.mail_merge_data_source_target is not None
        and not variant.mail_merge_data_source_target
    ):
        raise FixtureBuildError("mail-merge data-source target cannot be empty")
    if variant.subdocument_target is not None and not variant.subdocument_target:
        raise FixtureBuildError("subdocument target cannot be empty")
    if variant.frameset_source_target is not None and not variant.frameset_source_target:
        raise FixtureBuildError("frameset source target cannot be empty")
    if variant.vml_linked_ole_target is not None and not variant.vml_linked_ole_target:
        raise FixtureBuildError("VML linked-OLE target cannot be empty")
    if (
        variant.alternative_format_import_payload is not None
        and not variant.alternative_format_import_payload
    ):
        raise FixtureBuildError("alternative-format import payload cannot be empty")
    if (
        variant.drawing_linked_picture_target is not None
        and not variant.drawing_linked_picture_target
    ):
        raise FixtureBuildError("linked-picture target cannot be empty")
    if variant.activex_persistence_payload is not None and not variant.activex_persistence_payload:
        raise FixtureBuildError("ActiveX persistence payload cannot be empty")


def _write_case(case_dir: Path, files: dict[str, bytes], *, force: bool) -> None:
    if case_dir.exists() and (case_dir.is_symlink() or not case_dir.is_dir()):
        raise FixtureBuildError("fixture case destination must be a non-symlink directory")
    case_dir.mkdir(exist_ok=True)
    existing_names = {child.name for child in case_dir.iterdir()}
    if existing_names - set(files):
        raise FixtureBuildError("fixture case destination contains unknown entries")
    if existing_names and not force:
        raise FixtureBuildError(
            "fixture case destination is not empty; use --force to replace DCAB files"
        )
    for name, data in files.items():
        _write_regular_file(case_dir / name, data, force=force)


def _write_regular_file(path: Path, data: bytes, *, force: bool) -> None:
    if path.exists() or path.is_symlink():
        if path.is_symlink() or not path.is_file():
            raise FixtureBuildError("fixture output must be a regular file")
        if not force:
            raise FixtureBuildError(
                "fixture output already exists; use --force to replace DCAB files"
            )
    path.write_bytes(data)


def _content_types(variant: DocumentVariant) -> bytes:
    main_type = _DOCM_MAIN_CONTENT_TYPE if variant.extension == "docm" else _DOCX_MAIN_CONTENT_TYPE
    overrides = [
        ("/word/document.xml", main_type),
        ("/word/settings.xml", _SETTINGS_CONTENT_TYPE),
        ("/word/styles.xml", _STYLES_CONTENT_TYPE),
    ]
    if variant.custom_xml_payload is not None:
        overrides.append(("/customXml/itemProps1.xml", _CUSTOM_XML_PROPERTIES_CONTENT_TYPE))
    if variant.alternative_format_import_payload is not None:
        overrides.append(("/word/afchunk1.html", _ALT_CHUNK_CONTENT_TYPE))
    if variant.macro_payload is not None:
        overrides.append(("/word/vbaProject.bin", _VBA_PROJECT_CONTENT_TYPE))
    if variant.embedded_payload is not None:
        overrides.append(("/word/embeddings/oleObject1.bin", _OLE_CONTENT_TYPE))
    if variant.activex_persistence_payload is not None:
        overrides.extend(
            (
                ("/word/activeX/activeX1.xml", _ACTIVE_X_CONTENT_TYPE),
                ("/word/activeX/activeX1.bin", _ACTIVE_X_BINARY_CONTENT_TYPE),
            )
        )
    if variant.frameset_source_target is not None:
        overrides.append(("/word/webSettings.xml", _WEB_SETTINGS_CONTENT_TYPE))
    if variant.modern_comment_done is not None:
        overrides.extend(
            (
                ("/word/comments.xml", _COMMENTS_CONTENT_TYPE),
                ("/word/commentsExtended.xml", _COMMENTS_EXTENDED_CONTENT_TYPE),
            )
        )
    if variant.taskpane_auto_show is not None:
        overrides.extend(
            (
                (
                    "/word/webextensions/taskpanes.xml",
                    _TASKPANE_WEB_EXTENSION_TASKPANES_CONTENT_TYPE,
                ),
                (
                    "/word/webextensions/webextension1.xml",
                    _TASKPANE_WEB_EXTENSION_CONTENT_TYPE,
                ),
            )
        )
    values = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Types xmlns="{_CONTENT_TYPES_NS}">',
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
        '<Default Extension="xml" ContentType="application/xml"/>',
    ]
    values.extend(
        f'<Override PartName="{part_name}" ContentType="{content_type}"/>'
        for part_name, content_type in overrides
    )
    values.append("</Types>")
    return "".join(values).encode("utf-8")


def _root_relationships() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        f'<Relationship Id="rIdOfficeDocument" Type="{_OFFICE_DOCUMENT_RELATIONSHIP}" '
        'Target="word/document.xml"/>'
        "</Relationships>"
    ).encode()


def _document_relationships(variant: DocumentVariant) -> bytes:
    relationships = [
        ("rIdStyles", _STYLES_RELATIONSHIP, "styles.xml", "Internal"),
        ("rIdSettings", _SETTINGS_RELATIONSHIP, "settings.xml", "Internal"),
    ]
    if variant.direct_hyperlink_target is not None:
        relationships.append(
            ("rIdHyperlink", _HYPERLINK_RELATIONSHIP, variant.direct_hyperlink_target, "External")
        )
    if variant.drawing_linked_picture_target is not None:
        relationships.append(
            (
                "rIdLinkedPicture",
                _IMAGE_RELATIONSHIP,
                variant.drawing_linked_picture_target,
                "External",
            )
        )
    if variant.subdocument_target is not None:
        relationships.append(
            ("rIdSubDocument", _SUBDOCUMENT_RELATIONSHIP, variant.subdocument_target, "External")
        )
    if variant.frameset_source_target is not None:
        relationships.append(
            ("rIdWebSettings", _WEB_SETTINGS_RELATIONSHIP, "webSettings.xml", "Internal")
        )
    if variant.vml_linked_ole_target is not None:
        relationships.append(
            (
                "rIdLinkedOleObject",
                _OLE_OBJECT_RELATIONSHIP,
                variant.vml_linked_ole_target,
                "External",
            )
        )
    if variant.alternative_format_import_payload is not None:
        relationships.append(("rIdAltChunk", _ALT_CHUNK_RELATIONSHIP, "afchunk1.html", "Internal"))
    if variant.custom_xml_payload is not None:
        relationships.append(
            ("rIdCustomXml", _CUSTOM_XML_RELATIONSHIP, "../customXml/item1.xml", "Internal")
        )
    if variant.macro_payload is not None:
        relationships.append(
            ("rIdVbaProject", _VBA_PROJECT_RELATIONSHIP, "vbaProject.bin", "Internal")
        )
    if variant.embedded_payload is not None:
        relationships.append(
            ("rIdOleObject", _OLE_OBJECT_RELATIONSHIP, "embeddings/oleObject1.bin", "Internal")
        )
    if variant.activex_persistence_payload is not None:
        relationships.append(
            ("rIdActiveXControl", _CONTROL_RELATIONSHIP, "activeX/activeX1.xml", "Internal")
        )
    if variant.modern_comment_done is not None:
        relationships.extend(
            (
                ("rIdComments", _COMMENTS_RELATIONSHIP, "comments.xml", "Internal"),
                (
                    "rIdCommentsExtended",
                    _COMMENTS_EXTENDED_RELATIONSHIP,
                    "commentsExtended.xml",
                    "Internal",
                ),
            )
        )
    if variant.taskpane_auto_show is not None:
        relationships.append(
            (
                "rIdTaskpaneWebExtensions",
                _TASKPANE_WEB_EXTENSION_TASKPANES_RELATIONSHIP,
                "webextensions/taskpanes.xml",
                "Internal",
            )
        )
    values = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">',
    ]
    for relationship_id, relationship_type, target, target_mode in relationships:
        mode = ' TargetMode="External"' if target_mode == "External" else ""
        values.append(
            f'<Relationship Id="{html.escape(relationship_id, quote=True)}" '
            f'Type="{html.escape(relationship_type, quote=True)}" '
            f'Target="{html.escape(target, quote=True)}"{mode}/>'
        )
    values.append("</Relationships>")
    return "".join(values).encode("utf-8")


def _taskpane_web_extension_taskpanes() -> bytes:
    """Return one internal-only, invisible task-pane declaration."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<wetp:taskpanes xmlns:wetp="{_TASKPANE_WEB_EXTENSION_TASKPANES_NS}" '
        f'xmlns:r="{_REL_NS}">'
        '<wetp:taskpane dockstate="right" visibility="0" width="350" row="0" locked="false">'
        '<wetp:webextensionref r:id="rIdTaskpaneWebExtension"/>'
        "</wetp:taskpane></wetp:taskpanes>"
    ).encode()


def _taskpane_web_extension_relationships() -> bytes:
    """Return the internal task-pane-to-web-extension relationship."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        '<Relationship Id="rIdTaskpaneWebExtension" '
        f'Type="{_TASKPANE_WEB_EXTENSION_RELATIONSHIP}" '
        'Target="webextension1.xml"/>'
        "</Relationships>"
    ).encode()


def _taskpane_web_extension(auto_show: bool) -> bytes:
    """Return a complete Office web-extension part without executable content."""

    value = "true" if auto_show else "false"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<we:webextension xmlns:we="{_TASKPANE_WEB_EXTENSION_NS}" '
        f'id="{_TASKPANE_WEB_EXTENSION_ID}">'
        f'<we:reference id="{_TASKPANE_WEB_EXTENSION_REFERENCE_ID}" '
        f'version="{_TASKPANE_WEB_EXTENSION_REFERENCE_VERSION}" '
        f'store="{_TASKPANE_WEB_EXTENSION_REFERENCE_STORE}" '
        f'storeType="{_TASKPANE_WEB_EXTENSION_REFERENCE_STORE_TYPE}"/>'
        "<we:alternateReferences/><we:properties>"
        f'<we:property name="{_TASKPANE_AUTO_SHOW_PROPERTY_NAME}" value="{value}"/>'
        "</we:properties><we:bindings/>"
        f'<we:snapshot xmlns:r="{_REL_NS}"/>'
        "</we:webextension>"
    ).encode()


def _document_xml(variant: DocumentVariant) -> bytes:
    drawing_namespaces = (
        f' xmlns:a="{_DRAWING_NS}" xmlns:wp="{_WORDPROCESSING_DRAWING_NS}" '
        f'xmlns:pic="{_PICTURE_NS}"'
        if variant.drawing_linked_picture_target is not None
        else ""
    )
    hyperlink_run = _run(_HYPERLINK_DISPLAY_TEXT)
    hyperlink_markup = (
        f'<w:hyperlink r:id="rIdHyperlink">{hyperlink_run}</w:hyperlink>'
        if variant.direct_hyperlink_target is not None
        else hyperlink_run
    )
    vml_shape_hyperlink_markup = (
        _vml_shape_hyperlink_markup(variant.vml_shape_hyperlink_target)
        if variant.vml_shape_hyperlink_target is not None
        else ""
    )
    hyperlink_field_markup = (
        _simple_field(f' HYPERLINK "{variant.hyperlink_field_target}" ', _HYPERLINK_FIELD_RESULT)
        if variant.hyperlink_field_target is not None
        else _run(_HYPERLINK_FIELD_RESULT)
    )
    if variant.include_text_target is not None:
        include_field_markup = _simple_field(
            f' INCLUDETEXT "{variant.include_text_target}" ', _INCLUDE_TEXT_FIELD_RESULT
        )
    elif variant.complex_include_text_target is not None:
        include_field_markup = _complex_include_text_field(variant.complex_include_text_target)
    else:
        include_field_markup = _run(_INCLUDE_TEXT_FIELD_RESULT)
    dde_field_markup = (
        _simple_field(_dde_field_instruction(variant.dde_source_file), _DDE_FIELD_RESULT)
        if variant.dde_source_file is not None
        else ""
    )
    document_variable_field_markup = (
        _simple_field(_document_variable_field_instruction(), _DOCUMENT_VARIABLE_FIELD_RESULT)
        if variant.document_variable_value is not None
        else ""
    )
    permission_range_markup = (
        _permission_range_markup(variant.permission_range_editor)
        if variant.permission_range_editor is not None
        else ""
    )
    modern_comment_anchor_markup = (
        _modern_comment_anchor_markup() if variant.modern_comment_done is not None else ""
    )
    hidden_properties = "<w:rPr><w:vanish/></w:rPr>" if variant.hidden_text else ""
    hidden_markup = f"<w:r>{hidden_properties}<w:t>{_HIDDEN_TEXT}</w:t></w:r>"
    revision_run = _run(_REVISION_TEXT)
    revision_markup = (
        f'<w:ins w:id="1" w:author="DCAB" w:date="2026-08-03T00:00:00Z">{revision_run}</w:ins>'
        if variant.insertion_markup
        else revision_run
    )
    binding_markup = _binding_markup(variant.data_binding_xpath)
    ole_markup = _ole_markup() if variant.embedded_payload is not None else ""
    active_x_markup = _active_x_markup() if variant.activex_persistence_payload is not None else ""
    linked_ole_markup = _linked_ole_markup() if variant.vml_linked_ole_target is not None else ""
    linked_picture_markup = (
        _linked_picture_markup() if variant.drawing_linked_picture_target is not None else ""
    )
    subdocument_markup = (
        '<w:subDoc r:id="rIdSubDocument"/>' if variant.subdocument_target is not None else ""
    )
    alt_chunk_markup = (
        '<w:altChunk r:id="rIdAltChunk"/>'
        if variant.alternative_format_import_payload is not None
        else ""
    )
    value = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:document xmlns:w="{_WORD_NS}" xmlns:r="{_REL_NS}" '
        f'xmlns:v="{_VML_NS}" xmlns:o="{_OFFICE_VML_NS}"{drawing_namespaces}>'
        "<w:body><w:p>"
        f"{_run(_VISIBLE_TEXT)}{hyperlink_markup}{vml_shape_hyperlink_markup}"
        f"{hyperlink_field_markup}{include_field_markup}"
        f"{dde_field_markup}{document_variable_field_markup}{permission_range_markup}"
        f"{modern_comment_anchor_markup}{hidden_markup}{revision_markup}{binding_markup}"
        f"{ole_markup}{active_x_markup}{linked_ole_markup}{linked_picture_markup}"
        f'</w:p>{subdocument_markup}{alt_chunk_markup}<w:sectPr><w:pgSz w:w="12240" w:h="15840"/>'
        '<w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>'
        "</w:body></w:document>"
    )
    return value.encode("utf-8")


def _run(text: str) -> str:
    return f"<w:r><w:t>{html.escape(text)}</w:t></w:r>"


def _simple_field(instruction: str, result: str) -> str:
    return (
        f'<w:fldSimple w:instr="{html.escape(instruction, quote=True)}">'
        f"{_run(result)}</w:fldSimple>"
    )


def _complex_include_text_field(target: str) -> str:
    """Return one complete, fragmented complex INCLUDETEXT field."""

    instruction_markup = "".join(
        f'<w:r><w:instrText xml:space="preserve">{html.escape(chunk)}</w:instrText></w:r>'
        for chunk in _complex_include_text_instruction_chunks(target)
    )
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        f"{instruction_markup}"
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f"{_run(_INCLUDE_TEXT_FIELD_RESULT)}"
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def _complex_include_text_instruction_chunks(target: str) -> tuple[str, str, str]:
    """Return the fixed three-run instruction shape for a complex INCLUDETEXT field."""

    return (
        " INCLUDE",
        f'TEXT "{target[:8]}',
        f'{target[8:]}" ',
    )


def _dde_field_instruction(source_file: str) -> str:
    """Return the fixed-shape, nonexecuting DDE field instruction for a fixture."""

    return f' DDE {_DDE_APPLICATION} "{source_file}" "{_DDE_ITEM}" '


def _document_variable_field_instruction() -> str:
    """Return the fixed DOCVARIABLE field instruction for a fixture."""

    return f" DOCVARIABLE {_DOCUMENT_VARIABLE_NAME} "


def _permission_range_markup(editor: str) -> str:
    """Return one paired, synthetic editable-range permission boundary."""

    escaped_editor = html.escape(editor, quote=True)
    return (
        f'<w:permStart w:id="{_PERMISSION_RANGE_MARKER_ID}" w:ed="{escaped_editor}"/>'
        f"{_run(_PERMISSION_RANGE_TEXT)}"
        f'<w:permEnd w:id="{_PERMISSION_RANGE_MARKER_ID}"/>'
    )


def _modern_comment_anchor_markup() -> str:
    """Return one fixed classic-comment anchor for commentsExtended metadata."""

    return (
        f'<w:commentRangeStart w:id="{_MODERN_COMMENT_ID}"/>'
        f"{_run(_MODERN_COMMENT_ANCHOR_TEXT)}"
        f'<w:commentRangeEnd w:id="{_MODERN_COMMENT_ID}"/>'
        f'<w:r><w:commentReference w:id="{_MODERN_COMMENT_ID}"/></w:r>'
    )


def _comments_xml() -> bytes:
    """Return one fixed classic comment associated with commentsExtended metadata."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:comments xmlns:w="{_WORD_NS}" xmlns:w14="{_WORD_2010_WORDML_NS}">'
        f'<w:comment w:id="{_MODERN_COMMENT_ID}" '
        f'w:author="{_MODERN_COMMENT_AUTHOR}" w:initials="{_MODERN_COMMENT_INITIALS}" '
        f'w:date="{_MODERN_COMMENT_DATE}">'
        f'<w:p w14:paraId="{_MODERN_COMMENT_PARAGRAPH_ID}" '
        f'w14:textId="{_MODERN_COMMENT_TEXT_ID}">'
        f"<w:r><w:annotationRef/></w:r>{_run(_MODERN_COMMENT_TEXT)}"
        "</w:p></w:comment></w:comments>"
    ).encode()


def _comments_extended_xml(done: bool) -> bytes:
    """Return Office 2013 comment metadata with one explicit done-state value."""

    value = "1" if done else "0"
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w15:commentsEx xmlns:w15="{_WORD_2012_WORDML_NS}">'
        f'<w15:commentEx w15:paraId="{_MODERN_COMMENT_PARAGRAPH_ID}" '
        f'w15:done="{value}"/>'
        "</w15:commentsEx>"
    ).encode()


def _vml_shape_hyperlink_markup(target: str) -> str:
    """Return one direct legacy VML shape hyperlink without a relationship."""

    return (
        f'<w:r><w:pict><v:rect id="{_VML_SHAPE_ID}" '
        'style="width:1pt;height:1pt" filled="f" stroked="f" '
        f'href="{html.escape(target, quote=True)}" target="{_VML_SHAPE_TARGET_FRAME}"/>'
        "</w:pict></w:r>"
    )


def _binding_markup(xpath: str | None) -> str:
    if xpath is None:
        return _run(_BINDING_TEXT)
    return (
        '<w:sdt><w:sdtPr><w:id w:val="-1"/>'
        f'<w:dataBinding w:xpath="{html.escape(xpath, quote=True)}" '
        "w:prefixMappings=\"xmlns:dcab='urn:dcab:fixture'\" "
        f'w:storeItemID="{_CUSTOM_XML_STORE_ID}"/>'
        f"</w:sdtPr><w:sdtContent>{_run(_BINDING_TEXT)}</w:sdtContent></w:sdt>"
    )


def _ole_markup() -> str:
    return (
        '<w:r><w:object><v:shape id="DCABOleShape" '
        'style="width:0;height:0"><o:OLEObject Type="Embed" ProgID="DCAB.Synthetic" '
        'r:id="rIdOleObject"/></v:shape></w:object></w:r>'
    )


def _active_x_markup() -> str:
    """Return a fixed inline Word embedded-control anchor."""

    return (
        '<w:r><w:object><w:control r:id="rIdActiveXControl" '
        f'w:name="{_ACTIVE_X_CONTROL_NAME}"/></w:object></w:r>'
    )


def _active_x_control_xml() -> bytes:
    """Return a fixed ActiveX persistence XML part with one binary reference."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<ax:ocx xmlns:ax="{_ACTIVE_X_NS}" xmlns:r="{_REL_NS}" '
        f'ax:classid="{_ACTIVE_X_CLASS_ID}" ax:persistence="persistStorage" '
        'r:id="rIdActiveXBinary"/>'
    ).encode()


def _active_x_control_relationships() -> bytes:
    """Return the required internal relationship from persistence XML to binary data."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        f'<Relationship Id="rIdActiveXBinary" Type="{_ACTIVE_X_BINARY_RELATIONSHIP}" '
        'Target="activeX1.bin"/>'
        "</Relationships>"
    ).encode()


def _linked_ole_markup() -> str:
    """Return a fixed VML linked-OLE anchor backed by an external relationship."""

    return (
        f'<w:r><w:object><v:shape id="{_LINKED_OLE_SHAPE_ID}" '
        'style="width:1pt;height:1pt" o:ole=""/>'
        f'<o:OLEObject Type="Link" ProgID="{_LINKED_OLE_PROG_ID}" '
        f'ShapeID="{_LINKED_OLE_SHAPE_ID}" DrawAspect="Content" '
        f'ObjectID="{_LINKED_OLE_OBJECT_ID}" r:id="rIdLinkedOleObject" '
        f'UpdateMode="{_LINKED_OLE_UPDATE_MODE}"/>'
        "</w:object></w:r>"
    )


def _linked_picture_markup() -> str:
    """Return a compact, relationship-backed DrawingML linked-picture marker."""

    return (
        '<w:r><w:drawing><wp:inline><wp:extent cx="12700" cy="12700"/>'
        '<wp:docPr id="1" name="DCAB linked picture"/>'
        '<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/>'
        "</wp:cNvGraphicFramePr><a:graphic><a:graphicData "
        f'uri="{_PICTURE_NS}"><pic:pic><pic:nvPicPr>'
        '<pic:cNvPr id="0" name="dcab-linked-picture.png"/><pic:cNvPicPr/>'
        '</pic:nvPicPr><pic:blipFill><a:blip r:link="rIdLinkedPicture"/>'
        "<a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr>"
        '<a:xfrm><a:off x="0" y="0"/><a:ext cx="12700" cy="12700"/>'
        '</a:xfrm><a:prstGeom prst="rect"/></pic:spPr></pic:pic>'
        "</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>"
    )


def _settings_xml(variant: DocumentVariant) -> bytes:
    attached_template = (
        '<w:attachedTemplate r:id="rIdAttachedTemplate"/>'
        if variant.attached_template_target is not None
        else ""
    )
    mail_merge = (
        '<w:mailMerge><w:mainDocumentType w:val="formLetters"/>'
        '<w:dataSource r:id="rIdMailMergeSource"/></w:mailMerge>'
        if variant.mail_merge_data_source_target is not None
        else ""
    )
    document_variables = (
        _document_variables_markup(variant.document_variable_value)
        if variant.document_variable_value is not None
        else ""
    )
    relationship_namespace = f' xmlns:r="{_REL_NS}"' if attached_template or mail_merge else ""
    track_revisions = "<w:trackRevisions/>" if variant.track_revisions else ""
    protection = (
        '<w:documentProtection w:edit="readOnly" w:enforcement="1"/>'
        if variant.document_protection
        else ""
    )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:settings xmlns:w="{_WORD_NS}"{relationship_namespace}>'
        f"{attached_template}{mail_merge}{track_revisions}{protection}{document_variables}"
        "</w:settings>"
    ).encode()


def _document_variables_markup(value: str) -> str:
    """Return one deterministic, persisted document-variable declaration."""

    return (
        "<w:docVars>"
        f'<w:docVar w:name="{html.escape(_DOCUMENT_VARIABLE_NAME, quote=True)}" '
        f'w:val="{html.escape(value, quote=True)}"/>'
        "</w:docVars>"
    )


def _settings_relationships(variant: DocumentVariant) -> bytes:
    relationships: list[tuple[str, str, str]] = []
    if variant.attached_template_target is not None:
        relationships.append(
            (
                "rIdAttachedTemplate",
                _ATTACHED_TEMPLATE_RELATIONSHIP,
                variant.attached_template_target,
            )
        )
    if variant.mail_merge_data_source_target is not None:
        relationships.append(
            (
                "rIdMailMergeSource",
                _MAIL_MERGE_SOURCE_RELATIONSHIP,
                variant.mail_merge_data_source_target,
            )
        )
    if not relationships:
        raise FixtureBuildError("settings relationships require a relationship-backed setting")
    values = [
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">',
    ]
    values.extend(
        f'<Relationship Id="{relationship_id}" Type="{relationship_type}" '
        f'Target="{html.escape(target, quote=True)}" TargetMode="External"/>'
        for relationship_id, relationship_type, target in relationships
    )
    values.append("</Relationships>")
    return "".join(values).encode()


def _web_settings_xml() -> bytes:
    """Return one fixed root frameset with a single source-backed frame."""

    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:webSettings xmlns:w="{_WORD_NS}" xmlns:r="{_REL_NS}">'
        f'<w:frameset><w:frameLayout w:val="{_FRAME_LAYOUT}"/>'
        f'<w:frame><w:sz w:val="{_FRAME_SIZE}"/>'
        f'<w:name w:val="{_FRAME_NAME}"/>'
        '<w:sourceFileName r:id="rIdFrameSource"/>'
        "</w:frame></w:frameset></w:webSettings>"
    ).encode()


def _web_settings_relationships(variant: DocumentVariant) -> bytes:
    """Return the external source relationship for the fixed frameset anchor."""

    target = variant.frameset_source_target
    if target is None:
        raise FixtureBuildError("web settings relationships require a frame source")
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        f'<Relationship Id="rIdFrameSource" Type="{_FRAME_RELATIONSHIP}" '
        f'Target="{html.escape(target, quote=True)}" TargetMode="External"/>'
        "</Relationships>"
    ).encode()


def _styles_xml() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles xmlns:w="{_WORD_NS}"><w:docDefaults><w:rPrDefault><w:rPr/>'
        "</w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults></w:styles>"
    ).encode()


def _custom_xml_relationships() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<Relationships xmlns="{_PACKAGE_REL_NS}">'
        f'<Relationship Id="rIdItemProperties" Type="{_CUSTOM_XML_PROPERTIES_RELATIONSHIP}" '
        'Target="itemProps1.xml"/>'
        "</Relationships>"
    ).encode()


def _custom_xml_properties() -> bytes:
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<ds:datastoreItem xmlns:ds="{_CUSTOM_XML_PROPERTIES_NS}" '
        f'ds:itemID="{_CUSTOM_XML_STORE_ID}"/>'
    ).encode()


def _zip_members(members: dict[str, bytes]) -> bytes:
    stream = io.BytesIO()
    with zipfile.ZipFile(stream, "w", compression=zipfile.ZIP_STORED) as archive:
        for name in sorted(members):
            info = zipfile.ZipInfo(name, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_STORED
            info.create_system = 3
            info.external_attr = (stat.S_IFREG | 0o644) << 16
            archive.writestr(info, members[name])
    return stream.getvalue()
