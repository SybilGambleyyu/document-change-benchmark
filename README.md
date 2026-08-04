# Document Change Assurance Benchmark (DCAB)

DCAB is an open, deterministic corpus for evaluating static review tools that compare WordprocessingML packages. It supplies thirty-two paired synthetic `.docx`/`.docm` fixtures, a privacy-safe public oracle, an observation schema, and a scorer.

It is for a question ordinary text diffs do not answer well: did a document change in a stored review-sensitive surface even when ordinary text is unchanged? The corpus covers direct Word and legacy VML shape hyperlinks, simple and fragmented complex field instructions, DDE field sources and persisted document variables, editable-range permission markup, task-pane Office web-extension configuration, Office 2013 `commentsExtended` done metadata, attached-template, mail-merge data-source and recipient-selection settings, save-through-XSLT, automatic-field-recalculation-on-open, and automatic-template-style-update-on-open configuration, master-subdocument, frameset-source, legacy VML linked-OLE, and DrawingML linked-picture relationships, alternative-format import payloads, hidden text and revision markup, review settings and document protection, content-control bindings and custom XML, plus opaque macro, embedded-OLE, and ActiveX control-persistence payload boundaries.

DCAB is not a Word renderer, a macro scanner, a field evaluator, or a runtime behavior benchmark. It never resolves an external target, opens Word, updates a field, parses an opaque payload, instantiates an ActiveX control, or executes code.

## Install and use

```bash
python -m pip install document-change-benchmark
dcab validate
dcab observation-template --output observations.json
dcab score --observations observations.json
```

Use `--strict` with `score` to return a nonzero status unless every case is analyzed with every declared fact and the reference review convention.

```bash
dcab score --observations observations.json --strict --output score.json
```

The wheel bundles the corpus, so `dcab validate` needs no network access. To reproduce both repository and bundled copies from source:

```bash
dcab build --fixtures fixtures --force
dcab build --fixtures src/dcab/fixtures --force
dcab validate --fixtures fixtures
```

## Corpus contract

Each case contains:

- `baseline.docx` or `baseline.docm`
- `candidate.docx` or `candidate.docm`
- `truth.json`, a target-free public assertion

`manifest.jsonl` catalogues the thirty-two cases. Every pair has the same package-member set, differs only at a declared member boundary, and retains the same sequence of stored `w:t` values. That invariant is intentionally narrower than visual or client-runtime equivalence.

Version 0.21 adds a direct `w:linkStyles` transition from explicitly disabled to enabled while retaining a fixed `w:attachedTemplate` anchor, external Settings relationship, and synthetic target. Only `word/settings.xml` changes; the corpus does not resolve or load the template, open a client, propagate styles, or claim that a client will do so. Fixture schema version 1 and the truth and observation envelopes are unchanged. An earlier observation can still be parsed, but it is incomplete when scored against this thirty-two-case catalogue.

| Case | Declared fact | Reference convention |
| --- | --- | --- |
| `interaction.word_hyperlink_target_retargeted` | `word_hyperlink_target_changed` | block |
| `interaction.word_hyperlink_added` | `word_hyperlink_added` | block |
| `interaction.vml_shape_hyperlink_target_retargeted` | `vml_shape_hyperlink_target_changed` | block |
| `interaction.word_hyperlink_field_target_retargeted` | `field_target_changed` | block |
| `interaction.taskpane_auto_show_setting_enabled` | `taskpane_auto_show_setting_enabled` | review |
| `review.modern_comment_done_state_changed` | `modern_comment_done_state_changed` | review |
| `external.include_text_field_target_retargeted` | `external_field_source_changed` | block |
| `external.complex_include_text_field_target_retargeted` | `external_field_source_changed` | block |
| `external.dde_field_source_retargeted` | `external_field_source_changed` | block |
| `binding.document_variable_value_changed` | `document_variable_value_changed` | review |
| `external.attached_template_target_retargeted` | `external_document_dependency_target_changed` | block |
| `external.mail_merge_data_source_target_retargeted` | `mail_merge_data_source_target_changed` | block |
| `review.mail_merge_recipient_active_state_changed` | `mail_merge_recipient_active_state_changed` | block |
| `external.save_through_xslt_target_retargeted` | `save_through_xslt_target_changed` | block |
| `binding.attached_custom_xml_schema_namespace_changed` | `attached_custom_xml_schema_namespace_changed` | review |
| `review.field_recalculation_on_open_enabled` | `field_recalculation_on_open_enabled` | review |
| `review.template_style_update_on_open_enabled` | `template_style_update_on_open_enabled` | review |
| `external.subdocument_target_retargeted` | `external_document_dependency_target_changed` | block |
| `external.frameset_source_target_retargeted` | `external_document_dependency_target_changed` | block |
| `external.vml_linked_ole_object_target_retargeted` | `vml_linked_ole_object_target_changed` | block |
| `external.drawing_linked_picture_target_retargeted` | `drawing_linked_picture_target_changed` | block |
| `import.alternative_format_html_payload_changed` | `alternative_format_import_payload_changed` | block |
| `review.hidden_text_run_added` | `hidden_text_run_added` | review |
| `review.tracked_insertion_markup_added` | `revision_markup_added` | review |
| `review.track_revisions_setting_enabled` | `track_revisions_setting_enabled` | review |
| `review.document_protection_enabled` | `document_protection_enabled` | review |
| `review.permission_range_editor_changed` | `permission_range_editor_changed` | review |
| `binding.data_binding_xpath_retargeted` | `data_binding_mapping_changed` | review |
| `binding.custom_xml_payload_changed` | `custom_xml_payload_changed` | review |
| `macro.vba_project_payload_changed` | `macro_payload_changed` | block |
| `embedded.ole_payload_changed` | `embedded_ole_payload_changed` | block |
| `embedded.activex_control_persistence_payload_changed` | `activex_control_persistence_payload_changed` | block |

`block` and `review` are reference conventions for benchmark scoring, not universal policy advice. A tool may use stricter or looser policy; DCAB scores whether it can report the declared static fact and whether it agrees with the published convention.

## Safety and privacy

All URI-like relationship values use the reserved `example.invalid` domain, and the DDE source is a synthetic local-style string. Macro, embedded-OLE, and ActiveX persistence bytes are inert text markers, not valid executable, OLE, or control payloads. The public truth files deliberately exclude:

- targets, field instructions and fragmented field-code runs, VML shape IDs and target frames, linked-OLE ProgIDs, object IDs, and update modes, ActiveX control names, class IDs, persistence metadata, document-variable names and values, mail-merge recipient hashes and inclusion values, save-through-XSLT anchor and local solution identifiers, attached-custom-XML-schema namespace values, raw `w:linkStyles` serialization values, permission marker IDs and individual editor assignments, task-pane web-extension IDs, classic-comment anchors and paragraph IDs, comment authors, initials, dates and body text, raw `commentsExtended` serialization values, frameset layout/name/size values, references, store descriptors, property values, XPath expressions, relationship IDs, and relationship paths;
- custom XML values, payload bytes, and payload fingerprints;
- protection hashes, salts, passwords, and document content outside the fixed synthetic text.

The structural verifier compares generated bytes, validates ZIP/XML/package invariants, and refuses XML DTD/entity declarations. It does not interpret any stored value beyond the compact fixture contract.

## Why these surfaces

Word uses native OOXML packages, and direct `w:hyperlink` markup can bind its display text to a relationship target. [Microsoft's Open XML API documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hyperlink?view=openxml-3.0.1) shows that relationship-backed form. Word content controls can bind to custom XML data, so mapping and embedded-data changes are meaningful stored review surfaces. [Microsoft documents those bindings](https://learn.microsoft.com/en-us/visualstudio/vsto/content-controls?view=visualstudio), including their relationship to custom XML parts.

Legacy VML shapes are a separate direct-link surface. Microsoft documents a `v:shape`'s [`href` as a hyperlink target](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.shape?view=openxml-3.0.1), and its Word OOXML guidance explains that Word 2007 continued to use VML for shapes and text boxes. DCAB fixes one compact `w:pict`/`v:rect` shape, its target frame, and its styling attributes while changing only the direct `href`. It does not resolve the URL, select a rendering branch, simulate a click, or claim that any client will follow the link.

A VML [`o:OLEObject`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.vml.office.oleobject?view=openxml-3.0.1) is a distinct relationship-backed form. Its `UpdateMode` applies when its `Type` is `Link`, and Microsoft’s [Office compatibility notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/19190ae3-5734-48eb-a919-b89d0b3ce114) describe that type as determining whether the OLE object is included in or stored externally from the package. DCAB fixes one `w:object` carrier, VML placeholder, `Type="Link"`, `ProgID`, shape/object IDs, visual aspect, update mode, and relationship ID while changing only its external standard `oleObject` relationship target. It neither retrieves a source, parses an OLE payload, activates an object, launches an application, nor claims a client will update or display it.

An embedded [`w:control`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.control?view=openxml-3.0.1) can associate a Word `w:object` carrier with an internal control-properties relationship. Microsoft's [`ax:ocx` schema reference](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/48c99072-6cf7-4e69-84b1-3bea64f0ee3a) defines its class/persistence metadata and the binary relationship reference. The [Embedded Control Persistence Binary Data contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/81974a0b-05c0-4c46-b3b6-c96d0b3d3799) specifies the separate internal `activeXControlBinary` relationship from an ActiveX persistence part, and its [compatibility notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/754c8ac1-923d-412f-8fef-c85862e9124a) identify the ActiveX XML content type. DCAB fixes one inline `w:object`/`w:control` anchor, the internal `control` relationship, the `ax:ocx` persistence XML and its class/persistence metadata, both content types, and the required internal persistence-binary relationship while changing only inert synthetic binary bytes. It does not parse the bytes, load or instantiate a control, launch a client/server, render a placeholder, or claim any runtime behavior.

Mail-merge recipient selection is separate from retargeting a data source. Microsoft's [Mail Merge Recipient Data Part contract](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/af3dc913-8c08-4843-ab40-495a92170b96) specifies one internally related recipient-data part from Document Settings, its `w:recipients` root, and no relationships from that part. Its [`w:active` definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/7bbd9fb1-6181-481d-b29c-63842301455d) says a false value excludes the corresponding external record from a merge. DCAB fixes one synthetic external text source, the settings markup, the settings relationships, the content type, the recipient record hash, and package membership while changing only the stored `false`/`true` inclusion state. It does not retrieve or parse the source, identify a real record, perform a merge, or claim client behavior.

Microsoft's [`w:saveThroughXslt` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.savethroughxslt?view=openxml-3.0.1) describes a custom XSL transform used when saving a document as a single XML file, and [`w:useXSLTWhenSaving`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.usexsltwhensaving?view=openxml-3.0.1) controls whether the transform is applied. DCAB fixes the enabled marker, transform anchor, standard external relationship type, relationship ID, and package membership while changing only the synthetic target. It does not fetch, parse, or execute an XSL transform, save a document through one, or make a claim about emitted XML or client behavior.

Microsoft's [`w:attachedSchema` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.attachedschema?view=openxml-3.0.1) specifies a custom XML schema target-namespace association that a host may use when it loads a document if the matching schema is available. DCAB fixes one direct Settings leaf and all package members while changing only its synthetic namespace value. The value is an opaque declaration, not an OPC relationship or a fetch request: the pair does not locate, retrieve, load, or validate against a schema and makes no host-validation claim. Released DocFence 0.29 reports the same-count change through an aggregate-only inventory without exposing the namespace value.

Microsoft's [`w:updateFields` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.updatefieldsonopen?view=openxml-3.0.1) specifies whether fields should have their results recalculated from field codes when a supporting application opens the document. DCAB fixes the package-member set, all stored text, and one direct `CT_OnOff` Settings leaf while changing only its explicit `false`/`true` state. There is no Settings relationship part and no field evaluation in construction or validation. The case does not open Word, parse or evaluate a field instruction, access a field source, follow a link, start an application, or claim that any client will recalculate a field. Released DocFence 0.30 reports the aggregate state transition without evaluating a field.

Microsoft's [`w:linkStyles` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.linkstyles?view=openxml-3.0.1) describes the stored request to update document styles from an attached template when a document opens. DCAB fixes the `w:attachedTemplate` anchor, its external Settings relationship and synthetic target, package-member set, and stored text while changing only the direct `CT_OnOff` `w:linkStyles` leaf from explicit `false` to `true`. It does not resolve, retrieve, or load a template; open Word; propagate styles; or claim that a client will update a style. Released DocFence 0.31 reports the aggregate state transition without exposing the raw setting or template target.

Complex Word fields are a distinct parser boundary from `w:fldSimple`. Microsoft's [`w:fldChar` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fieldchar?view=openxml-3.0.1) defines the required begin/end markers and optional separator, while its [`w:instrText` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.fieldcode?view=openxml-3.0.1) says instruction text is field code only when it occurs in the code portion of a complex field. DCAB fixes one complete `INCLUDETEXT` field with a begin marker, three preserved-whitespace instruction runs that split the keyword itself, a separator, a fixed result, and an end marker. Only the private source fragment changes. It does not resolve or import the source, update/evaluate a field, open Word, or claim a client will process the instruction.

Framesets are a distinct external-document topology. Microsoft's [`w:sourceFileName` documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sourcefilereference?view=openxml-2.20.0) says that a frame source is identified by a relationship in the Web Settings part and requires the standard `frame` relationship type. The OOXML [Framesets contract](https://ooxml.info/docs/11/11.5/) further specifies that each frame target is external. DCAB fixes one content-type override, internal main-document-to-Web-Settings relationship, root frameset layout, frame size/name, source anchor, and relationship ID while changing only the private external target in `word/_rels/webSettings.xml.rels`. The standard says a document with a root frameset is a frameset definition rather than ordinary rendered document content; DCAB does not open, render, resolve, retrieve, import, authenticate to, or claim a client will display a frame.

Office's [web-extension XML format](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-owexml/29f59f30-b835-461a-bd8a-ca400a7bc717) stores Office Add-in structures in Word documents. Microsoft's [auto-open task-pane guidance](https://learn.microsoft.com/en-us/office/dev/add-ins/develop/automatically-open-a-task-pane-with-a-document) specifies internally related `webextension` and `taskpane` parts and the `Office.AutoShowTaskpaneWithDocument` property. DCAB fixes the entire internal topology, task-pane shape, reference, and property name while changing only that property's stored `false`/`true` value. It does not retrieve, install, authenticate, or execute an add-in, and it does not claim that a pane opens: Microsoft documents that the add-in must already be installed, and current auto-open availability depends on deployment/support conditions.

The [`commentsExtended` part](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/31f689cd-4192-4c2d-8d2f-202b1f8f20e9) carries additional information about comments represented by the classic comments part. Microsoft's [`w15:commentEx` schema documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2013.word.commentex?view=openxml-3.0.1) specifies that it is available in Office 2013 and later, ties `paraId` to the associated comment's final paragraph, and defines `done="1"` as a user indication that the comment is done. DCAB fixes one anchored classic comment, its matching paragraph identifier, both internal relationships, and all other extension inventory while changing only the explicit stored `done` value from `0` to `1`. It does not open a Word client, infer an authenticated identity, resolve a comment thread, synchronize with a service, or claim a particular comment UI/state-transition behavior.

An attached template is another relationship-backed setting: [Microsoft's Office Open XML notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7713efa6-b1ff-4cbd-9339-5bf9018433ac) specify that Word obtains its template path through the `attachedTemplate` relationship. A [`w:dataSource`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.datasourcereference?view=openxml-3.0.1) element identifies the external source connected for a mail merge through a `mailMergeSource` relationship. [Microsoft's Word field specification for `DDE`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a2c3a25a-1dba-40da-be7a-47cf63c78d55) defines separate application, source-file, and source-item arguments; DCAB fixes the first and third and changes only the stored source-file argument, without processing the field or starting an application. [`w:docVars`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentvariables?view=openxml-3.0.1) persists document-variable name/value pairs, and Microsoft documents that they can be shown by a [`DOCVARIABLE` field](https://learn.microsoft.com/en-us/office/vba/api/word.variable). DCAB fixes that field reference and variable name while changing only the persisted value, without evaluating a field. [`w:permStart`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.permstart?view=openxml-3.0.1) and [`w:permEnd`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.permend?view=openxml-3.0.1) form paired editable-range permission markup through a shared marker ID. DCAB fixes one paired boundary and its covered stored text while changing only a synthetic individual editor assignment; it does not infer an effective permission or authenticated identity. A [`w:subDoc`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.subdocumentreference?view=openxml-3.0.1) anchor identifies a separate master-document subdocument through an external relationship. A DrawingML [`a:blip` `r:link`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip.link?view=openxml-3.0.1) identifies an image outside the file. A [`w:altChunk`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1) anchor identifies internally stored alternate content for import. DCAB models those static relationship and payload boundaries without resolving or importing them or claiming a client will process them. The other cases follow explicit WordprocessingML constructs: [`w:vanish`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.vanish?view=openxml-3.0.1), [`w:ins`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.insertedrun?view=openxml-3.0.1), Track Changes and protection in [settings](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.settings?view=openxml-3.0.1), and [`w:documentProtection`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentprotection?view=openxml-3.0.1). Document protection is deliberately represented without password material and should not be interpreted as cryptographic protection.

## Tool-neutral observations

An observation report is JSON with schema version `1`. A tool can declare a case `analyzed`, `unsupported`, or `error`. An analyzed case provides zero or more fact objects and an optional `allow`, `review`, or `block` disposition. DCAB validates the envelope before scoring it.

```json
{
  "schema_version": 1,
  "benchmark": {"fixture_schema_versions": [1]},
  "tool": {"name": "example-reviewer", "version": "1.2"},
  "cases": [
    {
      "id": "review.hidden_text_run_added",
      "status": "analyzed",
      "facts": [{"fact": {"kind": "hidden_text_run_added"}}],
      "review": "review"
    }
  ]
}
```

Use `dcab observation-template` for a complete valid skeleton. Extra facts remain visible under `unrecognized_facts`; DCAB does not label them false positives because its oracle is intentionally partial.

## Optional DocFence adapter

The optional local adapter translates public, aggregate [DocFence](https://github.com/SybilGambleyyu/docfence) reports into DCAB observations. DCAB does not depend on DocFence and never receives the tool's private signatures, targets, or payload data.

```bash
dcab docfence-observations --executable docfence --output docfence-observations.json
dcab score --observations docfence-observations.json --strict
```

The adapter is evidence of one independent consumer, not a claim that a single tool defines the benchmark.

## Compatibility boundary

The test suite opens every `.docx` pair with `python-docx` and opens all `.docx`/`.docm` packages using its lower-level OPC reader. This is useful package-level interoperability evidence, not a claim of compatibility with every Word, Office, LibreOffice, renderer, macro engine, OLE client, or field-update behavior.

## Research and limits

[RESEARCH.md](RESEARCH.md) records the scoped gap investigation and adjacent work. DCAB does not claim that no other document benchmark or comparison tool exists. Its specific contribution is a compact, reproducible, paired static-review contract with target-free public truth.

## Development

```bash
python -m venv .venv
.venv/bin/python -m pip install -e '.[dev]'
.venv/bin/python -m pytest
.venv/bin/python -m ruff check .
.venv/bin/python -m ruff format --check .
```

MIT licensed. Contributions should preserve deterministic generation, the static/nonexecuting boundary, and target-free truth files.
