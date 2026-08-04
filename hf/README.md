---
license: mit
task_categories:
- other
tags:
- ooxml
- docx
- docm
- wordprocessingml
- benchmark
- static-analysis
- privacy
pretty_name: Document Change Assurance Benchmark
size_categories:
- n<1K
---

# Document Change Assurance Benchmark (DCAB)

This dataset mirrors the deterministic fixture corpus from [DCAB v0.21.0](https://github.com/SybilGambleyyu/document-change-benchmark/tree/v0.21.0). It contains 32 paired synthetic WordprocessingML cases for static document-change assurance.

Each case directory provides a baseline package, candidate package, and target-free `truth.json`. The public truth files describe only a narrow fact category and a reference review convention. They never disclose URI-like targets, field instructions or fragmented field-code runs, VML shape IDs or target frames, linked-OLE ProgIDs/object IDs/update modes, ActiveX control names/class IDs/persistence metadata, document-variable names or values, mail-merge recipient hashes or inclusion values, save-through-XSLT anchors or local solution identifiers, attached-custom-XML-schema namespace values, raw `w:linkStyles` serialization values, permission marker IDs or individual editor assignments, task-pane web-extension IDs, classic-comment anchors and paragraph IDs, comment author/initial/date/body values, raw `commentsExtended` serialization values, frameset layout/name/size values, references, store descriptors, property values, XPath values, relationship IDs, custom XML values, or opaque payload bytes.

## Layout

```text
fixtures/
  manifest.jsonl
  interaction.word_hyperlink_target_retargeted/
    baseline.docx
    candidate.docx
    truth.json
  ...
```

The corpus covers direct Word and legacy VML shape hyperlinks, `HYPERLINK`, simple and fragmented complex `INCLUDETEXT`, `DDE`, and `DOCVARIABLE` fields, persisted document variables, attached custom XML schema declarations, automatic field-recalculation-on-open and template-style-update-on-open settings, editable-range permission markup, task-pane Office web-extension auto-show configuration, Office 2013 `commentsExtended` done metadata, attached-template, mail-merge data-source/recipient-selection, and save-through-XSLT settings, master-subdocument and frameset-source dependencies, legacy VML linked-OLE objects, DrawingML linked-picture relationships, alternative-format import payloads, hidden text, tracked insertion markup, Track Changes and document protection settings, content-control/custom-XML bindings, VBA project payload boundaries, embedded OLE payload boundaries, and ActiveX control-persistence payload boundaries.

## Safety boundary

URI-like relationship values in packages use `example.invalid`, and the DDE source is a synthetic local-style string. Macro, OLE, and ActiveX persistence bytes are inert synthetic marker data. The corpus is static: use it without resolving a relationship, loading an attached template, updating a field or style, opening a Word client, parsing an opaque payload, activating OLE or ActiveX, or executing code.

## Reproducibility

The source repository includes a deterministic builder and independent structural verifier:

```bash
python -m pip install document-change-benchmark
dcab validate --fixtures fixtures
```

This mirror has no special execution requirement. It is provided under the MIT license; see `LICENSE`.

## Scope

DCAB does not claim client rendering/runtime compatibility or universal security policy. It is a narrow, tool-neutral static-review benchmark. In particular, the VML pair does not resolve or follow its direct link; the VML linked-OLE pair does not retrieve a source, parse an OLE payload, activate an object, launch an application, or assert a client update/display behavior; the ActiveX pair does not parse its persistence bytes, load or instantiate a control, render a placeholder, invoke a client/server, or assert runtime behavior; the mail-merge recipient-selection pair does not retrieve or parse a source, identify a record, perform a merge, or assert client behavior; the save-through-XSLT pair does not retrieve, parse, or execute a transform, save through one, or assert emitted XML/client behavior; the attached-schema pair does not locate, retrieve, load, or validate against a schema or assert host validation behavior; the field-recalculation-on-open pair does not open Word, parse/evaluate or update a field, access a source, follow a link, start an application, or assert host behavior; the template-style-update-on-open pair retains its attached-template anchor, relationship, and target but does not resolve, retrieve, or load the template, open Word, propagate styles, or assert client behavior; the fragmented complex-field pair does not resolve/import a source, update/evaluate a field, open Word, or assert client processing; the frameset pair does not open, render, resolve, retrieve, import, or authenticate to a frame source or assert client display; the task-pane pair does not install, retrieve, authenticate, or execute an add-in or assert that a pane opens; and the `commentsExtended` pair does not open Word, resolve a thread, infer an identity, synchronize with a service, or assert a client comment UI behavior. See the [repository README](https://github.com/SybilGambleyyu/document-change-benchmark) and [research notes](https://github.com/SybilGambleyyu/document-change-benchmark/blob/v0.21.0/RESEARCH.md) for contract details and limits.
