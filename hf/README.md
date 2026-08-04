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

This dataset mirrors the deterministic fixture corpus from [DCAB v0.26.0](https://github.com/SybilGambleyyu/document-change-benchmark/tree/v0.26.0). It contains 37 paired synthetic WordprocessingML cases for static document-change assurance.

Each case directory provides a baseline package, candidate package, and target-free `truth.json`. The public truth files describe only a narrow fact category and a reference review convention. They never disclose URI-like targets, field instructions or fragmented field-code runs, VML shape IDs or target frames, linked-OLE ProgIDs/object IDs/update modes, DrawingML nonvisual object names/descriptions/IDs, raw hidden serialization values, graphic-data URIs, ActiveX control names/class IDs/persistence metadata, document-variable names or values, mail-merge recipient hashes or inclusion values, save-through-XSLT anchors or local solution identifiers, attached-custom-XML-schema namespace values, raw `w:linkStyles` or `w:removePersonalInformation` serialization values, permission marker IDs or individual editor assignments, task-pane web-extension IDs, classic-comment anchors and paragraph IDs, comment author/initial/date/body values, raw `commentsExtended` serialization values, frameset layout/name/size values, thumbnail relationship sources/targets, content types, part paths, references, store descriptors, property values, XPath values, relationship IDs, custom XML values, thumbnail image bytes, Markup Compatibility branch bodies, feature-prefix or qualified-name values, compatibility-rule values, or opaque payload bytes.

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

The corpus covers direct Word and legacy VML shape hyperlinks, `HYPERLINK`, simple and fragmented complex `INCLUDETEXT`, `DDE`, and `DOCVARIABLE` fields, persisted document variables, attached custom XML schema declarations, automatic field-recalculation-on-open, template-style-update-on-open, and personal-information-removal-on-save settings, editable-range permission markup, task-pane Office web-extension auto-show configuration, Office 2013 `commentsExtended` done metadata, attached-template, mail-merge data-source/recipient-selection, and save-through-XSLT settings, master-subdocument and frameset-source dependencies, legacy VML linked-OLE objects, DrawingML linked-picture relationships and direct nonvisual visibility declarations, alternative-format import payloads, hidden text, tracked insertion markup, Track Changes and document protection settings, bound and unbound custom-XML payloads, relationship-bound OPC package-thumbnail payloads, OOXML Markup Compatibility choice requirements, VBA project payload boundaries, embedded OLE payload boundaries, and ActiveX control-persistence payload boundaries.

## Safety boundary

URI-like relationship values in packages use `example.invalid`, and the DDE source is a synthetic local-style string. Macro, OLE, and ActiveX persistence bytes are inert synthetic marker data. The package-thumbnail pair uses deterministic fully synthetic 1×1 PNG bytes and treats them as opaque. The corpus is static: use it without resolving a relationship, loading an attached template, updating a field or style, opening a Word client, parsing or decoding an opaque payload, activating OLE or ActiveX, or executing code.

The unbound custom-XML pair retains a conventional custom-XML data/properties topology and all stored Word text while changing only inert XML bytes. It contains no `w:dataBinding` marker and makes no claim that Word displays, uses, or removes the stored XML.

The package-thumbnail pair retains the standard root thumbnail relationship, `image/png` content type, member set, and all stored Word text while changing only a fully synthetic PNG's bytes. It does not decode the image or claim that it previews the document or will be displayed by any client.

The Markup Compatibility pair retains one `mc:AlternateContent`, one `mc:Choice`, one `mc:Fallback`, its package member set, and all stored Word text while changing only a private `Choice/@Requires` feature prefix. It does not validate MCE conformance, resolve the prefix, choose a branch, preprocess or save a package, or claim any client rendering behavior.

The DrawingML visibility pair retains one compact inline DrawingML carrier, its package topology, and all stored Word text while changing only direct `wp:docPr/@hidden` from `false` to `true`. It does not identify a visual object, calculate effective visibility, choose an MCE branch, lay out or render DrawingML, or claim client behavior.

## Reproducibility

The source repository includes a deterministic builder and independent structural verifier:

```bash
python -m pip install document-change-benchmark
dcab validate --fixtures fixtures
```

This mirror has no special execution requirement. It is provided under the MIT license; see `LICENSE`.

## Scope

DCAB does not claim client rendering/runtime compatibility or universal security policy. It is a narrow, tool-neutral static-review benchmark. In particular, it does not resolve or follow links; retrieve sources, templates, frame documents, or mail-merge data; parse or decode opaque macro, thumbnail, OLE, ActiveX, or alternative-format payloads; interpret unbound custom-XML data; resolve an MCE feature prefix or choose an MCE branch; calculate effective DrawingML visibility; update/evaluate fields; open Word; activate objects; execute code; synchronize comments or task panes; or assert client behavior. See the [repository README](https://github.com/SybilGambleyyu/document-change-benchmark) and [v0.26 research notes](https://github.com/SybilGambleyyu/document-change-benchmark/blob/v0.26.0/RESEARCH.md) for the exact contracts and limits.
