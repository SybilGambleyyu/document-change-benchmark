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

This dataset mirrors the deterministic fixture corpus from [DCAB v0.30.0](https://github.com/SybilGambleyyu/document-change-benchmark/tree/v0.30.0). It contains 41 paired synthetic WordprocessingML cases for static document-change assurance.

Each case directory provides a baseline package, candidate package, and target-free `truth.json`. The public truth files describe only a narrow fact category and a reference review convention. They never disclose URI-like targets, field instructions or fragmented field-code runs, VML shape IDs or target frames, linked-OLE ProgIDs/object IDs/update modes, DrawingML nonvisual object names/descriptions/IDs, raw hidden serialization values, graphic-data URIs, ActiveX control names/class IDs/persistence metadata, document-variable names or values, mail-merge recipient hashes or inclusion values, save-through-XSLT anchors or local solution identifiers, attached-custom-XML-schema namespace values, raw `w:linkStyles`, `w:removePersonalInformation`, `w:savePreviewPicture`, or content-control `w:lock` serialization values, content-control IDs/tags/text, permission marker IDs or individual editor assignments, task-pane web-extension IDs, classic-comment anchors and paragraph IDs, comment author/initial/date/body values, raw `commentsExtended` serialization values, frameset layout/name/size values, package-signature values, manifest object IDs, reference URIs, relationship selectors, signature/origin paths, thumbnail relationship sources/targets, content types, part paths, references, store descriptors, property values, XPath values, relationship IDs, custom XML values, thumbnail image bytes, Markup Compatibility branch bodies, feature-prefix or qualified-name values, compatibility-rule values, or opaque payload bytes.

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

The corpus covers direct Word and legacy VML shape hyperlinks, `HYPERLINK`, simple and fragmented complex `INCLUDETEXT`, `DDE`, `DOCVARIABLE`, and fixed legacy `FORMTEXT` fields, persisted document variables, attached custom XML schema declarations, automatic field-recalculation-on-open, template-style-update-on-open, personal-information-removal-on-save, form-data-only-save, preview-thumbnail-on-save, and content-control-lock declarations, editable-range permission markup, task-pane Office web-extension auto-show configuration, Office 2013 `commentsExtended` done metadata, attached-template, mail-merge data-source/recipient-selection, and save-through-XSLT settings, master-subdocument and frameset-source dependencies, legacy VML linked-OLE objects, DrawingML linked-picture relationships and direct nonvisual visibility declarations, alternative-format import payloads, hidden text, tracked insertion markup, Track Changes and document protection settings, bound and unbound custom-XML payloads, relationship-bound OPC package-thumbnail payloads, static OPC package-signature declaration coverage, OOXML Markup Compatibility choice requirements, VBA project payload boundaries, embedded OLE payload boundaries, and ActiveX control-persistence payload boundaries.

## Safety boundary

URI-like relationship values in packages use `example.invalid`, and the DDE source is a synthetic local-style string. Macro, OLE, and ActiveX persistence bytes are inert synthetic marker data. The package-thumbnail pair uses deterministic fully synthetic 1×1 PNG bytes and treats them as opaque. The corpus is static: use it without resolving a relationship, loading an attached template, updating a field or style, opening a Word client, parsing or decoding an opaque payload, activating OLE or ActiveX, or executing code.

The unbound custom-XML pair retains a conventional custom-XML data/properties topology and all stored Word text while changing only inert XML bytes. It contains no `w:dataBinding` marker and makes no claim that Word displays, uses, or removes the stored XML.

The package-thumbnail pair retains the standard root thumbnail relationship, `image/png` content type, member set, and all stored Word text while changing only a fully synthetic PNG's bytes. It does not decode the image or claim that it previews the document or will be displayed by any client.

The Markup Compatibility pair retains one `mc:AlternateContent`, one `mc:Choice`, one `mc:Fallback`, its package member set, and all stored Word text while changing only a private `Choice/@Requires` feature prefix. It does not validate MCE conformance, resolve the prefix, choose a branch, preprocess or save a package, or claim any client rendering behavior.

The DrawingML visibility pair retains one compact inline DrawingML carrier, its package topology, and all stored Word text while changing only direct `wp:docPr/@hidden` from `false` to `true`. It does not identify a visual object, calculate effective visibility, choose an MCE branch, lay out or render DrawingML, or claim client behavior.

The form-data-only-save pair retains one fixed legacy `FORMTEXT` carrier, its package topology, and all stored Word text while changing only direct `w:saveFormsData/@w:val` from `false` to `true`. It does not read or evaluate a form-field value, open Word, save a document, emit a delimited record, determine a delimiter, or claim client behavior.

The preview-thumbnail-on-save pair retains no thumbnail relationship or image part, its package topology, and all stored Word text while changing only direct `w:savePreviewPicture/@w:val` from `false` to `true`. It does not create, decode, render, or classify an image; open Word; save a document; generate a thumbnail; or claim client behavior.

The content-control-lock pair retains one direct SDT carrier, its fixed ID/tag and stored text, package topology, and all other stored Word text while changing only direct `w:sdtPr/w:lock/@w:val` from `unlocked` to `sdtContentLocked`. It does not infer a lock from omitted markup, classify a control type, read a control value, open Word, apply a lock, or claim client behavior.

The package-signature pair retains its signature-origin topology, package
membership, and stored Word text while changing only one private static
relationship-selection declaration in an XMLDSIG-shaped manifest. Its digest
and signature values are fixed synthetic placeholders: it exercises declared
coverage only and does not establish cryptographic validity, signer identity,
trust, integrity, or a client-side signature result.

## Reproducibility

The source repository includes a deterministic builder and independent structural verifier:

```bash
python -m pip install document-change-benchmark
dcab validate --fixtures fixtures
```

This mirror has no special execution requirement. It is provided under the MIT license; see `LICENSE`.

## Scope

DCAB does not claim client rendering/runtime compatibility or universal security policy. It is a narrow, tool-neutral static-review benchmark. In particular, it does not resolve or follow links; retrieve sources, templates, frame documents, or mail-merge data; parse or decode opaque macro, thumbnail, OLE, ActiveX, or alternative-format payloads; interpret unbound custom-XML data; resolve an MCE feature prefix or choose an MCE branch; calculate effective DrawingML visibility or content-control locking; calculate or validate a package signature, digest, certificate, trust, or integrity result; update/evaluate fields; read a form/control value; open Word; save a document; emit a form-data record; generate a thumbnail; activate objects; execute code; synchronize comments or task panes; or assert client behavior. See the [repository README](https://github.com/SybilGambleyyu/document-change-benchmark) and [v0.30 research notes](https://github.com/SybilGambleyyu/document-change-benchmark/blob/v0.30.0/RESEARCH.md) for the exact contracts and limits.
