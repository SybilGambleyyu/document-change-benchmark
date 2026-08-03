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

This dataset mirrors the deterministic fixture corpus from [DCAB v0.11.0](https://github.com/SybilGambleyyu/document-change-benchmark/tree/v0.11.0). It contains 22 paired synthetic WordprocessingML cases for static document-change assurance.

Each case directory provides a baseline package, candidate package, and target-free `truth.json`. The public truth files describe only a narrow fact category and a reference review convention. They never disclose URI-like targets, field instructions, VML shape IDs or target frames, document-variable names or values, permission marker IDs or individual editor assignments, task-pane web-extension IDs, references, store descriptors, property values, XPath values, relationship IDs, custom XML values, or opaque payload bytes.

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

The corpus covers direct Word and legacy VML shape hyperlinks, `HYPERLINK`, `INCLUDETEXT`, `DDE`, and `DOCVARIABLE` fields, persisted document variables, editable-range permission markup, task-pane Office web-extension auto-show configuration, attached-template and mail-merge data-source settings, master-subdocument, and DrawingML linked-picture relationships, alternative-format import payloads, hidden text, tracked insertion markup, Track Changes and document protection settings, content-control/custom-XML bindings, VBA project payload boundaries, and embedded OLE payload boundaries.

## Safety boundary

URI-like relationship values in packages use `example.invalid`, and the DDE source is a synthetic local-style string. Macro and OLE bytes are inert synthetic marker data. The corpus is static: use it without resolving a relationship, updating a field, opening a Word client, parsing an opaque payload, activating OLE, or executing code.

## Reproducibility

The source repository includes a deterministic builder and independent structural verifier:

```bash
python -m pip install document-change-benchmark
dcab validate --fixtures fixtures
```

This mirror has no special execution requirement. It is provided under the MIT license; see `LICENSE`.

## Scope

DCAB does not claim client rendering/runtime compatibility or universal security policy. It is a narrow, tool-neutral static-review benchmark. In particular, the VML pair does not resolve or follow its direct link, and the task-pane pair does not install, retrieve, authenticate, or execute an add-in or assert that a pane opens. See the [repository README](https://github.com/SybilGambleyyu/document-change-benchmark) and [research notes](https://github.com/SybilGambleyyu/document-change-benchmark/blob/v0.11.0/RESEARCH.md) for contract details and limits.
