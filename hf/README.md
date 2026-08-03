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

This dataset mirrors the deterministic fixture corpus from [DCAB v0.5.0](https://github.com/SybilGambleyyu/document-change-benchmark/tree/v0.5.0). It contains 16 paired synthetic WordprocessingML cases for static document-change assurance.

Each case directory provides a baseline package, candidate package, and target-free `truth.json`. The public truth files describe only a narrow fact category and a reference review convention. They never disclose URI-like targets, field instructions, XPath values, relationship IDs, custom XML values, or opaque payload bytes.

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

The corpus covers direct Word hyperlinks, `HYPERLINK` and `INCLUDETEXT` fields, attached-template, master-subdocument, and DrawingML linked-picture relationships, alternative-format import payloads, hidden text, tracked insertion markup, Track Changes and document protection settings, content-control/custom-XML bindings, VBA project payload boundaries, and embedded OLE payload boundaries.

## Safety boundary

All URI-like values in packages use `example.invalid`. Macro and OLE bytes are inert synthetic marker data. The corpus is static: use it without resolving a relationship, updating a field, opening a Word client, parsing an opaque payload, activating OLE, or executing code.

## Reproducibility

The source repository includes a deterministic builder and independent structural verifier:

```bash
python -m pip install document-change-benchmark
dcab validate --fixtures fixtures
```

This mirror has no special execution requirement. It is provided under the MIT license; see `LICENSE`.

## Scope

DCAB does not claim client rendering/runtime compatibility or universal security policy. It is a narrow, tool-neutral static-review benchmark. See the [repository README](https://github.com/SybilGambleyyu/document-change-benchmark) and [research notes](https://github.com/SybilGambleyyu/document-change-benchmark/blob/v0.5.0/RESEARCH.md) for contract details and limits.
