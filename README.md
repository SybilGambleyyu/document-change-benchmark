# Document Change Assurance Benchmark (DCAB)

DCAB is an open, deterministic corpus for evaluating static review tools that compare WordprocessingML packages. It supplies nineteen paired synthetic `.docx`/`.docm` fixtures, a privacy-safe public oracle, an observation schema, and a scorer.

It is for a question ordinary text diffs do not answer well: did a document change in a stored review-sensitive surface even when ordinary text is unchanged? The corpus covers direct hyperlinks and field instructions, DDE field sources and persisted document variables, attached-template, mail-merge data-source, master-subdocument, and DrawingML linked-picture relationships, alternative-format import payloads, hidden text and revision markup, review settings and document protection, content-control bindings and custom XML, plus opaque macro and embedded-OLE payload boundaries.

DCAB is not a Word renderer, a macro scanner, a field evaluator, or a runtime behavior benchmark. It never resolves an external target, opens Word, updates a field, parses an opaque payload, or executes code.

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

`manifest.jsonl` catalogues the nineteen cases. Every pair has the same package-member set, differs only at a declared member boundary, and retains the same sequence of stored `w:t` values. That invariant is intentionally narrower than visual or client-runtime equivalence.

Version 0.8 adds a case but retains fixture schema version 1: the truth and observation envelopes are unchanged. An earlier observation can still be parsed, but it is incomplete when scored against this nineteen-case catalogue.

| Case | Declared fact | Reference convention |
| --- | --- | --- |
| `interaction.word_hyperlink_target_retargeted` | `word_hyperlink_target_changed` | block |
| `interaction.word_hyperlink_added` | `word_hyperlink_added` | block |
| `interaction.word_hyperlink_field_target_retargeted` | `field_target_changed` | block |
| `external.include_text_field_target_retargeted` | `external_field_source_changed` | block |
| `external.dde_field_source_retargeted` | `external_field_source_changed` | block |
| `binding.document_variable_value_changed` | `document_variable_value_changed` | review |
| `external.attached_template_target_retargeted` | `external_document_dependency_target_changed` | block |
| `external.mail_merge_data_source_target_retargeted` | `mail_merge_data_source_target_changed` | block |
| `external.subdocument_target_retargeted` | `external_document_dependency_target_changed` | block |
| `external.drawing_linked_picture_target_retargeted` | `drawing_linked_picture_target_changed` | block |
| `import.alternative_format_html_payload_changed` | `alternative_format_import_payload_changed` | block |
| `review.hidden_text_run_added` | `hidden_text_run_added` | review |
| `review.tracked_insertion_markup_added` | `revision_markup_added` | review |
| `review.track_revisions_setting_enabled` | `track_revisions_setting_enabled` | review |
| `review.document_protection_enabled` | `document_protection_enabled` | review |
| `binding.data_binding_xpath_retargeted` | `data_binding_mapping_changed` | review |
| `binding.custom_xml_payload_changed` | `custom_xml_payload_changed` | review |
| `macro.vba_project_payload_changed` | `macro_payload_changed` | block |
| `embedded.ole_payload_changed` | `embedded_ole_payload_changed` | block |

`block` and `review` are reference conventions for benchmark scoring, not universal policy advice. A tool may use stricter or looser policy; DCAB scores whether it can report the declared static fact and whether it agrees with the published convention.

## Safety and privacy

All URI-like relationship values use the reserved `example.invalid` domain, and the DDE source is a synthetic local-style string. Macro and embedded-object bytes are inert text markers, not valid executable/OLE payloads. The public truth files deliberately exclude:

- targets, field instructions, document-variable names and values, XPath expressions, relationship IDs, and relationship paths;
- custom XML values, payload bytes, and payload fingerprints;
- protection hashes, salts, passwords, and document content outside the fixed synthetic text.

The structural verifier compares generated bytes, validates ZIP/XML/package invariants, and refuses XML DTD/entity declarations. It does not interpret any stored value beyond the compact fixture contract.

## Why these surfaces

Word uses native OOXML packages, and direct `w:hyperlink` markup can bind its display text to a relationship target. [Microsoft's Open XML API documentation](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hyperlink?view=openxml-3.0.1) shows that relationship-backed form. Word content controls can bind to custom XML data, so mapping and embedded-data changes are meaningful stored review surfaces. [Microsoft documents those bindings](https://learn.microsoft.com/en-us/visualstudio/vsto/content-controls?view=visualstudio), including their relationship to custom XML parts.

An attached template is another relationship-backed setting: [Microsoft's Office Open XML notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7713efa6-b1ff-4cbd-9339-5bf9018433ac) specify that Word obtains its template path through the `attachedTemplate` relationship. A [`w:dataSource`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.datasourcereference?view=openxml-3.0.1) element identifies the external source connected for a mail merge through a `mailMergeSource` relationship. [Microsoft's Word field specification for `DDE`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a2c3a25a-1dba-40da-be7a-47cf63c78d55) defines separate application, source-file, and source-item arguments; DCAB fixes the first and third and changes only the stored source-file argument, without processing the field or starting an application. [`w:docVars`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentvariables?view=openxml-3.0.1) persists document-variable name/value pairs, and Microsoft documents that they can be shown by a [`DOCVARIABLE` field](https://learn.microsoft.com/en-us/office/vba/api/word.variable). DCAB fixes that field reference and variable name while changing only the persisted value, without evaluating a field. A [`w:subDoc`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.subdocumentreference?view=openxml-3.0.1) anchor identifies a separate master-document subdocument through an external relationship. A DrawingML [`a:blip` `r:link`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip.link?view=openxml-3.0.1) identifies an image outside the file. A [`w:altChunk`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1) anchor identifies internally stored alternate content for import. DCAB models those static relationship and payload boundaries without resolving or importing them or claiming a client will process them. The other cases follow explicit WordprocessingML constructs: [`w:vanish`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.vanish?view=openxml-3.0.1), [`w:ins`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.insertedrun?view=openxml-3.0.1), Track Changes and protection in [settings](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.settings?view=openxml-3.0.1), and [`w:documentProtection`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentprotection?view=openxml-3.0.1). Document protection is deliberately represented without password material and should not be interpreted as cryptographic protection.

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
