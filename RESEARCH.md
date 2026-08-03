# Research notes: a paired WordprocessingML change-review gap

Research was performed on 2026-08-03 before defining DCAB. The conclusion is scoped: there are mature document comparison products and a growing set of office/document benchmarks, but the reviewed material did not provide an open, paired, target-free corpus and tool-neutral scorer for static WordprocessingML change assurance.

## Adjacent work is valuable but measures a different task

- [DocBank](https://arxiv.org/abs/2006.01038) is a large document-layout-analysis benchmark. Its token/layout labels solve a different image/document-understanding problem.
- [Office Comprehension Bench](https://arxiv.org/abs/2607.01245) evaluates understanding over native office files. It is a question-answering/comprehension benchmark rather than a paired stored-change contract.
- Commercial Word comparison tools can produce text/format deltas and native tracked changes. That is useful review functionality, but it is not an open scored corpus for static package surfaces such as relationship targets, field instructions, custom-XML bindings, or opaque payload boundaries.

None of that implies absence of other relevant work. It identifies a concrete interoperability/evaluation niche: a reviewer can detect some stored evidence, yet cannot readily report comparable coverage across tools without sharing sensitive targets or document content.

## Evidence that the selected surfaces are real stored semantics

- [`w:hyperlink`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hyperlink?view=openxml-3.0.1) can identify a relationship-backed hyperlink target while keeping display text in the document markup.
- An [`attachedTemplate` relationship](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/7713efa6-b1ff-4cbd-9339-5bf9018433ac) supplies the attached template path for a Word settings anchor. DCAB keeps the anchor and relationship ID stable while changing only that external relationship's synthetic target.
- A [`w:dataSource`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.datasourcereference?view=openxml-3.0.1) mail-merge setting identifies the external source through a required `mailMergeSource` relationship. DCAB keeps the `w:mailMerge`/`w:dataSource` markup and relationship ID stable while changing only that synthetic external target.
- [Microsoft's Word field specification for `DDE`](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/a2c3a25a-1dba-40da-be7a-47cf63c78d55) defines separate application, source-file location, and source-item arguments for information linked from another application. DCAB keeps the synthetic application and item arguments fixed while changing only a synthetic local-style source-file argument in a complete simple field. [Microsoft's DDE security advisory](https://learn.microsoft.com/en-us/security-updates/securityadvisories/2017/4053440) also documents Office controls for processing DDE fields; that supports review relevance but does not make DCAB a runtime or exploit benchmark.
- A [`w:subDoc`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.subdocumentreference?view=openxml-3.0.1) identifies a master-document subdocument location. Its relationship is a separately stored external document dependency, so DCAB keeps the anchor and relationship ID stable while changing only the synthetic target.
- A DrawingML [`a:blip` `r:link`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.blip.link?view=openxml-3.0.1) identifies a linked picture that does not reside in the file. DCAB keeps that marker and its external image relationship type stable while changing only the synthetic relationship target.
- A [`w:altChunk`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.altchunk?view=openxml-3.0.1) anchor identifies alternate content for import through an internal `afChunk` relationship. DCAB keeps the anchor and relationship fixed while changing only an inert HTML payload, and never parses or imports it.
- [Word content controls](https://learn.microsoft.com/en-us/visualstudio/vsto/content-controls?view=visualstudio) can bind content controls to embedded custom XML parts. [Microsoft's binding walkthrough](https://learn.microsoft.com/en-us/visualstudio/vsto/walkthrough-binding-content-controls-to-custom-xml-parts?view=visualstudio) explains that mapped values can display when a document opens.
- [`w:vanish`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.vanish?view=openxml-3.0.1) and [`w:ins`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.insertedrun?view=openxml-3.0.1) are explicit WordprocessingML markup classes.
- [`w:documentProtection`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.documentprotection?view=openxml-3.0.1) records editing restrictions but is documented as not being a security feature; DCAB therefore models only a password-free declaration and makes no security claim.
- [Word OOXML documentation](https://learn.microsoft.com/en-us/office/dev/add-ins/word/create-better-add-ins-for-word-with-office-open-xml) describes OOXML as Word's native file format and demonstrates structured document tag markup.

## Design implications

1. The oracle must be deliberately partial and target-free. A benchmark should not force tools to disclose URL-like destinations, field arguments, data-binding XPath values, or opaque payload fingerprints.
2. Fixture packages should be generated from source rather than copied from user documents. This makes byte-level regeneration and public inspection possible.
3. The corpus should be static. It must not download a resource, update a field, activate OLE, or execute a macro in construction, validation, or scoring.
4. Compatibility should be demonstrated separately from runtime claims. DCAB's independent `python-docx` checks prove an OPC/reader boundary, not Word rendering or behavior.
5. A tool-neutral observation schema is more useful than a benchmark coupled to one scanner. The optional DocFence adapter is a consumer, not the specification.

## Deferred surfaces

The initial releases intentionally do not claim coverage for every WordprocessingML surface. Candidate future work includes legacy VML hyperlinks, mail merge, remaining external document dependencies such as frameset sources, ActiveX controls, document variables, permission ranges, modern comments, task-pane web extensions, and style-resolution semantics. Each requires the same standard: a narrowly defined static fact, safe synthetic pair generation, a target-free oracle, and a credible independent reader/consumer test.

A direct external `w:objectLink` relationship was investigated and intentionally not added. The OOXML definition identifies `w:objectLink` as linked-object metadata but specifies that its `r:id` relationship targets an Embedded Object Part. A bare external target is therefore not a clean standards-conformance contract for this corpus. A future linked-OLE case would need a valid, inert embedded link payload and a separate safe static parser for that payload; it must not shortcut that requirement by treating an external relationship alone as equivalent.
