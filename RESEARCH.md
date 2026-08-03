# Research notes: a paired WordprocessingML change-review gap

Research was performed on 2026-08-03 before defining DCAB. The conclusion is scoped: there are mature document comparison products and a growing set of office/document benchmarks, but the reviewed material did not provide an open, paired, target-free corpus and tool-neutral scorer for static WordprocessingML change assurance.

## Adjacent work is valuable but measures a different task

- [DocBank](https://arxiv.org/abs/2006.01038) is a large document-layout-analysis benchmark. Its token/layout labels solve a different image/document-understanding problem.
- [Office Comprehension Bench](https://arxiv.org/abs/2607.01245) evaluates understanding over native office files. It is a question-answering/comprehension benchmark rather than a paired stored-change contract.
- Commercial Word comparison tools can produce text/format deltas and native tracked changes. That is useful review functionality, but it is not an open scored corpus for static package surfaces such as relationship targets, field instructions, custom-XML bindings, or opaque payload boundaries.

None of that implies absence of other relevant work. It identifies a concrete interoperability/evaluation niche: a reviewer can detect some stored evidence, yet cannot readily report comparable coverage across tools without sharing sensitive targets or document content.

## Evidence that the selected surfaces are real stored semantics

- [`w:hyperlink`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.hyperlink?view=openxml-3.0.1) can identify a relationship-backed hyperlink target while keeping display text in the document markup.
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

The initial release intentionally does not claim coverage for every WordprocessingML surface. Candidate future work includes linked pictures, legacy VML hyperlinks, mail merge, external document dependencies, ActiveX controls, document variables, permission ranges, modern comments, task-pane web extensions, and style-resolution semantics. Each requires the same standard: a narrowly defined static fact, safe synthetic pair generation, a target-free oracle, and a credible independent reader/consumer test.
