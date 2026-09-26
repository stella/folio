---
"@stll/folio-core": patch
---

A list command or list autoformat in a document with no list of the requested kind now defines a numbering instance of that kind (and the numbering part when the package has none) instead of referencing one the package does not define, so the document saves. A bullet command no longer joins a numbered list, or the reverse. Selective save falls back to a full save when it cannot write the numbering the paragraphs reference.
