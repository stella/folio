---
"@stll/folio-core": minor
---

Add `docxToMarkdown(bytes)` to the `/server` entry: a one-call, server-safe DOCX
bytes → markdown converter that composes `parseDocx` (font preloading disabled)
and `toMarkdown`, so non-browser callers get full DOCX fidelity without deep
imports or a hand-rolled OOXML walker. Also fixes the DOCX table-cell parser to
descend into block-level content controls (`w:tc > w:sdt > w:sdtContent`), so
controlled field text inside table cells is no longer dropped.
