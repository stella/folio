---
"@stll/folio-core": minor
---

`compareDocx` reports a changed numbering definition as a `numbering` change:
a list whose format, level template or start differs moves every label in the
list and no block's text, so a text comparison saw two identical documents.
Reported and not represented — OOXML has no tracked-change grammar for
`numbering.xml`, and Word does not track it either.

`FolioDocxReviewer.readNumberingDefinitions()` returns the package's numbering
flattened to one entry per instance and level, with overrides resolved.
