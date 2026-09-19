---
"@stll/folio-core": patch
---

`w:numId w:val="0"` on a paragraph style is ECMA-376's "no numbering" sentinel, not a reference: it switches off the numbering the style would inherit through `w:basedOn`. Saving a document whose styles carry it (a TOC heading based on a numbered heading, for example) no longer fails with "Style references missing numbering definition 0", and the sentinel is written back unchanged. Every numbering lookup now reads the sentinel through one shared predicate.
