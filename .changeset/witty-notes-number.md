---
"@stll/folio-core": patch
---

Footnote and endnote body reference marks now display the sequential reference-order number (1, 2, 3…) instead of the raw `w:id`, matching Word for documents with non-contiguous or out-of-order note ids. The body marker and the footnote-area number derive from one shared display-number map, and reserved notes (separators, continuation notices at positive ids) no longer shift numbering.
