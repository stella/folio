---
"@stll/folio-core": patch
---

The block manifest reads a bookmark that ends between cells, rows or blocks (outside any paragraph) as ending behind the paragraph read before it, so a cross-reference field to a numbered table row (a bookmark started in the row's first cell and ended in the row) reports the paragraph it marks.
