---
"@stll/folio-core": patch
---

A table nested in a cell now continues across pages between its own rows and lines, so rows of an inner table taller than a page are no longer clipped away. A page paints only the nested rows and cell blocks it shows, which also keeps a caret or selection in a continued row on the page where that row appears.
