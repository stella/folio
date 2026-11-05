---
"@stll/folio-core": patch
---

A table nested in a cell now continues across pages between its own rows and lines, so rows of an inner table taller than a page are no longer clipped away. A page paints only the nested rows it shows, which also keeps a caret in a nested row on the page where that row appears.
