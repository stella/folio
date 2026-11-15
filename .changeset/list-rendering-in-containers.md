---
"@stll/folio-core": patch
---

A paragraph an edit numbers inside a table cell, a content control or a custom XML wrapper gets its list rendering in the live document, as it does in the body and as the saved file reopens. The parser and the edit normalization derive the rendering's definition fields with the same function, and a list's numbering continues through table cells in document order, row by row and cell by cell, instead of each cell's paragraphs reading the level's template.
