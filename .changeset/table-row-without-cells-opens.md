---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A document with a table row holding no cells, which the file format allows, now opens instead of failing with an invalid document model. Row and column edits elsewhere in the table leave such a row where it is.
