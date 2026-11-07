---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A comment range marker written between blocks, beside a content control for example, is kept as a typed block in the document model and both editors, and is saved back between the same two blocks. The comment used to read as a point comment, and a save after any edit wrote its start twice under the same id.
