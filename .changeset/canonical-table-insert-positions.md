---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Inserting a table in the experimental canonical session now works from a selection of the whole document and from selected cells, placing the table where the editor does. A whole-document selection puts it at the start of the document, before the first paragraph. Selected cells put it after their table, with an empty paragraph on each side so the two tables never touch. The table block intent gains `side` and `separated`.
