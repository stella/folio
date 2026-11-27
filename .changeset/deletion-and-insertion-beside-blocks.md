---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A paragraph deleted directly before a table, a preserved block or a content control now goes whole after a block or past block-level bookmarks too, as accepting its tracked deletion removes it, so the list numbering after it matches in every mode; it stays only between two tables, bookmarks between them included. Rejecting a paragraph put after one whose break was added before a table now runs that paragraph into the table, as without the new one, also in a body ending in a content control.
