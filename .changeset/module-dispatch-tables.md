---
"@stll/folio-core": patch
---

Parse DOCX files faster: table, row, cell, section, numbering, font, content-control, comment, hyperlink and paragraph-content child handlers are built once per module instead of once per element.
