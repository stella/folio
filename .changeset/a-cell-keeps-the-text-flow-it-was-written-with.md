---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Generate `TableCellTextDirection` from `ST_TextDirection`. The union omitted
`lrTb`, `lrTbV` and `tbLrV`, so a cell written with one lost the attribute at
parse time and saved without it. The editor's writing-mode map is now total
over the enumeration, and the display list reports an unpainted vertical flow
for every direction the painter turns rather than the two that were listed.
