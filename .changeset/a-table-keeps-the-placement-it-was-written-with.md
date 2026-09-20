---
"@stll/docx-core": minor
"@stll/folio-core": minor
---

Read `w:tblPr/w:jc` and `w:trPr/w:jc` against `ST_JcTable`. The two had a
reader of their own that accepted `left`, `center` and `right`, folded `start`
onto `left` at the table and refused it at the row, so a table written `start`
saved as `left` and one written `end` saved with no `w:jc` at all. `start` and
`end` are members now, resolved against the table's `w:bidiVisual` at layout
and written back as authored.
