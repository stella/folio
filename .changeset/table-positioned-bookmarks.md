---
"@stll/docx-core": patch
---

Row and column edits, and resolving row and cell changes, work on a table holding bookmark markers beside its rows or a row's cells: a marker keeps its place among the rows or cells that remain (a start before the one that followed it, an end after the one that preceded it), and each edit's inverse restores it exactly. Row edits on such a table were refused, and removed cells left the markers beside them out of range.
