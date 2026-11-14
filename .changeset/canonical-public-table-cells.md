---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

The canonical session applies the public `insertTableColumn`, `mergeTableCells` and `splitTableCell` operations, directly or as tracked changes. A new column's cells hold the given texts, a line break starting a new paragraph as in a new row, in both engines. Merging cells leaves out cells that hold only an empty paragraph, and splitting a vertically merged cell gives each row its own cells again, each keeping its content. A suggestion the default engine cannot record, such as a split or a column through a merged cell, applies and gives the direct edit once accepted.
