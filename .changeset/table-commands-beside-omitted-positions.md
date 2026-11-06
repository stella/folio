---
"@stll/folio-core": minor
---

The editor's row and column commands, and merge and split, now work on tables with short rows, `w:gridBefore` or `w:gridAfter`, and a tracked column deletion works through such a gap. A new column at a position a row omits grows the omission instead of adding a cell. A merge through an omitted position is refused.
