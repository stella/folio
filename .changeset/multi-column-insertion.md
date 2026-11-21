---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

Add `columnCount` to `insertTableColumn`: one operation inserts that many adjacent columns, filled from `cellTexts` in reading order, and tracked or suggested they resolve together. The canonical session refuses a second tracked column insertion in a table whose columns already have a pending change.
