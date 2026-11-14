---
"@stll/folio-core": patch
---

The canonical session applies the public `deleteTable`, `insertTableRow`, `deleteTableRow` and `deleteTableColumn` operations, directly or as tracked changes, giving the package the default engine gives. A new row is built as the default engine builds it, from the row beside it and the cell texts. A row or column operation on a block in no table is skipped as the default engine skips it, without naming a canonical gap.
