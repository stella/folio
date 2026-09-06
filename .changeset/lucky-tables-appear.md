---
"@stll/folio-core": minor
"@stll/folio-agents": minor
---

`insertTable` and `deleteTable` add and remove a whole table: every row marked
`w:trIns` or `w:trDel` in tracked mode, which is how Word says it. `compareDocx`
emits them as `table-insert` / `table-delete`, so a pair whose table count
differs is compared instead of refused.

Its segment pairing no longer goes by index. A table pairs with the table
opposite it unless the next table on one side matches it better, so removing
the first table no longer shifts every later one and rewrites each table's
contents into the next. A segment is the outermost table plus everything nested
in it; `FolioAIBlock.table` gains `outerTableIndex` to carry that distinction.
