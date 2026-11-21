---
"@stll/folio-core": patch
---

On the canonical engine, table operations of one batch now agree across modes as on the legacy engine. An operation a tracked batch refuses only because no tracked change records it (a merge across a row, a split of a cell spanning columns, a column that would widen or narrow a spanning cell) still claims the cells it would edit directly, so a later operation of the batch on those cells is refused as overlapping in either mode. A merge or split in a table the same batch inserts or deletes a column of is refused, whatever their order. A table, row, column or cell edit named through a paragraph the same batch deletes finds its target through a paragraph the batch keeps.
