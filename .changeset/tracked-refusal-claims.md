---
"@stll/folio-core": patch
---

On the canonical engine, a table operation a tracked batch refuses only because no tracked change records it (a merge across a row, a split of a cell spanning columns, a column that would widen or narrow a spanning cell) still claims the cells it would edit directly, so a later operation of the batch on those cells is refused as overlapping in either mode, as on the legacy engine.
