---
"@stll/folio-core": patch
---

In the canonical session, a suggested `deleteTableColumn` through a cell that spans other columns is skipped as `unsupportedBlock`, as the default engine skips it, instead of being reported as a canonical capability gap.
