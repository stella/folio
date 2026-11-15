---
"@stll/folio-core": patch
---

In the canonical session, suggestions that would narrow, widen or split a cell spanning several columns are skipped as `unsupportedBlock`, as the default engine skips them: a suggested `deleteTableColumn` through such a cell, a suggested `insertTableColumn` that runs through it, and a suggested `splitTableCell` of it. The same operations applied directly still work.
