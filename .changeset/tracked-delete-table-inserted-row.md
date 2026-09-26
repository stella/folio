---
"@stll/folio-core": patch
---

A tracked `deleteTable` over a table holding a pending inserted row (for example a table inserted in `"tracked-changes"` mode and not yet accepted) is refused with `unsupportedBlock`, as `deleteTableRow` refuses that row. It used to apply and leave rows marked both inserted and deleted, which no save could write (`Expected at most one structural revision marker`).
