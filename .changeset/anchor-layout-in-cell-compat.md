---
"@stll/folio-core": patch
---

Confine table-cell floating anchors to their cell when the document's compatibility mode is 15 or higher, ignoring an authored `layoutInCell="0"`, and keep the painter and row-break geometry consistent about which anchors are cell-scoped.
