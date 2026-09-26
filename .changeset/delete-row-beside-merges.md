---
"@stll/folio-core": patch
---

The editor's `deleteRow` closes a vertical merge that sits right of a wide merged cell over the removed row, as the table operations already do, instead of leaving the merge one row longer than the table.
