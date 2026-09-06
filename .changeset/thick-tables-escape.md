---
"@stll/folio-core": patch
---

A block inserted next to a block inside nested tables now lands at document
level. `insertBeforeBlock`, `insertAfterBlock`, and `insertSignatureTable`
escape the table their anchor sits in; they escaped only the innermost one, so
an anchor two tables deep left the new block inside the outer cell. This is
what made `compareDocx` refuse a pair whose base ends with a nested table and
whose target appends a paragraph after it.
