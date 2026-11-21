---
"@stll/docx-core": patch
---

Resolve a decision that takes every cell of a table away (deleted cells accepted, inserted cells rejected) by removing the table, as removing its last row does, instead of refusing it.
