---
"@stll/folio-core": patch
---

`compareDocx` and `generateRedlineDocx` take `session` and run on the canonical engine: run formatting and link changes, kept undefined references, and copied tables and rows.
