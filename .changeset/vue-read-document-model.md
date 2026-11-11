---
"@stll/folio-vue": patch
---

The paged editor handle gains `readDocumentModel()`, as in React: in a canonical session it returns the session's own document without a copy, otherwise null.
