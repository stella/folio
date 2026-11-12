---
"@stll/folio-react": patch
"@stll/folio-core": patch
---

In a canonical session the React editor holds the session's own document as its document state and hands the host one copy per settled change, as the Vue editor does. Layout reads a model that keeps its identity between edits, and a settled change copies the document once instead of twice. `PagedEditor`'s `onDocumentChange` now receives the model and the host's document together; the model is typed read-only (`ReadonlyDocument`).
