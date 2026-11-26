---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Preserve the public comment and review reads and their existing command exports. Move reference-only ProseMirror revision transactions, table merge resolution and cleanup, and supporting paragraph and comment helpers into test support. The former deep imports for those helpers and the legacy `folioRevisionResolution` transaction-step registration are removed.
