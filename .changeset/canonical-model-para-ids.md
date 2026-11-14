---
"@stll/folio-core": patch
"@stll/folio-vue": patch
---

A document model whose paragraphs have no ids now opens in a canonical session: its paragraphs get exactly the ids its package gets when opened as bytes, and keep them through edits, saves and reopening.
