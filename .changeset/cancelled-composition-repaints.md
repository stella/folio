---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A cancelled composition no longer stays painted on the page. A layout pass that ran while composing could draw the provisional text, and nothing repainted the page once the composition was cancelled or refused.
