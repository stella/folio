---
"@stll/folio-core": patch
---

A document with a text box written with a fallback branch now saves after an edit that rewrites that text box. The save failed because the rewritten paragraphs used namespace prefixes the captured branches did not declare.
