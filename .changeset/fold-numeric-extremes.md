---
"@stll/docx-core": patch
"@stll/folio-core": patch
"@stll/folio-cli": patch
---

Fold over values for maxima and minima instead of spreading arrays into `Math.max` and `Math.min`, so large tables, lists and selections stay within the engine's argument limit.
