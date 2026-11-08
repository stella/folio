---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

Internal cleanup for the stricter shared lint rules: merged duplicate imports, block-scoped `case` bodies and `flatMap` in place of `map().flat()`. No behavior change.
