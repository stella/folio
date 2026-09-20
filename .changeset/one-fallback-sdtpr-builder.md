---
"@stll/folio-core": patch
---

Serialize a programmatic content control's `<w:sdtPr>` the same way wherever it sits: one builder for block and inline, writing `w:id` only for a decimal number and `w:lock` only when something is locked.
