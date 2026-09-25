---
"@stll/folio-core": patch
---

Justified lines (`w:jc="both"`, `compatibilityMode` 15) keep an overflowing word by shrinking spaces only when the stretch the shorter line would need is large compared with the shrink, and list paragraphs use the same quarter-of-a-space shrink limit as other prose.
