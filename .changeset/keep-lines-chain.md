---
"@stll/folio-core": patch
---

Honour `w:keepLines` by moving a paragraph that does not fit whole to the next page, and keep a `w:keepNext` paragraph with the full opening its following paragraph cannot split (the whole paragraph under `w:keepLines` or when widow control leaves it fewer than four lines, otherwise two lines under widow control).
