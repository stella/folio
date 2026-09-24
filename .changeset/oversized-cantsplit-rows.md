---
"@stll/folio-core": patch
---

Split a `w:cantSplit` table row that is taller than a page: it moves to a fresh page and continues across pages under repeated header rows instead of overflowing the page. Exact-height rows still stay whole.
