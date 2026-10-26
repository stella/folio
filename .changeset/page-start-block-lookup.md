---
"@stll/folio-core": patch
---

`findPageShellForPmPos` treats a page's position range as half-open, so a block that starts a page resolves to that page rather than the page before it; this corrects `getPageNumberForPmPos`, navigation to such a block and its painted block geometry.
