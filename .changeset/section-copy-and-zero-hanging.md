---
"@stll/folio-core": patch
---

Accepting a pending join of a section-ending paragraph no longer gives the section to a paragraph split from it or inserted before it, so the live document and its saved package resolve alike. A `w:hanging` of zero now reads as a first-line indent of `0`, never `-0`.
