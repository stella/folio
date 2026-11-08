---
"@stll/folio-core": patch
---

A full save writes an existing note's reference mark only when the note still has one: a note whose mark was deleted, or that never had one, no longer gets a new mark when the package is rebuilt. A note new to the package is still given its mark.
