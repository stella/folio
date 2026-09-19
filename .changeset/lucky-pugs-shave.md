---
"@stll/folio-core": minor
---

Version the node attrs Folio persists in a collaboration document. A snapshot written by a newer attr schema is now refused with a typed error instead of being rebuilt attr by attr, and `migrateFolioYjsSnapshot` carries a stored snapshot forward offline.
