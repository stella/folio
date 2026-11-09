---
"@stll/folio-core": patch
---

Layout no longer serializes every paragraph after an edit to look up its measurement: a paragraph the edit only moved takes the key of the block it was moved from.
