---
"@stll/folio-core": patch
---

Accepting or rejecting a tracked change that joins a paragraph into one the same resolution empties keeps the first paragraph's identity, so a deleted last paragraph or a rejected paragraph after it no longer takes over the id of the paragraph that stays.
