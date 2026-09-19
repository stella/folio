---
"@stll/folio-core": patch
---

Make saving a threaded comment set a fixed point. `word/commentsExtended.xml` and the paraId minting it keys on were built by walking the model separately from `word/comments.xml`, so the first save wrote an entry order the second save could not reproduce. Both parts, and the minting, now come from one plan.
