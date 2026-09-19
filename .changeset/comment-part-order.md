---
"@stll/folio-core": patch
---

Make saving a threaded comment set a fixed point. `word/comments.xml` is written with the top-level comments first and the replies after, so the next parse returns the comments in that order, but `word/commentsExtended.xml` was built by walking the model's own order: the first save wrote an entry order the second save could not reproduce. Both parts, and the paraId minting they key on, now walk one order.
