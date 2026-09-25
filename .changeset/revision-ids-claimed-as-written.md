---
"@stll/folio-core": patch
---

AI edit batches without a revision stamp claim the shared revision ids they actually write, operation by operation, instead of reserving four per operation up front. A replacement that clears the background of many runs no longer overruns its reservation, a batch started from a comment-id callback cannot reuse an id the outer batch writes afterwards, and consecutive batches follow one another without a gap. Stamped batches are unchanged: their ids stay contiguous from `idSeed`.
