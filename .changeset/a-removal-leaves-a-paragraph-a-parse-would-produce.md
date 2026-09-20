---
"@stll/folio-core": patch
---

Re-consolidate a paragraph after a parse-time normaliser removes an inline item.

`parseParagraph` consolidates a paragraph's runs, and every inline item that is not a run is a merge boundary. Three normalisers then remove items from that already-consolidated array: a comment marker naming a comment the package does not define, a move-range marker with no other half, and the per-paragraph range markers a multi-paragraph comment is cut into. Each left two mergeable runs adjacent, which the parse that consolidated them would never have produced.

That is an oscillation rather than a loss. Save 1 wrote the pair, the next parse merged it, save 2 wrote one run, so the second save differed from the first. The removal now restores the invariant where it happens, in `InlineContentRemovals.apply`, so every normaliser that removes an inline item gets it and none has to remember.
