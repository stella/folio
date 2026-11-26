---
"@stll/folio-core": minor
---

A template plan can keep the tracked changes of its source (`trackedChanges: "keep"`): the plan is made for the accepted view, and the output shows the same insertions and deletions inside its paragraphs, by the same authors and dates, with tracking switched off. Verification holds the output to each change. So far only changes inside paragraphs are kept; a source with an inserted or deleted paragraph mark, table row or cell, a move, or a change inside another change, a drawing or an equation is refused as `tracked-changes` with what stands in the way. `analyzeDocxBlocks` takes `keep` too and gives the same answer before a plan is made.
