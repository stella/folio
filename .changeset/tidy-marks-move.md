---
"@stll/folio-core": minor
---

Two operations for the edit that moves a paragraph mark and no words:
`splitBlock` writes an inserted mark on the paragraph the break now ends, and
`mergeBlockWithNext` a deleted one. Both carry the `separator` the break stands
in for, deletion-marked on a split and insertion-marked on a merge, so either
direction of accept/reject reproduces the right spacing. A deleted mark is
refused where there is no sibling to join with — the last paragraph of a table
cell, or of a story.

`compareDocx` emits them, so a split is reported as `split` and a merge as
`merge` rather than as a rewrite of the half that stayed put plus an insertion
or deletion of the other. On a 320-paragraph document of splits and merges the
change list drops from 145 entries to 87 and the redline's text-carrying runs
from 156 to 87.
