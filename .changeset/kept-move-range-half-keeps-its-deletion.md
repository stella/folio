---
"@stll/folio-core": patch
---

Rejecting an insertion that holds another reviewer's pending deletion of a move range end, with the range's start just before the insertion, keeps the end in that deletion in the editor, as the engine does, instead of joining both halves into one point outside it.
