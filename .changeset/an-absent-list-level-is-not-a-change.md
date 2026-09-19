---
"@stll/folio-core": patch
---

Treat a list level neither document defines as the same absence rather than a difference. A `w:numPr` naming an abstract numbering with no such `w:ilvl` resolved to no level on either side, and the staging read that as a changed definition: the paragraph was remapped onto a freshly minted `numId`, so a document compared with itself reported a numbering change and any direct `w:ind` the new numbering displaced could no longer be moved back.
