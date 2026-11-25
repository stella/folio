---
"@stll/docx-core": patch
---

Sequencing rebases an edit over a concurrent deletion whose ends name zero-width leaves when the edit has no position at an offset the deletion spans, instead of refusing every such pair. `trackedOrderPair` leaves out the pairs the transform still refuses, through the same `zeroWidthSpanOverlap` check.
