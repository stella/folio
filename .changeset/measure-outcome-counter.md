---
"@stll/folio-core": patch
---

The layout instrumentation's `onMeasureBlock` now reports only blocks that were measured. A block that took a measure from the cache reports `onReuseMeasure` instead, so a count of measured blocks is no longer every block a layout pass visited.
