---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Accepting some of a run's tracked changes and rejecting others now gives the same document in any order. A range boundary a rejection keeps stays whatever other occurrences of it the story holds, so `planWellFormedRanges` no longer drops occurrences: `WellFormedRangePlan` loses `dropped` and `BoundaryOccurrence` loses `boundary`. A boundary an earlier rejection kept at a seam goes inside the piece the seam later joins, and a pending seam no longer names an insertion that resolving the changes inside it removed.
