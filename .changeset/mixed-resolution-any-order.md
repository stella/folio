---
"@stll/docx-core": patch
---

Accepting some of a run's tracked changes and rejecting others gives the same document in more orders. A pending seam no longer names an insertion that resolving the changes inside it removed, which made the next resolution refuse; and a range boundary an earlier rejection kept at a seam goes inside the piece the seam later joins.
