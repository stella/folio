---
"@stll/docx-core": patch
---

A direct deletion whose range holds a note's own reference mark keeps the mark, and now records that in its sequenced effect (`keptMark` in a new `deleteRange` effect). A concurrent edit rebased over it lands after the kept mark, as it does when the edit applies first, so a format over the same range covers the mark in either order.
