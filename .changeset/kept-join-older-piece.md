---
"@stll/docx-core": patch
---

Stop deferring an accepted or kept revision's seam on pending insertions made before it: they are pieces of the record it cut, so rejecting one later keeps a range boundary it held where it was instead of moving it into the record across the seam.
