---
"@stll/docx-core": patch
---

Defer an accepted or kept revision's seam only on pending insertions made after it: an older insertion beside it is a piece of the record it cut, so rejecting that piece later keeps a range boundary it held where it was instead of moving it into the record across the seam.
