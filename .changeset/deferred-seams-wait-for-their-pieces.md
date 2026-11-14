---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Range boundaries a rejection keeps at a seam now go only into the two pieces of the record the removed content cut. A pending insertion beside the seam, or a piece still held in another reviewer's pending deletion or move, does not take them: they stand at the seam until its pieces meet, and the seam waiting on that insertion records the date of the revision they were kept from, saved with the document. So the boundaries land in the same place whichever order the changes are accepted or rejected in.
