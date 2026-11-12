---
"@stll/docx-core": patch
---

Rejecting a paragraph mark and the insertions beside it at once now gives what rejecting them one at a time gives: a range boundary kept at the seam the paragraphs join at goes inside the piece the join merges, instead of keeping the pieces apart.
