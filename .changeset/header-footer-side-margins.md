---
"@stll/folio-core": patch
---

Headers and footers no longer cut off text indented into the side margins. Their height cap clipped both axes, so a paragraph with a negative `w:ind` (a right-aligned "Version" line, a licence footer spanning the margins) lost the part that belongs in the margin.
