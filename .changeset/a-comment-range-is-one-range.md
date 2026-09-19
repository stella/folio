---
"@stll/folio-core": patch
---

Keep a comment range whole across the paragraphs it covers. The conversion
tracked the open ranges per paragraph, so a comment on paragraphs 1 to 3 marked
the first and the last (where its two boundaries sit) and left the middle
unhighlighted, and the save path, reading each paragraph on its own, wrote one
range per marked paragraph where the author wrote one. The open ranges now flow
with the block walk, into table cells, text boxes and content controls, and a
save emits one `w:commentRangeStart` at a comment's first marked position and
one `w:commentRangeEnd` at its last.
