---
"@stll/folio-core": patch
---

Read every model content union by exhaustion. The conversion, serializer and markdown modules narrowed `ParagraphContent`, `BlockContent`, a hyperlink's children and the table structural-change union by chains of `else if`, so a member added to any of them reached whichever branch happened to be last and the build still passed. Each is now a `switch` with a `never` default, and a lint rule fails the next chain written in those directories.

Three bugs the conversion surfaced: a bookmark pair inside `w:bdo`/`w:dir` was not recognised as a pair, so the start fell back to the legacy paragraph attribute and the end was dropped; a text-box anchor inside one was neither resolved nor removed and reached the saved package; and a footnote or comment holding an equation rendered as markdown without it. Two unused plain-text flatteners that disagreed with the one owner are gone.
