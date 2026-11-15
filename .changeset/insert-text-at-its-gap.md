---
"@stll/docx-core": patch
---

`insertText` puts its text exactly at the gap its position names, after the zero-width leaves `zeroWidthBefore` counts and before the others, and joins the run holding the leaf before that gap, else the one holding the leaf after it, whatever kind of leaf either is. Typing after a run that holds only a note's own reference mark joins that run, and text typed after a bookmark's end or a cached page break stays after it. A sequenced position at the typing's offset moves past the text exactly when it lies after that gap.
