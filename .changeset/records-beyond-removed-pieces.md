---
"@stll/docx-core": patch
---

A removal's seam waits on the insertions standing between pieces of a pending insertion only, not of a pending deletion, and a removal resolved with the pieces it cut reads that cut no deeper than removing the pieces first leaves: a range marker it keeps lands in the same place whichever order the neighbouring changes are resolved in.
