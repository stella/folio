---
"@stll/folio-core": patch
---

Formatting changed over part of a line that holds a line break or a tab resolves to the same runs whether or not the document was saved and reopened in between, and rejecting it gives back the authored run.
