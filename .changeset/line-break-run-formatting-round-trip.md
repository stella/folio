---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Formatting changed over part of a line that holds a line break or a tab resolves to the same runs whether or not the document was saved and reopened in between: alike runs split beside a regular tab or a line break now read as one run, as alike text runs already did.
