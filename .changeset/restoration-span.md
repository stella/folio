---
"@stll/docx-core": patch
---

A paragraph restoration (an undo that gives a paragraph back its content, and each paragraph a comment restoration rewrites) no longer conflicts with every concurrent edit of that paragraph. It conflicts only with an edit reaching into the stretch it changes; any other typing, deletion, format or inline record split or join in that paragraph passes it in either order, in the body, a table cell, a header or footer, or a note.
