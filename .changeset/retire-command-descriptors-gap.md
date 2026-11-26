---
"@stll/folio-core": minor
---

Inserting a footnote, endnote or table of contents over selected table cells now applies: a note reference goes after the last selected cell's text and a table of contents at the start of the first selected cell, in editing and suggesting modes alike. Removing a link from part of a field the link holds takes the link off the whole field instead of refusing. Toolbar and structure commands the editor cannot apply now report a `{ condition }` ground (`invalidRequest` or `notReady`) instead of a gap, and the `command-descriptors` gap is retired from `CANONICAL_GAP` and `CANONICAL_CAPABILITIES`.
