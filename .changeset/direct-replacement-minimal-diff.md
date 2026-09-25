---
"@stll/folio-core": patch
---

Apply direct-mode `replaceInBlock`, `replaceRange` and `replaceBlock` edits as the difference between the matched text and its replacement. Characters the replacement keeps keep their runs, formatting, content controls, note references, fields, bookmarks and drawings; appending one character to a paragraph no longer rewrites every run with the first run's formatting or drops the paragraph's inline content controls and note references.
