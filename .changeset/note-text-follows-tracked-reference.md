---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A canonical tracked operation deleting a note's last live reference also deletes the note's text and paragraph marks, as operations of their own; a comparison states each note's changes as a story of its own. docx-core adds `noteDeletionOps`.
