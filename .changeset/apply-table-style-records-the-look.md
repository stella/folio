---
"@stll/folio-core": patch
---

Record the `w:tblLook` that applying a table style asks for. `applyTableStyle` painted the first row, the last row and the bands the style names but left the table's own look untouched, so a save kept whatever the document arrived with and Word resolved the conditional formatting differently from what was on screen. The command now states each of the six regions explicitly and leaves `w:val` as the author wrote it: the attribute form is what a reader resolves first, and re-encoding the bitmask would rewrite bits folio does not model.
