---
"@stll/folio-core": patch
---

Indentation set in the editor (increase/decrease indent, left/right/first-line indent) is saved on a paragraph that already carries its own `w:pPr`; it used to save the source `w:ind` whatever the editor showed. The block snapshot's `directIndentation` now reports indentation a command set, as the saved package does, with a hanging indent read as a negative first line.
