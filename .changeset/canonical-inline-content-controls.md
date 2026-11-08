---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The experimental canonical session now opens a document with inline content controls. Typing inside a control keeps the text in it, and a caret just inside a control's edge types into the control while one just outside types beside it, in editing and suggesting mode. Deleting inside a control keeps the control, and locks hold for inline controls as for block ones. Replacing a selection of text inside a control refuses for now, since it would cut the control in two.
