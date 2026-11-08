---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The experimental canonical session now opens a document with inline content controls. Typing inside a control keeps the text in it, and a caret just inside a control's edge types into the control while one just outside types beside it. The same holds for typing inside a hyperlink, which no longer cuts the link in two. Locks hold for inline controls as for block ones. A suggestion over the whole text of a control refuses, since it would be recorded as the control's deletion.
