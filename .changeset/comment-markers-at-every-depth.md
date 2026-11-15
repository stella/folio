---
"@stll/folio-core": patch
"@stll/docx-core": patch
---

A comment whose range holds only empty runs keeps its anchors and its definition when the document is prepared for editing and saved. When a save removes a comment, its range markers and reference go from inside content controls, hyperlinks, fields, tracked and moved changes, inline wrappers and text boxes as well, so the save is not refused.
