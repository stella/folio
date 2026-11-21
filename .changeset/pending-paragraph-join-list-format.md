---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Refuse a list format change of a paragraph whose mark is pending deletion with the new docx-core refusal reason `pendingParagraphJoin` (was `revisionConflict`): `changeListFormat` decides it before compiling, in every mode, for the story's paragraphs and those of its text boxes. The canonical session reports the reason as `pendingParagraphMarkDeletion`, the skip reason both engines give a property edit of such a paragraph. The compiler groups a text box's paragraphs joined by a pending mark deletion when a position names them through the story drawing the box; docx-core exports that grouping as `editorParagraphGroupsThroughTextBoxes`.
