---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Refuse a list format change of a paragraph whose mark is pending deletion with the new docx-core refusal reason `pendingParagraphJoin` (was `revisionConflict`): `changeListFormat` decides it before compiling, in every mode. The canonical session reports the reason as `pendingParagraphMarkDeletion`, the skip reason both engines give a property edit of such a paragraph.
