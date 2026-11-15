---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Undo of a committed document operation on the canonical session restores the exact prior document when the operation removed notes with their references and list rendering then renumbered the paragraphs after it: `inverseMeetingDocument` rebases each whole-part restore of an inverse onto the document it meets when the inverse runs. A rejected undo of a document-operation batch names its cause: `documentChanged` when the document changed after the batch, `storyUnavailable` when the story the batch edited is not open, `inverseRefused` (with the session's message) when the batch's inverse cannot be staged, and `documentNotEditable` when the host does not take the undone document.
