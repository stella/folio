---
"@stll/folio-core": patch
---

A rejected undo of a document-operation batch names its cause: `documentChanged` when the document changed after the batch, `storyUnavailable` when the story the batch edited is not open, `inverseRefused` (with the session's message) when the batch's inverse cannot be staged, and `documentNotEditable` when the host does not take the undone document.
