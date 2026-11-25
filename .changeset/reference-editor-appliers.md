---
"@stll/folio-core": minor
"@stll/folio-react": patch
---

Move the ProseMirror reference appliers into test support. Deep `./*` imports of the former applier, revision-stretch, and note-inventory modules break, and `restoreNotes` and `updateDocumentContent` are no longer exported. Use canonical document operations for editing. Canonical save diagnostics now use the `save-serialization` gap.
