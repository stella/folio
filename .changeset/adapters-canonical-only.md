---
"@stll/folio-react": patch
"@stll/folio-vue": patch
"@stll/folio-core": patch
---

The React and Vue editors run review, document operations, suggestions, comments and section properties only through the canonical session. Without one (a document the session refused to open, which is read-only), these calls change nothing: document operations skip every operation, accept and reject return `false`, and comment edits report a read-only attempt, instead of editing ProseMirror state directly.
