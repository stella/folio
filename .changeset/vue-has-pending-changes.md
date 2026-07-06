---
"@stll/folio-vue": minor
---

Implement `DocxEditorRef.hasPendingChanges()` in the Vue adapter (previously
stubbed to `false`). It now returns whether the live editor has edits not yet
serialized by `save()`, mirroring the React adapter's set/clear points: a
doc-dirty flag set on every doc-changing transaction and cleared on a successful
`save()` and on document load/swap, OR-ed with a comment-list dirty flag set on
comment mutations (add/reply/resolve/unresolve/delete) and cleared on load and
save. Moves `hasPendingChanges` from deferred to paired in the cross-adapter
parity contract.
