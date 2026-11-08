---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
"@stll/folio-cli": patch
---

Accepting or rejecting a tracked change now honours a content control's `w:lock`. A change stays pending as it is when resolving it would change the content of a control locked with `contentLocked` or `sdtContentLocked`, or remove one locked with `sdtLocked` or `sdtContentLocked`.

- `FolioDocxReviewer.acceptChange` and `rejectChange` return `false` for such a change.
- `acceptAll()` and `rejectAll()` resolve every other change and now return `{ resolved, skipped }`, listing each held change with the reason `contentControlLock`. They previously returned the number of changes resolved, which is now `resolved`.
- The editor commands `acceptChange`, `rejectChange`, `acceptAIEditRevision`, `rejectAIEditRevision`, `acceptAllChanges` and `rejectAllChanges` refuse the same changes. `changeResolutionLockViolation` and `lockHeldRevisions` tell a host which changes a lock holds.
- In the React editor, the toolbar's accept and reject buttons and the context menu's Accept Change and Reject Change are disabled for such a change, with the reason. In the Vue editor, the change card's buttons are.
