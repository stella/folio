---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A canonical commit reports what it does to the session's undo history (`history`: `opened`, `joined`, `undo` or `redo`), and `canonicalHistoryEffect` reads it from the commit's transaction, so a host that keeps its own undo stack can mirror the session's steps. With `hostShortcuts` including `"history"`, the undo and redo keys no longer undo in a canonical session's views; the host undoes through the editor's API.
