---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

A canonical commit reports what it does to the session's undo history (`history`: `opened`, `joined`, `undo` or `redo`), and the React `DocxEditor` passes it to `onCanonicalHistory` after each published edit, undo or redo, so a host that keeps its own undo stack can mirror the session's steps. With `hostShortcuts` including `"history"`, the undo and redo keys no longer undo in a canonical session's views; the host undoes through the editor's API.

No layout pass paints the React editor's pages before the initial wait for the document's fonts ends; a pass asked for during the wait is covered by the initial one.
