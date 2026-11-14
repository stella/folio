---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

In a canonical room, an editor open on a header, a footer or a note takes a peer's edit to that story at once, with its selection moved over the edit. The React and Vue editors (and `useDocxEditor`) take `onCanonicalHistoryDrop`, called with how many local undo and redo steps a peer's edit made unusable, so a host that mirrors the steps in its own undo stack removes as many.
