---
"@stll/docx-core": patch
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

In a canonical room, edits to headers, footers and notes travel between peers: the wire decoder refused every position outside the main story, so a peer editing a note stopped following its room. The React and Vue editors (and `useDocxEditor`) take `onCanonicalHistoryDrop`, called with how many local undo and redo steps a peer's edit made unusable, so a host that mirrors the steps in its own undo stack removes as many.
