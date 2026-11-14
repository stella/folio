---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

The canonical session collaborates. The `collaboration` prop of the React and Vue editors (and the `collaboration` option of `useDocxEditor`) takes a canonical room (`CanonicalRoomCollaboration`: the DOCX's `snapshot`, this peer's `actor`, `connect`, and the optional `user` and `onEvent`) when the session is canonical, and the Yjs fragment (`DocxEditorYjsCollaboration`) when it is legacy; each engine refuses the other's. The editor opens the room with `connect` once it has read its document, joins the session to it, takes local edits once the room granted the session its id lease, and closes the connection when it opens another document or unmounts. The hidden editor of a canonical session no longer uses y-prosemirror: peers' edits are sequenced operations published as canonical commits, and their selections are drawn from canonical positions (`RemoteSelection.peer` names the peer). In a canonical room the controlled `comments` prop only reports the room's comments, which travel in the journal. `CanonicalSession.adoptIdLease` gives an unedited session the lease a room granted it.
