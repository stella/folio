---
"@stll/folio-core": minor
---

A server can keep a canonical room: `FolioDocxReviewer.openRoomBase` reads a room's revision 0 as an editor's canonical session reads the DOCX, `readCanonicalYjsJournal` (`@stll/folio-core/controller/canonicalYjsRoom`) reads the authority's journal back from a stored Yjs room, and `FolioDocxReviewer.applyRoomJournal` replays it on a canonical reviewer, whose `toBuffer` then saves the room. `canonicalSnapshotId` and `CanonicalRoomBase` move to `@stll/folio-core/canonical/canonicalRoomBase`. A canonical headless reviewer now opens a password-protected package with its `password` option.
