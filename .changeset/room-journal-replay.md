---
"@stll/folio-core": minor
---

A server can keep a canonical room: `FolioDocxReviewer.openRoomBase` reads a room's revision 0 as an editor's canonical session reads the DOCX, `readCanonicalYjsJournal` (`@stll/folio-core/controller/canonicalYjsRoom`) reads the authority's journal back from a stored Yjs room, and `FolioDocxReviewer.applyRoomJournal` replays it on a canonical reviewer, whose `toBuffer` then saves the room. `CANONICAL_YJS_AWARENESS_FIELD` names the awareness field a peer's presence travels in, and a room's `pending` event tells the host how many local batches it has not sequenced yet. `canonicalSnapshotId` and `CanonicalRoomBase` move to `@stll/folio-core/canonical/canonicalRoomBase`. A canonical headless reviewer now opens a password-protected package with its `password` option.
