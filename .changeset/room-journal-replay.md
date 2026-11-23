---
"@stll/folio-core": minor
---

A server can keep a canonical room: `FolioDocxReviewer.openRoomBase` reads a room's revision 0 as an editor's canonical session reads the DOCX, and `FolioDocxReviewer.applyRoomJournal` replays a room's sequenced batches on a canonical reviewer, whose `toBuffer` then saves the room. `canonicalSnapshotId` and `CanonicalRoomBase` move to `@stll/folio-core/canonical/canonicalRoomBase`. A canonical headless reviewer now opens a password-protected package with its `password` option.
