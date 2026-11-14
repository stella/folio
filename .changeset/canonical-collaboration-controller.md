---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A canonical session reports every published commit through `onCommitted`: its operations, their inverse, the version it reached, its origin (`input`, `undo`, `redo` or `remote`) and an operation id, minted by the session's `newOpId` option for a local commit (a random UUID by default) and the sequenced batch's own for a remote one.

`joinCanonicalRoom` joins a canonical session to a sequenced room over a `CanonicalRoomTransport` (base snapshot identity, head revision, submit, journal subscription, id lease grant and renewal), with `createCanonicalMemoryRoom` as the in-memory reference transport around `createSequencer` and `canonicalSnapshotId` as the snapshot identity of a DOCX. Local commits are enqueued on the optimistic client and submitted in order; sequenced batches are rebased by the client and applied through `prepareRemote`; a refused local edit is taken back and reported to the host; a reconnect resubmits the unanswered batch under the same operation id; a joining session replays the journal from revision 0 and is refused on a snapshot mismatch. `@stll/docx-core/ops` exports the `BatchAck`, `BatchReject`, `BatchSubmission` and `ClientNotice` types.
