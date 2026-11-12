---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

The optimistic client's `receiveBroadcast` and `receiveReject` report how its document moved (`ClientDocumentChange`): a `rebased` change, a foreign revision rebased over the pending local batches or a refused batch's inverse rebased over the ones after it, or a `replaced` change when a pending batch was dropped. `transformPosition` moves a position across sequenced batches by paragraph id and offset, and `applyBatch` is exported.

A canonical session applies another actor's sequenced edit with `prepareRemote(state, change)`, through the same staging and publishing path as a local edit. The edit is not an undo step: the local undo and redo steps are rebased over it, a step that no longer rebases is dropped with the steps beneath it and reported to `onHistoryDrop`, and the selection moves by paragraph id and offset. A commit carries the operations it applies (`ops`).
