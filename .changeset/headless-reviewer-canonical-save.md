---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` saves through the canonical editor's save: `toBuffer` and `save` write the session's committed snapshot as a selective patch of the loaded package when it applies and as a repack otherwise (`save({ repack: "refuse" })` then reports the reason and returns no buffer), with pending suggestions written as rejected. `toResolvedDocument` resolves every change on a second session over the current document, leaving the reviewer's own session and undo history as they are.
