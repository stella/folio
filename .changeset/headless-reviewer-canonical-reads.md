---
"@stll/folio-core": minor
---

`FolioDocxReviewer.fromBuffer` takes `session: "legacy" | "canonical"` (default `"legacy"`). A canonical reviewer loads the document into the canonical session the editors' canonical mode runs, and its reads (snapshots, content, stories, changes, comments, numbering and `toDocument`) come from that session. Its edits, review resolution, undo, comparison and save throw a `CanonicalSessionError` grounded in the `publicOps.headlessSession` gap until they move to the session.
