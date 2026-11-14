---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` applies operations to the body (`applyOperations`, `applyDocumentOperations`, and `applyDocumentOperationsToStory` on the main story) through the canonical editor's public operation compiler, undoes them with `undoDocumentOperations` from the session's journal, and resolves the body's tracked changes (`acceptChange`, `rejectChange`, `acceptAll`, `rejectAll`) through the session's review intents. Its reads follow the session after every edit. Operations in headers, footers and notes, suggested batches, and resolutions beside pending suggestions, in other stories or in the final section's properties throw a `CanonicalSessionError` grounded in the `publicOps.headlessSession` gap.
