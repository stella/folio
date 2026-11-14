---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` applies operations to headers, footers and notes (`applyDocumentOperationsToStory`) through the canonical editor's public operation compiler, reads those stories from the session's story projections, and resolves their tracked changes in the same session commit as the body's on `acceptAll` and `rejectAll`. `replyTo` and `resolveComment` compile to the session's semantic comment operations. Resolved views (`readReviewedStory` and `storyTables` with `"original"` or `"final"`) are read from a second session resolved to the view, and `resolveReviewedStory` resolves one story's changes in one session commit. Suggested batches stage on the session and are recorded for `exportPendingSuggestions`; `loadPendingSuggestions` restages records through the session, and `acceptSuggestion`, `rejectSuggestion` and the suggested revisions of `acceptChange`, `rejectChange`, `acceptAll` and `rejectAll` resolve through its suggestion intent. A suggestion staged in a header, footer or note refuses to resolve.

`FolioDocxReviewer.fork()` gives a live, independent copy of a reviewer on either engine, pending suggestions included, with an empty undo history.
