---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` applies operations to headers, footers and notes (`applyDocumentOperationsToStory`) through the canonical editor's public operation compiler, reads those stories from the session's story projections, and resolves their tracked changes in the same session commit as the body's on `acceptAll` and `rejectAll`. `replyTo` and `resolveComment` compile to the session's semantic comment operations. Resolved views (`readReviewedStory` and `storyTables` with `"original"` or `"final"`) are read from a second session resolved to the view, and `resolveReviewedStory` resolves one story's changes in one session commit.
