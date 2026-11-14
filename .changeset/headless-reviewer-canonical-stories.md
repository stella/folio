---
"@stll/folio-core": minor
---

A `FolioDocxReviewer` opened with `session: "canonical"` applies operations to headers, footers and notes (`applyDocumentOperationsToStory`) through the canonical editor's public operation compiler, reads those stories from the session's story projections, and resolves their tracked changes in the same session commit as the body's on `acceptAll` and `rejectAll`. `replyTo` and `resolveComment` compile to the session's semantic comment operations.
