---
"@stll/folio-core": minor
"@stll/folio-react": minor
---

Reset the hidden editor on every document load, not only when the host's `documentKey` changes. `DocumentLoaderManager` now publishes a per-load identity (`setLoadedDocumentIdentity`) in the same commit as the history reset and orders parsed-document loads against in-flight buffer parses; the React `DocxEditor` feeds that identity to the paged editor, so loading a new `document`/`documentBuffer` (or `loadDocument`) into an editor with a live view replaces the painted content instead of leaving it on the previous document. The `documentKey` prop is deprecated and ignored.
