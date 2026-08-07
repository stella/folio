---
"@stll/folio-core": patch
---

Seed headless `paraId` attributes by rebuilding the document instead of applying one ProseMirror step per paragraph, removing two quadratics from `FolioDocxReviewer.fromBuffer` (-29% on a real DOCX corpus, -45% on the largest file). Also seeds RTL base direction explicitly, which previously rode on the paraId transaction and was skipped entirely for documents that already carried Word-authored paraIds.
