---
"@stll/folio-core": patch
---

On the legacy engine, `FolioDocxReviewer` resolved views (`readReviewedStory` and `storyTables` with `final` or `original`, `toResolvedDocument`) read the document resolved by `resolvedDocumentView`, as the canonical engine and clean Markdown do.
