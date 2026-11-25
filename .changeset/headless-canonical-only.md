---
"@stll/folio-core": minor
---

`FolioDocxReviewer` runs only on the canonical session: the legacy ProseMirror engine and its `session: "legacy"` option are removed, `HeadlessReviewerSession` is `"canonical"`, and `DEFAULT_HEADLESS_REVIEWER_SESSION` and `headlessReviewerUsesCanonicalSession` are gone. `FolioRoomJournalError` no longer has the `legacyEngine` reason.
