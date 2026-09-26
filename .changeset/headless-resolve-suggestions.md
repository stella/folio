---
"@stll/folio-core": patch
---

`FolioDocxReviewer` resolves `"suggested"` edits in `acceptAll()` / `rejectAll()` and in `acceptChange()` / `rejectChange()`. Accepting kept a suggested paragraph insertion in the reviewer but the next save dropped it (the paragraph was still flagged as a pending suggestion), and rejecting left an empty suggested paragraph behind. Rejecting one change of a suggestion now rejects the whole suggestion.
