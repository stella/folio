---
"@stll/folio-core": patch
---

`FolioDocxReviewer.acceptAll()` / `rejectAll()` resolve `"suggested"` edits too. Accepting kept a suggested paragraph insertion in the reviewer but the next save dropped it (the paragraph was still flagged as a pending suggestion), and rejecting left an empty suggested paragraph behind.
