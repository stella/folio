---
"@stll/docx-core": patch
---

A range marker kept by a rejection never lands inside a pending tracked change; it stands beside it, and a seam waiting on pending insertions waits across such markers. Rejecting content made inside a piece of a pending insertion cancels that insertion's join facing it, and rejected pieces of one insertion meet around the markers standing between them, so the result no longer depends on resolution order. A marker kept again keeps the revision it was first kept from and the deeper cut.
