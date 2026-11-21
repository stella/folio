---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep a run's pending property change id when a tracked patch folds the whole run into it, and report only revision ids a canonical batch leaves standing. Refuse a suggestion's fold into another pending property change with `pendingRunPropertyChange` or `pendingParagraphPropertyChange`, and a tracked join into a table under review with `joinIntoPendingTable` (reported as `pendingParagraphMarkDeletion`).
