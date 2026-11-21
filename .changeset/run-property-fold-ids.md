---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Keep a run's pending property change id when a tracked patch folds the whole run into it, report only revision ids a canonical batch leaves standing, and refuse a suggestion's fold into another pending property change with `pendingRunPropertyChange` or `pendingParagraphPropertyChange`.
