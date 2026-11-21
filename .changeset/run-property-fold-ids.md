---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Keep a run's pending property change id when a tracked patch folds the whole run into it, and report only revision ids a canonical batch leaves standing. Refuse a suggestion's fold into another pending property change with the new `pendingRunPropertyChange` and `pendingParagraphPropertyChange` refusal reasons, which the canonical session reports as the skip reasons of the same names.
