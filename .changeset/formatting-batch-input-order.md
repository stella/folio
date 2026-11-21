---
"@stll/docx-core": minor
"@stll/folio-core": patch
---

Run a batch's formatting operations ahead of its comments and in input order, so a later write of a property replaces an earlier one wherever their ranges meet, as applying them one at a time does. A batch order address of an annotation names its kind (`BATCH_ANNOTATION_KINDS`). A canonical batch reports an operation that the batch's earlier operations leave with nothing to change as `noopOperation`.
