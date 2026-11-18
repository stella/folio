---
"@stll/docx-core": patch
---

Keep a document's own range boundaries in place when rejecting changes joins a seam beside them: a rejected deletion's boundary content, and a boundary after a cut whose pieces the same rejection removes, no longer move into the neighbouring record.
