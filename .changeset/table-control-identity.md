---
"@stll/docx-core": patch
---

A table row or cell inside a content control no longer counts that control twice. The parser records such a control in the row's or cell's control stack and again in its wrapper order, and the identity check took the two for two controls carrying one id, so the experimental canonical session refused the document.
