---
"@stll/docx-core": patch
---

Rejecting content that cut a pending insertion joins the insertion's two pieces again, so rejecting those pieces later leaves the same document as rejecting everything at once.
