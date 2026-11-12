---
"@stll/docx-core": patch
---

Resolving revisions places the range boundaries that removed content leaves by the records that remain: rejecting content that cut a pending insertion joins the insertion's pieces again, and a wrapper resolved in the same pass no longer counts as a level, so resolving revisions at once or one operation at a time gives the same document.
