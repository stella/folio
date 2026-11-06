---
"@stll/folio-core": patch
---

Rejecting a tracked inserted paragraph no longer leaves a blank one behind when it ended a content control after a table, followed a paragraph with a text box drawn in it, or followed a table that ended the body. Accepting a removed paragraph break on a paragraph with a text box now joins the two paragraphs.
