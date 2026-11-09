---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Rejecting a tracked or suggested paragraph inserted after a paragraph that a block content control (or any block other than a paragraph or a table) follows no longer leaves an empty paragraph behind. As at the end of a container, the anchor's mark is recorded as the inserted one, so rejecting it joins the two paragraphs back into the anchor.
