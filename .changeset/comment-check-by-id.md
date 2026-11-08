---
"@stll/docx-core": patch
---

The comment anchor check reads each comment's anchors from one pass over the document's anchors instead of filtering all of them once per comment, so documents with many comments check in time proportional to their anchors.
