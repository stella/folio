---
"@stll/docx-core": patch
---

Rejecting changes that remove the pieces a paragraph split left at both of its edges now ends that split's cut there, so rejecting the split's mark joins the paragraphs as plain neighbours: an empty wrapper the author made beside the removed piece stays, and a boundary kept from an insertion at that edge stays beside the next record, whether the revisions are rejected at once or one at a time. A split seam with pending insertions standing at it is again kept on them when the record's other piece lies beyond a later paragraph, so the record joins again once they are rejected.
