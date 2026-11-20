---
"@stll/docx-core": patch
---

Track the deletion of a paragraph that holds a comment's reference: the reference stands outside the deletion as the comment's range anchors do, and a direct deletion keeps it too, so a comment that goes on over earlier words keeps its reference.
