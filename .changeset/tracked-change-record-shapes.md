---
"@stll/folio-core": minor
"@stll/docx-core": minor
---

Suggesting the deletion of a paragraph that holds a point comment now works: the deletion takes the comment's reference with it, accepting removes the comment as deleting the paragraph directly does, and rejecting keeps it. Tracked insertions and deletions inside a simple field's result can be accepted and rejected instead of being refused.
