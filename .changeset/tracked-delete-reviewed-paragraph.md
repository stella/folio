---
"@stll/docx-core": patch
---

Track the deletion of a paragraph whose mark a pending deletion already took, or whose content holds tracked changes: the earlier deletion record stands, and the content is deleted as a tracked range deletion over it is.
