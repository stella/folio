---
"@stll/docx-core": patch
---

A tracked deletion of whole paragraphs now takes paragraphs holding comment range anchors: the anchors stand outside the deletion, and accepting it hands them to the nearest paragraph that stays, as deleting the paragraphs directly does. Rejecting it restores the paragraphs as they were.
