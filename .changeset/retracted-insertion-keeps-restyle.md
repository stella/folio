---
"@stll/docx-core": patch
---

A tracked deletion that takes back an author's own paragraph inserted after a container's last paragraph keeps the pending property changes that author made to that last paragraph since, so accepting reads as deleting the inserted paragraph directly, and rejecting still restores the original.
