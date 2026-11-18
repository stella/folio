---
"@stll/docx-core": patch
---

Refuse concurrent direct and tracked paragraph insertions after the same anchor, and direct insertions after a pending first-survivor paragraph mark, so rejecting a tracked insertion cannot discard a direct paragraph's formatting or merge away its paragraph boundary.
