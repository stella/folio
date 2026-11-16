---
"@stll/docx-core": patch
---

Refuse concurrent direct and tracked paragraph insertions after the same anchor, and sequential transfers of a pending inserted mark to differently formatted content, so rejecting a tracked insertion cannot discard a direct paragraph's formatting. Same-format sequential insertions preserve the pending mark and remain reversible.
