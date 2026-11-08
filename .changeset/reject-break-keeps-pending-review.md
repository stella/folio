---
"@stll/docx-core": patch
---

Rejecting only the inserted paragraph mark of a paragraph that keeps its identity, after paragraphs were inserted at the end of a container, now keeps the next paragraph's pending mark and pending property changes instead of dropping them.
