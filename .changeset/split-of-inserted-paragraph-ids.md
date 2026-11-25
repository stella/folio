---
"@stll/docx-core": patch
---

Rejecting a split of an inserted paragraph together with the insertion, in either order, gives the paragraph after it back under its own id. A rejected break keeps the paragraph before it only while the paragraph after it still holds the mark it took over.
