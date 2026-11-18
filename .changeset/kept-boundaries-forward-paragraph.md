---
"@stll/docx-core": patch
---

Remove a paragraph that goes whole even when no paragraph stands before it: the range boundaries and comment anchors it holds move to the start of the first paragraph after it in the same list, instead of keeping the paragraph to carry them.
