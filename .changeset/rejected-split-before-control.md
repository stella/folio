---
"@stll/docx-core": patch
---

Rejecting a paragraph split together with the paragraph's own inserted break and a pending property change, before a block content control or another block, restores the paragraph's properties: a paragraph joining it in the same resolution does not stand for one that stays, also with a range boundary between them.
