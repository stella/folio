---
"@stll/docx-core": patch
---

Resolving a paragraph mark whose next block, past bookmark or other range boundaries standing between them, is a block content control or custom XML now resolves as it does with nothing between: an empty paragraph goes and the boundaries stay; still-pending words run into the control's first paragraph with the boundaries where the two meet.
