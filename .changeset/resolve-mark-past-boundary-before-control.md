---
"@stll/docx-core": patch
---

Resolving a paragraph mark with bookmark or other range boundaries standing after it now resolves as it does with nothing between. Before a block content control or custom XML, an empty paragraph goes and the boundaries stay, and still-pending words run into the control's first paragraph with the boundaries where the two meet. Before only boundaries at the container's end, the paragraph stays, so the container still ends with a paragraph.
