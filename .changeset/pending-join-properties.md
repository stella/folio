---
"@stll/docx-core": patch
---

Deleting a paragraph break after an empty paragraph that a pending join runs words into keeps that paragraph's properties, in direct, tracked and suggested edits alike, so accepting a tracked deletion of a last paragraph leaves what the direct deletion does. A tracked deletion that retires the break before a container's pending end records those properties too, and one after a break already pending deletion leaves the last paragraph's properties as they are.
