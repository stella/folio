---
"@stll/folio-core": patch
---

On the canonical session, a paragraph or table inserted before or after the paragraph shown in an empty content control goes beside the control, which stays as the source wrote it; any other operation on that paragraph is refused as `unsupportedBlock` with a message naming the control, not as a missing block.
