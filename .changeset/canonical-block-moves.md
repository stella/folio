---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

Tracked block moves now apply in the experimental canonical session. A `deleteBlock` and an `insertAfterBlock` or `insertBeforeBlock` that share a `moveId` in one batch are recorded as `w:moveFrom` and `w:moveTo`, paired as the default engine pairs them, and an unpaired id is reported as `unpairedMove`. `INSERT_BLOCKS` and `DELETE_BLOCKS` take `move`. A block deletion the following paragraph survives now removes the paragraph by its identity.
