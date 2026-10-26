---
"@stll/folio-core": patch
---

Give a paragraph without a `w14:paraId` the same id in the live editor and in the headless reviewer: one load-time allocator derives it from the paragraph's text and position for both, so a block id read in the editor resolves when an edit is applied headlessly.
