---
"@stll/folio-core": patch
---

A tracked or suggested move of the paragraph that ends its container to beside itself (a `deleteBlock` and an `insertAfterBlock` or `insertBeforeBlock` sharing a `moveId`) now applies in the canonical session instead of refusing one half. The moved words and the words moved there share one paragraph while pending, so no tracked mark stands on the container's final paragraph; accepting gives the direct result and rejecting gives the source's paragraphs back.
