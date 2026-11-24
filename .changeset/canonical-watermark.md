---
"@stll/docx-core": minor
"@stll/folio-core": minor
"@stll/folio-vue": patch
---

Edit watermarks in the canonical session: `watermarkOperations` (from `storyPartsOpOf`, in `controller/canonicalWatermark`) journals a watermark change as one operation, and the Vue watermark dialog applies it instead of refusing.
