---
"@stll/folio-core": patch
---

A batch with two `setBlockParagraphProperties` (or a style-only `replaceBlock`) on one paragraph now refuses the later one as `overlappingOperation`, in direct and tracked mode alike. Applied directly, both used to report applied while the earlier one's values won; tracked, the earlier one was refused as `pendingParagraphPropertyChange` and the later one kept.
