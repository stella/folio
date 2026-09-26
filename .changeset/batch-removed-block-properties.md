---
"@stll/folio-core": patch
---

A `setBlockParagraphProperties` whose block an earlier operation of the same batch removed (a direct `deleteBlock`, or a merge that joined it away) is skipped as `missingBlock`. It used to throw a `RangeError` out of `applyDocumentOperations` / `suggest_changes`, or restyle the paragraph that took the removed block's place.
