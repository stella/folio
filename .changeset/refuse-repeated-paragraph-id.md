---
"@stll/docx-core": patch
---

Applying an operation refuses an edit whose document would hold one paragraph id twice, with the seed contract's `duplicateBlockId` reason. The check reads a paragraph id index of the package, derived where it can from a package indexed earlier, so it usually walks only what the edit changed.
