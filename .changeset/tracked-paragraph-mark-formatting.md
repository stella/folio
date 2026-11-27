---
"@stll/docx-core": patch
"@stll/folio-core": patch
---

A suggested change to a paragraph mark's run formatting is now recorded as its own revision (`w:rPrChange` under `w:pPr/w:rPr`) instead of being refused. It accepts, rejects and undoes apart from any paragraph property change, folds later edits by the same suggestion, follows the mark through tracked splits, insertions and pastes, and survives a save.
