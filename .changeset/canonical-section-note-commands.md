---
"@stll/folio-core": minor
"@stll/docx-core": minor
---

Section break and footnote and endnote commands run through canonical document operations in canonical sessions; every inserted note paragraph receives a fresh paragraph id. The no-op `extractSelectionContext` editor command is removed: call the exported `extractSelectionContext(state)` function instead. `@stll/docx-core/ops` exports `paragraphIdsIn`.
