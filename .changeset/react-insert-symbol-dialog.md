---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": patch
---

Port the Insert Symbol dialog to the React adapter, reaching full parity with Vue. The symbol catalog and search now live in a shared framework-neutral `@stll/folio-core/symbols` module (six categories, `filterSymbols`); the Vue dialog is refactored onto it so both adapters share one source of truth instead of duplicating the catalog. React gains `InsertSymbolDialog` (exported from the package root), an "Insert Symbol" toolbar button (`onInsertSymbol`), and inserts the chosen character at the cursor.
