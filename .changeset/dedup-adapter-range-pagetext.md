---
"@stll/folio-core": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Hoist the `getPageTextFromLayout` and AI-edit block-range helpers (`resolveFolioAIBlockRange`, `clampRangeToDocSize`) into framework-neutral core modules (`@stll/folio-core/paged-layout/pageText`, `@stll/folio-core/ai-edits/blockRange`). Both were previously duplicated verbatim in the React and Vue adapters; they now share one implementation (and one test suite) in core, so the two adapters can never drift on page-text extraction or block-range resolution.
