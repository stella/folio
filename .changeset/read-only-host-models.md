---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Both editors hold their document model typed read-only (`ReadonlyDocument`): the React editor's undo history and `PagedEditor`'s `document` prop, and the Vue editor's `modelDocument` and the toolbar, ruler, page setup and watermark values derived from it. `readModel` in `@stll/folio-core/types/documentModel` is the one place a model is read as a plain `Document`, for readers that never write it; a writer copies the model first. `cloneDocumentWithParagraphPropertySources` accepts a model, `cloneCommentsWithParagraphPropertySources` copies a comment list, the find-and-replace and table-action writers take a model and return a document that shares nothing with it, and `getDocumentWatermark` returns a read-only watermark for a read-only document. The Vue editor's change notifications and `getDocument` now always hand the host a copy, in the default session too.
