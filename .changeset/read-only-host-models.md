---
"@stll/folio-core": patch
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Both editors hold their document model typed read-only (`ReadonlyDocument`): the React editor's undo history and `PagedEditor`'s `document` prop, and the Vue editor's `modelDocument` and the toolbar, ruler, page setup and watermark values derived from it. `readModel` in `@stll/folio-core/types/documentModel` is the one place a model is read as a plain `Document`, for readers that never write it; a writer copies the model first. `cloneDocumentWithParagraphPropertySources` accepts a model, `cloneCommentsWithParagraphPropertySources` copies a comment list, the find-and-replace and table-action writers take a model and return a document that shares nothing with it, and `getDocumentWatermark` returns a read-only watermark for a read-only document. `resolveEffectiveSectionProperties` gives read-only section properties for a read-only document, and the table selection's `table` and `context.table` are typed read-only.

A host gets one copy of the editor's document per settled change, in both sessions, and never the document the editor holds: the change listeners (`onChange` and `docChange`) and the Vue editor's `getDocument`. A note edit, and in the React editor a margin drag, reaches the host once it settles, as a body edit does. The Vue editor's `readDocument` reads the current document without a copy, and a document a Vue host loads or sets is held as raw data even when it holds reactive proxies.

A Vue host that passes `modelDocument.value` on where a `Document` is expected reads it through `readModel`, or copies it first when it writes.
