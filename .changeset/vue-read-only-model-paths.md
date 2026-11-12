---
"@stll/folio-vue": patch
---

The Vue editor passes its document model on read-only to core's `createHiddenEditorState`, `runLayoutPipeline` and `waitForInitialLayoutFonts`, and reads it through `readModel` only where a reader takes plain styles, section properties or a base document. `getEditorRef().getDocument()` returns a `ReadonlyDocument`, as React's `PagedEditorRef.getDocument` does: it shares the model's parts, so a host that writes it copies it first (`DocxEditorRef.getDocument` hands a copy).
