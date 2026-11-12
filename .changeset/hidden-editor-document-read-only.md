---
"@stll/folio-core": patch
"@stll/folio-vue": patch
---

The hidden editor's `getDocument`, and with it `FolioEditor.getDocument`, returns the document typed read-only: in the default session it is built over the model and shares every part no edit reached with it, and in a canonical session it is the session's own document, now read without a copy. Vue's `PagedEditorRef.getDocument` is typed read-only for the same reason. A writer copies the document first with `cloneDocumentWithParagraphPropertySources`.
