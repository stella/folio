---
"@stll/folio-react": patch
---

In a canonical session the editor holds the session's own document as its document state from the moment the editor is ready, not from the first settled edit, so the first keys after loading no longer convert and measure the whole document again. `PagedEditorRef.readDocumentModel()` reads that model without a copy.
