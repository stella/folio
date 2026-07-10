---
"@stll/folio-vue": minor
---

Port `applyDocumentOperations` to the Vue `DocxEditorRef`. The Vue adapter now implements the versioned document-operation batch API at full parity with React (delegating to core's `applyFolioDocumentOperations`, mirroring the existing `applyAIEditOperations`), and re-exports `DocxEditorApplyDocumentOperationsOptions` from the package root. Previously this ref method existed only on the React adapter.
