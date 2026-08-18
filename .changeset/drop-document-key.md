---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Remove the `documentKey` prop from the React and Vue `DocxEditor` (deprecated and ignored since the previous release) and the `useDocxEditor` option. Both adapters now hand the hidden-editor manager a per-load identity from their own loaders (`HiddenEditorManagerDeps.getDocumentIdentity`, required); the document-metadata fallback signature is gone.
