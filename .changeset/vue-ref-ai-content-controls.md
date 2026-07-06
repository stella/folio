---
"@stll/folio-vue": minor
---

Wire the AI-edit and content-control `DocxEditorRef` methods in the Vue adapter
over `@stll/folio-core`, mirroring the React adapter. `createAIEditSnapshot`,
`applyAIEditOperations`, `acceptAIEditOperation`, `rejectAIEditOperation`,
`scrollToAIEditOperation`, `scrollToBlock`, `getContentControls`,
`scrollToContentControl`, `setContentControlContent`, `setContentControlValue`,
and `removeContentControl` were previously stubbed to `null` / `false` / `[]`;
they now drive the live editor view. Only `getEditorRef` and `hasPendingChanges`
remain deferred pending the Vue `PagedEditor` component and PM-serialized-vs-live
tracking.
