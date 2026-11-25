---
"@stll/folio-core": minor
"@stll/folio-react": minor
"@stll/folio-vue": minor
---

Remove `applyFolioAIEditOperations`, `applyFolioDocumentOperations` and `ApplyFolioDocumentOperationsOptions` from the public exports. They applied operations to a ProseMirror view; apply operations through the editor ref's `applyDocumentOperations` / `applyAIEditOperations` or `FolioDocxReviewer`, which run on the canonical session.
