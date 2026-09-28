---
"@stll/folio-core": minor
"@stll/folio-agents": minor
"@stll/folio-react": patch
"@stll/folio-vue": patch
---

Document operations no longer commit a result the save would refuse. An operation whose `numbering.numId` names an instance the document does not define (`insertAfterBlock`, `insertBeforeBlock`, `setBlockParagraphProperties`, `splitBlock`, `mergeBlockWithNext`) is skipped before anything is applied with the new `missingNumbering` reason (`retryable: true`, `recovery: "refreshDocument"`), in every mode and through `suggest_changes`. Every batch result is then checked with the save-time model validator before it is committed; an operation whose result would not save is skipped with the new `invalidResult` reason and nothing from it reaches the document, the live editors included. Skipped operations and issues carry an optional `message` with the detail (the undefined instance and the ones the document defines, or the validator's path and reason). The React and Vue editors keep only the comments of applied operations, and the Vue editor's `applyAIEditOperations` now goes through the same checked applier as `applyDocumentOperations`.
