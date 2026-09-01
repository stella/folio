---
"@stll/folio-agents": minor
---

Make `suggest_changes` host-configurable: `getFolioToolDefinitions({ suggestChanges })` and `executeFolioToolCall(name, args, bridge, { suggestChanges })` take `operationTypes`, `reviewMeta`, `maxOperations`, and `documentVersion` options that drive the schema, the parser, and `describeSuggestChangesCapabilities()` from one list. The parser now decodes leniently and reports `normalizations`, passes `severity`/`area` through, mints ids unique across calls, delegates per-operation rules to the core contract parser, and pins batches to a host document version checked against the new optional `FolioAgentBridge.getDocumentVersion()`. Results gain `queued` for host review-queue bridges; `suggestionId` and the batch `precondition` join the JSON Schema projections.
