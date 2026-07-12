---
"@stll/folio-agents": minor
---

Export a standard-schema projection of the versioned document-operation contract: `FOLIO_DOCUMENT_OPERATION_JSON_SCHEMA` (full operation union), `FOLIO_DOCUMENT_OPERATION_BATCH_JSON_SCHEMA` (versioned batch envelope), and `folioDocumentOperationBatchSchema` — a Standard Schema V1 object whose validation delegates to `parseFolioDocumentOperationBatch` (the parser stays the single source of truth) with the JSON schema attached for LLM tool definitions. The `suggest_changes` tool schema now derives from the shared constants, with its deliberate narrowings expressed explicitly; downstream consumers can drop hand-maintained mirrors of the contract.
