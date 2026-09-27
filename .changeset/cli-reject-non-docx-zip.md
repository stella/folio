---
"@stll/folio-cli": patch
---

Every command and MCP tool refuses, as `invalid_document` (exit 2), a file that is not a WordprocessingML package: one without `[Content_Types].xml`, a package relationship to a main document part, that part, a WordprocessingML main-document content type, or a `document` root. A ZIP of unrelated entries used to read as an empty document, and `folio save --from` committed it over a real one. `folio serve` now reports a document it cannot open with the document's own error code rather than `invalid_input`.
