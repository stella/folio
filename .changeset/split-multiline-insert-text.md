---
"@stll/folio-core": patch
"@stll/folio-agents": patch
---

Split `insertAfterBlock` / `insertBeforeBlock` text on line breaks into consecutive paragraphs instead of one paragraph with embedded newlines, and report the split as a `splitMultilineText` normalization (`FolioAIEditApplyResult.normalizations`, surfaced through `suggest_changes`' `normalizations`). Only the first paragraph keeps `styleId` / `inheritFormatting`; later paragraphs get body formatting.
