---
"@stll/folio-agents": minor
"@stll/folio-core": minor
---

Add `@stll/folio-agents`: a framework-neutral LLM tool layer over folio-core's AI-edits engine — provider-neutral function-calling tool definitions (`read_document`, `find_text`, `read_comments`, `read_changes`, `add_comment`, `suggest_changes`, `reply_comment`, `resolve_comment`, plus live-editor capability tools), an `executeFolioToolCall` executor, Anthropic/OpenAI schema mappers, and bridges for both the headless `FolioDocxReviewer` and a live `DocxEditorRef`. In core, `FolioDocxReviewer` gains `resolveComment(commentId, { resolved? })`, which round-trips `w15:done` through `getComments` and `toBuffer`. It also gains a document version-diff engine — `compareDocxVersions` aligns two `.docx` buffers block by block (stable ids, then text LCS, then positional fallback) into added/deleted/modified changes, and `formatVersionDiffForLLM` renders the result as compact, deterministic text for a model prompt.
