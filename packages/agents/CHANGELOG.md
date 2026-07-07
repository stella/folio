# @stll/folio-agents

## 0.1.0

### Minor Changes

- [#65](https://github.com/stella/folio/pull/65) [`ca9d64b`](https://github.com/stella/folio/commit/ca9d64bd36bbce78b3dad9aab72092cceebc4919) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Add `@stll/folio-agents`: a framework-neutral LLM tool layer over folio-core's AI-edits engine — provider-neutral function-calling tool definitions (`read_document`, `find_text`, `read_comments`, `read_changes`, `add_comment`, `suggest_changes`, `reply_comment`, `resolve_comment`, plus live-editor capability tools), an `executeFolioToolCall` executor, Anthropic/OpenAI schema mappers, and bridges for both the headless `FolioDocxReviewer` and a live `DocxEditorRef`. In core, `FolioDocxReviewer` gains `resolveComment(commentId, { resolved? })`, which round-trips `w15:done` through `getComments` and `toBuffer`. It also gains a document version-diff engine — `compareDocxVersions` aligns two `.docx` buffers block by block (stable ids, then text LCS, then positional fallback) into added/deleted/modified changes, and `formatVersionDiffForLLM` renders the result as compact, deterministic text for a model prompt.

- [#81](https://github.com/stella/folio/pull/81) [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c) Thanks [@jan-kubica](https://github.com/jan-kubica)! - Broaden the AI-edits read surface for live-editor reuse. Core gains pure `getTrackedChangesFromDoc` / `getCommentAnchorsFromDoc` readers and a `buildAnnotatedBlockText` redline renderer over any ProseMirror `doc`, a `getContentAsText({ annotated })` option that inlines tracked changes and comment anchors as `<ins>`/`<del>`/`<comment>` tags, and a read-only `getNotesAsText` surfacing header/footer and footnote/endnote text. The React and Vue `DocxEditorRef` gain a matching read surface — `getTrackedChanges`, `getCommentAnchors`, `getSelectionText`, and `getPageText(page)` — so a live-editor agent tool can read the current document state without a fresh AI-edit snapshot. `@stll/folio-agents` now exports `parseSuggestChangesInput` / `parseAddCommentInput` — the `suggest_changes` / `add_comment` argument-validation rules factored out of `executeFolioToolCall`, for hosts with their own review-queue UX — and its live-editor bridge now reads real tracked changes and comment anchors from that new `DocxEditorRef` surface instead of reporting them as unavailable.

### Patch Changes

- Updated dependencies [[`ca9d64b`](https://github.com/stella/folio/commit/ca9d64bd36bbce78b3dad9aab72092cceebc4919), [`439a8ea`](https://github.com/stella/folio/commit/439a8ea12227aafc49e99693755fc03577d4b54c), [`ce39713`](https://github.com/stella/folio/commit/ce397130771ec8f1271b7d0cfde9e6dc0367e857)]:
  - @stll/folio-core@0.3.0
