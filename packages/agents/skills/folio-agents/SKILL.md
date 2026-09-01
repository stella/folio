---
name: folio-agents
description: >-
  Integrate and use @stll/folio-agents, the framework-neutral LLM tool layer
  over @stll/folio-core: function-calling tools that read and mutate a .docx
  document, with every mutation landing as a tracked change or comment
  pending human review. Load this skill when wiring folio's tools into an
  agent's tool-use loop (TanStack AI, Vercel AI SDK, raw Anthropic/OpenAI
  SDKs, or a custom loop), choosing between the headless reviewer bridge and
  the live-editor-ref bridge, or summarizing what changed between two
  document versions.
metadata:
  type: core
  library: "@stll/folio-agents"
  library_version: "0.0.0"
---

# @stll/folio-agents

Framework-neutral LLM tool layer over `@stll/folio-core`'s ai-edits engine.
Gives a model function-calling tools to read and mutate a `.docx` document;
every mutation (`add_comment`, `suggest_changes`) lands as a tracked change or
comment pending human review — nothing is silently finalized.

## Integration recipe

1. `getFolioToolDefinitions(options?)` returns `FolioAgentToolDefinition[]` —
   plain `{ name, description, inputSchema }` objects with a conservative JSON
   Schema subset (`type: "object"`, `properties`, `required`, `enum`,
   `additionalProperties: false`, plain arrays — safe for providers that
   reject uncommon keywords, e.g. Gemini's OpenAPI-3.0 subset). Pass
   `{ suggestChanges }` options to shape `suggest_changes` for the surface
   (see below) and pass the same object to `executeFolioToolCall`.
2. Register them with your framework:
   - TanStack AI: pass `inputSchema` directly as `toolDefinition({ inputSchema })` — no wrapper needed.
   - Raw Anthropic/OpenAI SDKs: `toAnthropicTools(defs)` / `toOpenAITools(defs)`.
   - Vercel AI SDK: wrap each `inputSchema` in `jsonSchema()` from `ai`.
3. In your tool-use loop, for every `tool_use` / `tool_call` the model emits:
   call `executeFolioToolCallUntyped(name, args, bridge, { suggestChanges })`
   with the same options object as step 1 → returns
   `{ ok: true, result } | { ok: false, error }`.
4. Feed that result straight back to the model as the tool result message —
   both branches are meant to reach the model, not just the `ok: true` one:
   error strings are plain language ("blockId not found; re-read...") that
   the model can act on directly.

## Bridge selection

A `FolioAgentBridge` is the structural seam `executeFolioToolCall` drives.
Two are shipped:

- `createReviewerBridge(FolioDocxReviewer.fromBuffer(...))` — server/headless,
  full capability. Use for a document that isn't open in a live editor (a
  backend job, a batch review pass). Defaults to `mode: "tracked-changes"`
  (pass `{ mode: "direct" }` to edit in place instead — rarely what you want
  for an unreviewed agent).
- `createEditorRefBridge({ ref, author, getComments, setComments })` — drives
  a live `DocxEditorRef` from `@stll/folio-react` (or anything structurally
  matching `FolioAgentEditorRefLike`). Comments live in host React state, not
  on the ref, hence the `getComments` / `setComments` pair (the same ones the
  host already passes to `DocxEditor`).

The editor-ref bridge's capability now depends on what the underlying
`DocxEditorRef` implements. Against a ref with the read surface
(`getTrackedChanges`, `getCommentAnchors`, `getSelectionText`, `getPageText`),
it has full parity with the headless one: `read_changes` returns real tracked
changes, comment entries carry a resolved `blockId` / `quote`, and `read_page`
/ `read_selection` work against the live view. Against an older ref that
predates those methods, each falls back independently: `read_changes` returns
`[]`, comment entries keep `blockId: null` / `quote: ""`, and `read_page` /
`read_selection` report an unsupported-capability error (see
`src/bridges/editor-ref.ts` for the exact per-method fallback). A bridge
omits an optional capability member entirely rather than implementing it as a
no-op, and `executeFolioToolCall` reports the corresponding tool call as
unsupported rather than throwing — this is how the headless reviewer bridge
signals `read_page` / `read_selection` / `scroll_to_block` as unsupported too,
since a headless document has no live page/selection/scroll surface at all.

## Configuring `suggest_changes`

`suggest_changes` is configured per host surface, never re-derived. Pass one
`FolioSuggestChangesOptions` object to both `getFolioToolDefinitions({
suggestChanges })` and `executeFolioToolCall(name, args, bridge, {
suggestChanges })`:

- `operationTypes`: any subset of the contract's operation types (default:
  everything except `commentOnBlock`, covered by `add_comment`, and
  `insertSignatureTable`, which is direct-only). The schema advertises only
  the properties those types accept.
- `reviewMeta: "required"`: every operation must carry `severity` and `area`
  (a review queue that sorts and groups needs them). Default `"optional"`.
- `maxOperations`: 1 to 200 per call, default 50.
- `documentVersion: { current }`: the model must echo the token; the executor
  compares it to `bridge.getDocumentVersion()` and skips the whole batch with
  `documentVersionMismatch` when the document moved on.

`describeSuggestChangesCapabilities(options)` returns the capability text the
tool description carries; paste it into a system prompt instead of
hand-writing a "what this tool can do" paragraph that drifts.

The parser decodes leniently (a `kind` key, an operation given as a JSON
string, a property that does not apply to the type) and reports each step in
the result's `normalizations`; it then enforces the caps, mints ids unique
across calls, and delegates every per-operation rule to
`parseFolioDocumentOperationBatch` in `@stll/folio-core`.

## Host-managed review queue

A host that already has its own review-queue UX — its own place to store a
model's proposed edits pending approval, separate from folio's
tracked-changes redlines — writes a `FolioAgentBridge` whose
`applyDocumentOperations` enqueues the batch and returns `status: "queued"`
with the operations listed in `queued` (build `receipts` with
`getFolioDocumentOperationReceipts` from `@stll/folio-core/server`). The
executor then produces the standard envelope for the model (`queued`,
`skipped` with plain-language reasons, `receipts`, `normalizations`) and keeps
the automatic `precondition` stamping and the document-version check. A
surface with nothing editable reports every operation as skipped with
`documentNotEditable`. `parseSuggestChangesInput` / `parseAddCommentInput`
stay exported for validation-only paths.

## Ground rules for the model

- Block ids (`blockId`) and comment ids (`commentId`) only ever come from a
  prior tool call in the same conversation — `read_document`, `find_text`, or
  `read_comments`. Never invent or reuse one from outside the conversation;
  ids are not guessable and change whenever the document's structure changes.
- `suggest_changes` operations that get skipped (`skipped: [{ id, reason }]`)
  return a plain-language reason, not a machine code (e.g. "the block changed
  since your snapshot; re-read the document and retry with fresh ids"). Treat
  a skip as a retry signal: re-read (`read_document` or `find_text`) and
  reissue with fresh ids, not a dead end.
- Successful mutations return input-ordered `receipts`. Use their typed
  `affected` targets to identify updated blocks, stable ranges, insertions,
  and created comments without inferring effects from document text.
- `suggest_changes` defaults to tracked-changes mode on both shipped bridges
  — it proposes redlines for a human to accept or reject, never edits the
  visible text directly.
- The operation kinds a surface accepts are exactly the ones in its
  `suggest_changes` schema's `type` enum (by default `replaceInBlock`,
  `replaceRange`, `commentOnRange`, `formatRange`, `insertAfterBlock`,
  `insertBeforeBlock`, `replaceBlock`, `deleteBlock`, `insertTableRow`,
  `deleteTableRow`, `insertTableColumn`, `deleteTableColumn`,
  `mergeTableCells`, `splitTableCell`) plus comment/reply/resolve. Do not
  invent an operation kind or a directive marker; if a document needs a
  structural operation the contract lacks, extend `@stll/folio-core`'s
  ai-edits engine and the contract, then the schema follows.
- A `queued` list in a `suggest_changes` result means the host parked those
  operations in its own review queue; treat it like `applied` for the purpose
  of "the edit has been proposed", never as a failure.
- **Untrusted documents:** `read_document`, `read_comments`, `read_changes`,
  and `find_text` return document content verbatim, so a `.docx` from an
  untrusted party can inject prompt instructions straight into the model's
  context via its text — hosts should treat any document-derived tool result
  as untrusted model input. Mutations stay safe regardless: `suggest_changes`
  and `add_comment` only ever produce tracked changes or comments pending
  human review, so an injected instruction can propose an edit, never apply
  one.

## Summarizing what changed

Two different questions, two different tools:

- **Pending redlines in one document** — what a reviewer would see as tracked
  changes right now: use the `read_changes` tool on a capable bridge or call
  `reviewer.getChanges()` directly.
- **Between two saved `.docx` versions** — `compareDocxVersions(previousBuffer,
currentBuffer)` + `formatVersionDiffForLLM(diff)`. Both are plain async
  functions, not tool definitions: a model can't attach two buffers to one
  tool call (they aren't JSON-serializable arguments). Wrap them in a
  host-side tool keyed by version identifiers instead — the model calls it
  with two version ids, your backend resolves the buffers and returns the
  formatted diff text.
- **Reviewable output between two saved versions** —
  `generateRedlineDocx(previousBuffer, currentBuffer, options)` returns the
  generated package, operation results, unmatched stories, and a structured
  package-metadata privacy report.

`compareDocxVersions` compares the AS-ACCEPTED view of each buffer: any
tracked changes already pending in either one count as already applied before
the diff runs, so diffing two versions that each carry their own uncommitted
redlines still produces a clean diff instead of raw markup noise.

## Reference

Tool names, full JSON Schemas, and the headless/live-editor quickstarts are
in this package's README.md. Read `src/tools.ts` for the exact input schema
of every tool and `src/execute.ts` for the skip-reason vocabulary.
