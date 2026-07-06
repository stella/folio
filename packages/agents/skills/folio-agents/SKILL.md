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

1. `getFolioToolDefinitions()` returns `FolioAgentToolDefinition[]` — plain
   `{ name, description, inputSchema }` objects with a conservative JSON
   Schema subset (`type: "object"`, `properties`, `required`, `enum`,
   `additionalProperties: false`, plain arrays — safe for providers that
   reject uncommon keywords, e.g. Gemini's OpenAPI-3.0 subset).
2. Register them with your framework:
   - TanStack AI: pass `inputSchema` directly as `toolDefinition({ inputSchema })` — no wrapper needed.
   - Raw Anthropic/OpenAI SDKs: `toAnthropicTools(defs)` / `toOpenAITools(defs)`.
   - Vercel AI SDK: wrap each `inputSchema` in `jsonSchema()` from `ai`.
3. In your tool-use loop, for every `tool_use` / `tool_call` the model emits:
   call `executeFolioToolCall(name, args, bridge)` → returns
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

The editor-ref bridge has real capability gaps versus the headless one, not
bugs: `read_changes` always returns `[]` (no ref-level API enumerates tracked
changes from ProseMirror mark attributes), and comment entries carry
`blockId: null` / `quote: ""` (no ref-level anchor-resolution accessor). If
you're wiring tools for a live-editor host, do not register `read_changes` or
rely on comment anchors from that bridge until the ref surface grows — either
omit the tool from the model's tool list, or tell the model in its system
prompt that those fields are unavailable there. `read_page` and
`read_selection` are unsupported on both bridges shipped here (see
`src/bridges/editor-ref.ts` for the exact gap list); a bridge simply omits an
optional capability member, and `executeFolioToolCall` reports the
corresponding tool call as an unsupported-capability error rather than
throwing.

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
- `suggest_changes` defaults to tracked-changes mode on both shipped bridges
  — it proposes redlines for a human to accept or reject, never edits the
  visible text directly.
- The tool surface here is fixed: five edit operation kinds
  (`replaceInBlock`, `insertAfterBlock`, `insertBeforeBlock`, `replaceBlock`,
  `deleteBlock`) plus comment/reply/resolve. There is no
  `insertSignatureTable`-style structural op, and none should be invented —
  compose the five primitives, or extend `@stll/folio-core`'s ai-edits engine
  itself if a document needs a structural operation this package doesn't
  expose.
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
  changes right now: use the `read_changes` tool (headless bridge only; see
  the capability gap above) or call `reviewer.getChanges()` directly.
- **Between two saved `.docx` versions** — `compareDocxVersions(previousBuffer,
currentBuffer)` + `formatVersionDiffForLLM(diff)`. Both are plain async
  functions, not tool definitions: a model can't attach two buffers to one
  tool call (they aren't JSON-serializable arguments). Wrap them in a
  host-side tool keyed by version identifiers instead — the model calls it
  with two version ids, your backend resolves the buffers and returns the
  formatted diff text.

`compareDocxVersions` compares the AS-ACCEPTED view of each buffer: any
tracked changes already pending in either one count as already applied before
the diff runs, so diffing two versions that each carry their own uncommitted
redlines still produces a clean diff instead of raw markup noise.

## Reference

Tool names, full JSON Schemas, and the headless/live-editor quickstarts are
in this package's README.md. Read `src/tools.ts` for the exact input schema
of every tool and `src/execute.ts` for the skip-reason vocabulary.
