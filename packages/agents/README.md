# @stll/folio-agents

A framework-neutral (React-free, DOM-free) LLM tool layer over
[`@stll/folio-core`](https://www.npmjs.com/package/@stll/folio-core)'s AI-edits
engine: function-calling tool definitions and an executor so a model can read
and mutate a `.docx` document, with every mutation landing as a tracked change
or comment pending human review.

Part of [stella](https://github.com/stella/stella), an open-source legal workspace.

## Install

```sh
bun add @stll/folio-agents
```

`@stll/folio-core` is installed automatically as a dependency.

## Tools

| Tool              | What it does                                                                     |
| ----------------- | --------------------------------------------------------------------------------- |
| `read_document`    | Read the document body as `{ blockId, kind, text }` blocks                        |
| `find_text`        | Search block text for a string; returns block id, occurrence, and context per match |
| `read_comments`    | Read comment threads (author, text, resolved, anchored block, replies)            |
| `read_changes`     | Read pending tracked changes (insertions/deletions) awaiting review               |
| `add_comment`      | Attach a comment to a block, optionally quoting specific text                     |
| `suggest_changes`  | Propose `replaceInBlock` / `insertAfterBlock` / `insertBeforeBlock` / `replaceBlock` / `deleteBlock` edits as tracked changes |
| `reply_comment`    | Reply to a comment thread                                                         |
| `resolve_comment`  | Resolve or reopen a comment thread                                                |
| `read_page`        | Read a page's plain text (live editor only)                                       |
| `read_selection`   | Read the current text selection (live editor only)                                |
| `scroll_to_block`  | Scroll the live editor to a block (live editor only)                              |

Block ids and comment ids always come from a prior tool call
(`read_document`, `find_text`, `read_comments`) within the same conversation —
never guess them. `suggest_changes` reports a plain-language reason when an
operation is skipped (e.g. the block changed since it was last read), so the
model can re-read and retry.

## Headless quickstart

```ts
import { FolioDocxReviewer } from "@stll/folio-core/server";
import {
  createReviewerBridge,
  executeFolioToolCall,
  getFolioToolDefinitions,
  toAnthropicTools,
} from "@stll/folio-agents";

const reviewer = await FolioDocxReviewer.fromBuffer(docxBuffer, { author: "AI" });
const bridge = createReviewerBridge(reviewer);
const tools = toAnthropicTools(getFolioToolDefinitions());

// Inside your tool-use loop, for each tool_use block the model emits:
const result = executeFolioToolCall(toolName, toolInput, bridge);
// result: { ok: true, result } | { ok: false, error } — feed either back to the model.

const reviewedBuffer = await reviewer.toBuffer();
```

## Live-editor quickstart

```ts
import { createEditorRefBridge, executeFolioToolCall } from "@stll/folio-agents";

// `docxEditorRef` is a DocxEditorRef from @stll/folio-react (or any object
// structurally matching FolioAgentEditorRefLike).
const bridge = createEditorRefBridge({
  ref: docxEditorRef.current,
  author: "AI",
  getComments: () => comments,
  setComments: (next) => setComments(next),
});

const result = executeFolioToolCall("suggest_changes", { operations: [...] }, bridge);
```

The editor-ref bridge has known gaps versus the headless one: `read_changes`
always returns `[]` (no ref-level API enumerates tracked changes), and comment
entries carry no `blockId` / `quote` (no ref-level anchor lookup). `read_page`
and `read_selection` are unsupported on either bridge shipped here — see
`src/bridges/editor-ref.ts` for details.

## Summarizing changes

Two different questions come up under "what changed":

**1. Pending tracked changes in one document** — what a human reviewer would
see as redlines right now. Use the `read_changes` tool (or
`reviewer.getChanges()` directly) and hand the insertions/deletions to the
model:

```ts
const changes = reviewer.getChanges();
const prompt = `Summarize these pending edits for a reviewer:\n${changes
  .map((c) => `${c.type === "insertion" ? "+" : "-"} [${c.blockId}] ${c.text}`)
  .join("\n")}`;
```

**2. Between two saved versions** — what changed across two `.docx` buffers,
independent of whether either one has any tracked changes at all. Use
`compareDocxVersions` + `formatVersionDiffForLLM`:

```ts
import { compareDocxVersions, formatVersionDiffForLLM } from "@stll/folio-agents";

const diff = await compareDocxVersions(previousVersionBuffer, currentVersionBuffer);
const prompt = `Summarize what changed between these two document versions:\n${formatVersionDiffForLLM(diff)}`;
// -> feed `prompt` to your model as a normal user/system message.
```

Both recipes compare the AS-ACCEPTED view of a document: any tracked changes
already pending in a buffer count as already applied before the comparison
runs (`compareDocxVersions` parses each buffer through the same
`FolioDocxReviewer` snapshot `read_document` uses). Diffing two versions that
each have their own uncommitted redlines still produces a clean, readable diff
instead of raw markup noise.

`compareDocxVersions` ships as a plain async function, not a tool definition:
a model can describe a tool call, but it can't attach two document buffers to
one — buffers aren't JSON-serializable tool arguments a model could produce.
The natural shape is a host-side tool keyed by version identifiers instead
(e.g. a server tool the model calls with two stored version ids, which your
backend resolves to buffers, diffs, and returns the formatted text for).

## TanStack AI

TanStack AI's `toolDefinition` accepts a raw JSON Schema object as
`inputSchema`, so the definitions plug in without any wrapper:

```ts
import { toolDefinition } from "@tanstack/ai";
import { getFolioToolDefinitions } from "@stll/folio-agents";

const defs = getFolioToolDefinitions().map((def) =>
  toolDefinition({
    name: def.name,
    description: def.description,
    inputSchema: def.inputSchema,
  }),
);
// Client-executed tools: run executeFolioToolCall(name, args, bridge) where the
// live editor lives and report the payload back via your chat client's
// addToolResult; server-executed tools: chain .server((args) => ...) instead.
```

The schemas stay within a conservative JSON Schema subset (`type: "object"`,
`properties`, `required`, `enum`, `additionalProperties: false`, plain arrays).
Some providers (e.g. Gemini's OpenAPI-3.0 subset) reject less common keywords;
if your stack projects tool schemas through a provider-safe filter, these
definitions pass through it unchanged.

## Vercel AI SDK

This package ships no `ai` dependency; map its tool definitions with the AI
SDK's own `jsonSchema()` / `tool()` helpers:

```ts
import { jsonSchema, tool } from "ai";
import { executeFolioToolCall, getFolioToolDefinitions } from "@stll/folio-agents";

const tools = Object.fromEntries(
  getFolioToolDefinitions().map((def) => [
    def.name,
    tool({
      description: def.description,
      inputSchema: jsonSchema(def.inputSchema),
      execute: async (input) => executeFolioToolCall(def.name, input, bridge),
    }),
  ]),
);
```

## Acknowledgements

folio began as a private fork of [Eigenpal](https://eigenpal.com)'s
[docx-editor](https://github.com/eigenpal/docx-editor). The original license
and copyright are preserved in
[`NOTICE.md`](https://github.com/stella/folio/blob/main/packages/core/NOTICE.md).

## License

Apache-2.0
