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

| Tool                   | What it does                                                           |
| ---------------------- | ---------------------------------------------------------------------- |
| `read_document`        | Read the document body as `{ blockId, kind, text }` blocks             |
| `get_document_outline` | Read heading hierarchy and stable section handles                      |
| `read_section`         | Read a bounded logical section, with block cursor pagination           |
| `list_stories`         | List main, header, footer, footnote, and endnote story handles         |
| `read_story`           | Read one story by its typed handle                                     |
| `find_text`            | Search by document, section, real page, or story scope                 |
| `read_comments`        | Read comment threads (author, text, resolved, anchored block, replies) |
| `read_changes`         | Read pending tracked changes (insertions/deletions) awaiting review    |
| `add_comment`          | Attach a comment to a block, optionally quoting specific text          |
| `suggest_changes`      | Propose block or stable-range edits as tracked changes                 |
| `reply_comment`        | Reply to a comment thread                                              |
| `resolve_comment`      | Resolve or reopen a comment thread                                     |
| `read_page`            | Read a page's plain text (live editor only)                            |
| `read_selection`       | Read the current text selection (live editor only)                     |
| `scroll_to_block`      | Scroll the live editor to a block (live editor only)                   |
| `show_in_document`     | Reveal a stable block or exact text range (live editor only)           |

Block ids and comment ids always come from a prior tool call
(`read_document`, `find_text`, `read_comments`) within the same conversation —
never guess them. `suggest_changes` reports a plain-language reason when an
operation is skipped (e.g. the block changed since it was last read), so the
model can re-read and retry. Successful mutation results include `receipts`
that identify affected blocks, ranges, insertions, and created comments.

For document questions, start with `get_document_outline`, then call
`read_section` or scoped `find_text`. This keeps unrelated contract text out
of model context. Section handles use Folio block identities, heading depth,
and a text hash, so structurally changed, renamed, or deleted headings fail
stale instead of resolving to the wrong content. Page scopes and page numbers
use Folio's live layout; they are never approximated from character counts and
therefore require a live, paginated editor.

### Untrusted documents

`read_document`, `read_section`, `read_story`, `read_page`, `read_comments`,
`read_changes`, and `find_text` return document content verbatim. If a `.docx`
comes from an untrusted party, its text can carry prompt-injection payloads
straight into the model's context —
treat any document-derived tool result as untrusted model input, the same way
you would treat a fetched web page. Mutations stay safe by design regardless:
`suggest_changes` and `add_comment` land as tracked changes or comments
pending human review, so an injected instruction can propose an edit but
cannot silently apply one.

## Headless quickstart

```ts
import { FolioDocxReviewer } from "@stll/folio-core/server";
import {
  createReviewerBridge,
  executeFolioToolCallUntyped,
  getFolioToolDefinitions,
  toAnthropicTools,
} from "@stll/folio-agents";

const reviewer = await FolioDocxReviewer.fromBuffer(docxBuffer, { author: "AI" });
const bridge = createReviewerBridge(reviewer);
const tools = toAnthropicTools(getFolioToolDefinitions());

// Inside your tool-use loop, for each tool_use block the model emits:
const result = executeFolioToolCallUntyped(toolName, toolInput, bridge);
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

On a `DocxEditorRef` that implements the read surface (`getTrackedChanges`,
`getCommentAnchors`, `getSelectionText`, `getPageText`, `getTargetPage`, and
`showInDocument`), the editor-ref bridge has full parity with the headless one:
`read_changes` returns real tracked changes, comment entries carry a resolved
`blockId` / `quote`, and `read_page` / `read_selection` work against the live
view. Against an older ref that predates those methods, the bridge degrades
per-member: `read_changes`
returns `[]`, comment entries fall back to `blockId: null` / `quote: ""`, and
`read_page` / `read_selection` report an unsupported-capability error — see
`src/bridges/editor-ref.ts` for the exact fallback per method.

### Configuring `suggest_changes`

`suggest_changes` is a projection of the document-operation contract that a
host configures per surface instead of re-deriving. Pass the same options to
`getFolioToolDefinitions` (they shape the JSON Schema the model sees) and to
`executeFolioToolCall` (they drive the parser), so the two cannot drift:

```ts
import {
  describeSuggestChangesCapabilities,
  executeFolioToolCall,
  getFolioToolDefinitions,
  type FolioSuggestChangesOptions,
} from "@stll/folio-agents";

const suggestChanges: FolioSuggestChangesOptions = {
  // Any subset of the contract types, e.g. a text-only surface:
  operationTypes: ["replaceInBlock", "replaceBlock", "deleteBlock"],
  // Force `severity` and `area` on every operation for a review queue:
  reviewMeta: "required",
  // Per-call cap, 1 to 200 (default 50):
  maxOperations: 200,
  // Pin the batch to a host document version the model must echo back:
  documentVersion: { current: entityVersionId },
};

const tools = getFolioToolDefinitions({ suggestChanges });
const result = executeFolioToolCall("suggest_changes", args, bridge, { suggestChanges });
// `describeSuggestChangesCapabilities({ ...suggestChanges })` returns the same
// capability text the tool description carries, for reuse in a system prompt.
```

Defaults: every contract type except `commentOnBlock` (use `add_comment`) and
`insertSignatureTable` (direct-only, so never a tracked change); review
metadata optional; 50 operations per call; no version pin. The schema
advertises only the properties the allowed types accept, derived from
`FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE` in `@stll/folio-core`.

The parser is a front door over the contract parser, not a second parser. It
decodes leniently first (a `kind` key read as `type`, an operation supplied as
a JSON string, a property that does not apply to the operation type) and
reports every such step in the result's `normalizations`, then enforces the
caps and the options above, mints ids that stay unique across calls
(`op-<nonce>-<n>`), wraps plain-string `comment`s, and hands the batch to
`parseFolioDocumentOperationBatch`.

With `documentVersion` set, the batch carries `precondition.documentVersion`
and the executor compares it to the bridge's `getDocumentVersion()` before
anything is applied: a mismatch skips every operation with
`documentVersionMismatch`, and a bridge without `getDocumentVersion` refuses
version-pinned batches instead of guessing.

### Host-managed review queue

A host with its own review-queue UX (its own place to store proposed edits
pending approval, distinct from folio's tracked-changes redlines) implements
`FolioAgentBridge` with an `applyDocumentOperations` that enqueues instead of
applying and reports the operations as `queued`. The executor then returns
the same envelope the model gets everywhere else, receipts and skip prose
included, and the host keeps the automatic `precondition` stamping and the
document-version check:

```ts
import { getFolioDocumentOperationReceipts } from "@stll/folio-core/server";
import type { FolioAgentBridge } from "@stll/folio-agents";

const queueBridge: FolioAgentBridge = {
  snapshot: () => lastSnapshotShownToTheModel,
  getDocumentVersion: () => currentEntityVersionId,
  applyDocumentOperations: (batch) => {
    const queued = batch.operations.map(({ id }) => ({ id }));
    reviewQueue.enqueue(batch.operations);
    return {
      version: batch.version,
      status: "queued",
      applied: [],
      queued,
      skipped: [],
      issues: [],
      receipts: getFolioDocumentOperationReceipts(batch.operations, queued),
      undoHandle: null,
    };
  },
  getComments: () => [],
  getChanges: () => [],
  replyToComment: () => false,
  resolveComment: () => false,
};
```

A surface with no editable document at the moment reports every operation as
skipped with `documentNotEditable` instead. `parseSuggestChangesInput` and
`parseAddCommentInput` remain exported for a host that only needs validation.

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

To produce a reviewable package from the same pair of buffers, use
`generateRedlineDocx`. Optional package-metadata privacy transforms are applied
to the generated output and returned as a structured report:

```ts
import { generateRedlineDocx } from "@stll/folio-agents";

const result = await generateRedlineDocx(previousVersionBuffer, currentVersionBuffer, {
  privacy: { transforms: ["remove-attribution", "remove-timestamps"] },
});

await storeGeneratedPackage(result.buffer);
console.log(result.privacyReport);
```

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
import { executeFolioToolCallUntyped, getFolioToolDefinitions } from "@stll/folio-agents";

const tools = Object.fromEntries(
  getFolioToolDefinitions().map((def) => [
    def.name,
    tool({
      description: def.description,
      inputSchema: jsonSchema(def.inputSchema),
      execute: async (input) => executeFolioToolCallUntyped(def.name, input, bridge),
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
