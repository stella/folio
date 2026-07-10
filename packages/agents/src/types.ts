import type {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FolioAITextRangeHandle,
  FolioDocumentOperationIssue,
  FolioDocumentOperationReceipt,
} from "@stll/folio-core/server";

/**
 * Provider-neutral tool-layer types. `FolioAgentToolDefinition` describes a
 * tool the same way every model provider ends up wanting it described (name +
 * prose description + a JSON-Schema-shaped input schema); `providers.ts` maps
 * that single shape onto each SDK's own tool-definition envelope.
 */

/** Every tool name this package exposes, as a stable string union (no enums). */
export const FOLIO_AGENT_TOOL_NAMES = {
  readDocument: "read_document",
  listStories: "list_stories",
  readStory: "read_story",
  findText: "find_text",
  readComments: "read_comments",
  readChanges: "read_changes",
  addComment: "add_comment",
  suggestChanges: "suggest_changes",
  replyComment: "reply_comment",
  resolveComment: "resolve_comment",
  readPage: "read_page",
  readSelection: "read_selection",
  scrollToBlock: "scroll_to_block",
} as const;

export type FolioAgentToolName =
  (typeof FOLIO_AGENT_TOOL_NAMES)[keyof typeof FOLIO_AGENT_TOOL_NAMES];

/**
 * A provider-neutral tool definition: name, an LLM-facing description (when to
 * call it, what it returns), and a JSON-Schema (draft-07-ish) object schema for
 * its arguments. `providers.ts` maps a list of these onto Anthropic's or
 * OpenAI's tool-definition shape.
 */
export type FolioAgentToolDefinition = {
  name: FolioAgentToolName;
  description: string;
  inputSchema: Record<string, unknown>;
};

/**
 * Result of executing one tool call. `ok: false` covers every EXPECTED failure
 * (bad arguments, an unknown tool, a capability the bridge does not support, a
 * domain-level skip) — never throw for these, since the whole point is to feed
 * the reason back to the model so it can retry or adjust. `executeFolioToolCall`
 * is the only boundary layer allowed to catch an unexpected throw and fold it
 * into this shape too.
 */
export type FolioToolCallResult = { ok: true; result: unknown } | { ok: false; error: string };

/** One document block as exposed to a model: id, kind, and its plain text. */
export type FolioAgentBlock = {
  blockId: string;
  kind: string;
  text: string;
};

/** One `find_text` match: which block, which occurrence within it, and a text window around the match. */
export type FolioAgentTextMatch = {
  blockId: string;
  /** Stable handle that can be passed directly to a `replaceRange` operation. */
  range: FolioAITextRangeHandle;
  /** 0-based index of this occurrence within its block (multiple matches in one block increment this). */
  occurrenceInBlock: number;
  /** The match plus ~40 characters of surrounding context on each side. */
  context: string;
};

/**
 * Result of {@link FOLIO_AGENT_TOOL_NAMES.findText}. `matches` is capped at a
 * fixed limit; when the query has more hits than that, `truncated` is `true`
 * and `totalMatches` reports the real count so the model knows to narrow the
 * query instead of assuming it saw everything.
 */
export type FolioAgentFindTextResult = {
  matches: FolioAgentTextMatch[];
  truncated: boolean;
  totalMatches: number;
};

/** One reply within a comment thread. */
export type FolioAgentCommentReply = {
  id: string;
  author: string;
  text: string;
};

/**
 * A comment thread, shaped to match what {@link FolioDocxReviewer.getComments}
 * (from `@stll/folio-core/server`) returns, minus fields a tool-calling model
 * has no use for (raw dates). `resolved` mirrors the underlying model's
 * `done` flag; `quote` mirrors `anchoredText` (the text the comment is
 * attached to, empty when unavailable).
 */
export type FolioAgentComment = {
  id: string;
  author: string;
  text: string;
  resolved: boolean;
  blockId: string | null;
  quote: string;
  replies: FolioAgentCommentReply[];
};

/** A pending tracked change (insertion or deletion), shaped to match {@link FolioDocxReviewer.getChanges}. */
export type FolioAgentChange = {
  id: string;
  type: "insertion" | "deletion";
  author: string;
  text: string;
  blockId: string | null;
};

/** Filter for {@link FOLIO_AGENT_TOOL_NAMES.readComments}. */
export type FolioAgentCommentFilter = "all" | "open" | "resolved";

/** Result of {@link FOLIO_AGENT_TOOL_NAMES.addComment} / `suggest_changes`. */
export type FolioAgentApplyOperationsSummary = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: { id: string }[];
  skipped: { id: string; reason: string }[];
  issues: FolioDocumentOperationIssue[];
  receipts: FolioDocumentOperationReceipt[];
};
