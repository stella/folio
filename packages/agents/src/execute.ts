import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioAIBlock,
  type FolioDocumentOperation,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "./bridge";
import {
  explainTextTooLong,
  MAX_OPERATION_TEXT_LENGTH,
  parseAddCommentInput,
  parseSuggestChangesInput,
} from "./parse";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentApplyOperationsSummary,
  FolioAgentCommentFilter,
  FolioAgentFindTextResult,
  FolioAgentTextMatch,
  FolioToolCallResult,
} from "./types";

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0;

const ok = (result: unknown): FolioToolCallResult => ({ ok: true, result });
const fail = (error: string): FolioToolCallResult => ({ ok: false, error });

const VALID_TOOL_NAMES: readonly string[] = Object.values(FOLIO_AGENT_TOOL_NAMES);

/** `find_text` requires a short `query` and caps how much of a large match set it returns in one call. */
const MAX_QUERY_LENGTH = 1_000;
const MAX_FIND_MATCHES = 200;

/**
 * Execute one tool call against a {@link FolioAgentBridge}. Every EXPECTED
 * failure — bad arguments, an unknown tool name, a capability the bridge does
 * not implement, a domain-level skip — comes back as `{ ok: false, error }`
 * with a message meant to be fed straight back to the model; it is never
 * thrown. This function is the sole boundary layer allowed to catch an
 * unexpected throw from the bridge and fold it into that same shape (per
 * AGENTS.md, try/catch is acceptable at boundary layers).
 *
 * Every bridge member used here ({@link FolioDocxReviewer} and the live-editor
 * ref) is synchronous, so this stays synchronous too.
 */
export const executeFolioToolCall = (
  name: string,
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResult => {
  if (!VALID_TOOL_NAMES.includes(name)) {
    return fail(`Unknown tool "${name}". Valid tools: ${VALID_TOOL_NAMES.join(", ")}.`);
  }

  try {
    return dispatch(name, args, bridge);
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
};

const dispatch = (name: string, args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (name === FOLIO_AGENT_TOOL_NAMES.readDocument) {
    return readDocument(bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.findText) {
    return findText(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.readComments) {
    return readComments(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.readChanges) {
    return ok(bridge.getChanges());
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.addComment) {
    return addComment(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.suggestChanges) {
    return suggestChanges(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.replyComment) {
    return replyComment(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.resolveComment) {
    return resolveComment(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.readPage) {
    return readPage(args, bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.readSelection) {
    return readSelection(bridge);
  }
  if (name === FOLIO_AGENT_TOOL_NAMES.scrollToBlock) {
    return scrollToBlock(args, bridge);
  }
  // Unreachable: `name` was checked against VALID_TOOL_NAMES above.
  return fail(`Unknown tool "${name}".`);
};

const readDocument = (bridge: FolioAgentBridge): FolioToolCallResult => {
  const { blocks } = bridge.snapshot();
  return ok(blocks.map((block) => ({ blockId: block.id, kind: block.kind, text: block.text })));
};

const CONTEXT_RADIUS = 40;

/** All matches for `query`, capped at {@link MAX_FIND_MATCHES}; `totalMatches` still counts every occurrence. */
type FindTextMatches = { matches: FolioAgentTextMatch[]; totalMatches: number };

const findTextMatches = (
  blocks: FolioAIBlock[],
  query: string,
  matchCase: boolean,
): FindTextMatches => {
  const matches: FolioAgentTextMatch[] = [];
  let totalMatches = 0;
  const needle = matchCase ? query : query.toLowerCase();
  for (const block of blocks) {
    const haystack = matchCase ? block.text : block.text.toLowerCase();
    let occurrence = 0;
    let fromIndex = 0;
    for (;;) {
      const at = haystack.indexOf(needle, fromIndex);
      if (at === -1) {
        break;
      }
      totalMatches += 1;
      if (matches.length < MAX_FIND_MATCHES) {
        const contextStart = Math.max(0, at - CONTEXT_RADIUS);
        const contextEnd = Math.min(block.text.length, at + query.length + CONTEXT_RADIUS);
        matches.push({
          blockId: block.id,
          occurrenceInBlock: occurrence,
          context: block.text.slice(contextStart, contextEnd),
        });
      }
      occurrence += 1;
      fromIndex = at + needle.length;
    }
  }
  return { matches, totalMatches };
};

const findText = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!isPlainObject(args)) {
    return fail("find_text expects an object with a non-empty `query` string.");
  }
  const query = args["query"];
  const matchCase = args["matchCase"];
  if (!isNonEmptyString(query)) {
    return fail("find_text requires a non-empty string `query`.");
  }
  if (query.length > MAX_QUERY_LENGTH) {
    return fail(
      `find_text's \`query\` is ${query.length.toLocaleString()} characters, over the ${MAX_QUERY_LENGTH.toLocaleString()}-character limit; shorten it.`,
    );
  }
  if (matchCase !== undefined && typeof matchCase !== "boolean") {
    return fail("find_text's `matchCase` must be a boolean when provided.");
  }
  const { blocks } = bridge.snapshot();
  const { matches, totalMatches } = findTextMatches(blocks, query, matchCase === true);
  const result: FolioAgentFindTextResult = {
    matches,
    truncated: totalMatches > matches.length,
    totalMatches,
  };
  return ok(result);
};

const isCommentFilter = (value: unknown): value is FolioAgentCommentFilter =>
  value === "all" || value === "open" || value === "resolved";

const readComments = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  const filter = isPlainObject(args) ? args["filter"] : undefined;
  if (filter !== undefined && !isCommentFilter(filter)) {
    return fail('read_comments\' `filter` must be one of "all", "open", "resolved" when provided.');
  }
  const comments = bridge.getComments();
  if (filter === undefined || filter === "all") {
    return ok(comments);
  }
  const wantResolved = filter === "resolved";
  return ok(comments.filter((comment) => comment.resolved === wantResolved));
};

/**
 * Turn a `FolioAIEditSkipReason` into a plain-language reason a model can act
 * on — what to change and retry, not just a machine code.
 */
const explainSkipReason = (reason: string): string => {
  if (reason === "missingBlock") {
    return "blockId not found; re-read the document (read_document or find_text) and retry with a fresh block id.";
  }
  if (reason === "changedBlock") {
    return "the block changed since your snapshot; re-read the document and retry with fresh ids.";
  }
  if (reason === "ambiguousFind") {
    return "`find` matches more than once in this block; narrow it so it matches exactly once, or use a replaceBlock operation instead.";
  }
  if (reason === "missingFind") {
    return "`find` was not found in this block; re-read the block's current text and retry.";
  }
  if (reason === "unsupportedBlock") {
    return "this block kind does not support this operation.";
  }
  if (reason === "unsupportedMode") {
    return "this operation does not support the requested mutation mode; inspect document operation capabilities and retry with a supported mode.";
  }
  if (reason === "preconditionFailed") {
    return "the target block changed after the operation was prepared; re-read the document and retry with a fresh operation.";
  }
  if (reason === "emptyOperation") {
    return "this operation has no effect; nothing to apply.";
  }
  if (reason === "noopOperation") {
    return "this operation would not change the document (the text already matches what you asked for).";
  }
  return reason;
};

const summarizeApplyResult = (result: {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: { id: string }[];
  skipped: { id: string; reason: string }[];
}): FolioAgentApplyOperationsSummary => ({
  version: result.version,
  applied: result.applied.map((entry) => ({ id: entry.id })),
  skipped: result.skipped.map((entry) => ({
    id: entry.id,
    reason: explainSkipReason(entry.reason),
  })),
});

const applyOperations = (
  bridge: FolioAgentBridge,
  operations: FolioDocumentOperation[],
): FolioAgentApplyOperationsSummary => {
  const snapshot = bridge.snapshot();
  const guardedOperations: FolioDocumentOperation[] = [];
  for (const operation of operations) {
    const blockTextHash = snapshot.anchors[operation.blockId]?.textHash;
    guardedOperations.push({
      ...operation,
      ...(blockTextHash !== undefined && { precondition: { blockTextHash } }),
    });
  }
  return summarizeApplyResult(
    bridge.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: guardedOperations,
    }),
  );
};

const addComment = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  const parsed = parseAddCommentInput(args);
  if (!parsed.ok) {
    return fail(parsed.error);
  }
  return ok(applyOperations(bridge, [parsed.operation]));
};

const suggestChanges = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  const parsed = parseSuggestChangesInput(args);
  if (!parsed.ok) {
    return fail(parsed.error);
  }
  return ok(applyOperations(bridge, parsed.operations));
};

const replyComment = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!isPlainObject(args)) {
    return fail("reply_comment expects an object with `commentId` and `text` strings.");
  }
  const commentId = args["commentId"];
  const text = args["text"];
  if (!isNonEmptyString(commentId)) {
    return fail("reply_comment requires a non-empty string `commentId`.");
  }
  if (!isNonEmptyString(text)) {
    return fail("reply_comment requires a non-empty string `text`.");
  }
  if (text.length > MAX_OPERATION_TEXT_LENGTH) {
    return fail(explainTextTooLong("reply_comment's `text`", text.length));
  }
  const replied = bridge.replyToComment(commentId, text);
  if (!replied) {
    return fail(`No comment with id "${commentId}" was found.`);
  }
  return ok({ replied: true });
};

const resolveComment = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!isPlainObject(args)) {
    return fail("resolve_comment expects an object with a `commentId` string.");
  }
  const commentId = args["commentId"];
  const reopen = args["reopen"];
  if (!isNonEmptyString(commentId)) {
    return fail("resolve_comment requires a non-empty string `commentId`.");
  }
  if (reopen !== undefined && typeof reopen !== "boolean") {
    return fail("resolve_comment's `reopen` must be a boolean when provided.");
  }
  const resolved = reopen !== true;
  const changed = bridge.resolveComment(commentId, resolved);
  if (!changed) {
    return fail(`No comment with id "${commentId}" was found.`);
  }
  return ok({ resolved });
};

const readPage = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!bridge.getPageText) {
    return fail("This editor surface does not support read_page (no live paginated view).");
  }
  const page = isPlainObject(args) ? args["page"] : undefined;
  if (typeof page !== "number" || !Number.isInteger(page) || page < 1) {
    return fail("read_page requires an integer `page` >= 1.");
  }
  const pageCount = bridge.getPageCount?.();
  if (pageCount !== undefined && page > pageCount) {
    return fail(`read_page's \`page\` (${page}) exceeds the document's page count (${pageCount}).`);
  }
  return ok({ page, text: bridge.getPageText(page) });
};

const readSelection = (bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!bridge.getSelectionText) {
    return fail("This editor surface does not support read_selection (no live selection).");
  }
  return ok({ text: bridge.getSelectionText() });
};

const scrollToBlock = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!bridge.scrollToBlock) {
    return fail("This editor surface does not support scroll_to_block (no live editor view).");
  }
  const blockId = isPlainObject(args) ? args["blockId"] : undefined;
  if (!isNonEmptyString(blockId)) {
    return fail("scroll_to_block requires a non-empty string `blockId`.");
  }
  return ok({ scrolled: bridge.scrollToBlock(blockId) });
};
