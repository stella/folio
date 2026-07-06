import type { FolioAIBlock, FolioAIComment, FolioAIEditOperation } from "@stll/folio-core/server";

import type { FolioAgentBridge } from "./bridge";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentApplyOperationsSummary,
  FolioAgentCommentFilter,
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

const findTextMatches = (
  blocks: FolioAIBlock[],
  query: string,
  matchCase: boolean,
): FolioAgentTextMatch[] => {
  const matches: FolioAgentTextMatch[] = [];
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
      const contextStart = Math.max(0, at - CONTEXT_RADIUS);
      const contextEnd = Math.min(block.text.length, at + query.length + CONTEXT_RADIUS);
      matches.push({
        blockId: block.id,
        occurrenceInBlock: occurrence,
        context: block.text.slice(contextStart, contextEnd),
      });
      occurrence += 1;
      fromIndex = at + needle.length;
    }
  }
  return matches;
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
  if (matchCase !== undefined && typeof matchCase !== "boolean") {
    return fail("find_text's `matchCase` must be a boolean when provided.");
  }
  const { blocks } = bridge.snapshot();
  return ok(findTextMatches(blocks, query, matchCase === true));
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
  if (reason === "emptyOperation") {
    return "this operation has no effect; nothing to apply.";
  }
  if (reason === "noopOperation") {
    return "this operation would not change the document (the text already matches what you asked for).";
  }
  return reason;
};

const summarizeApplyResult = (result: {
  applied: { id: string }[];
  skipped: { id: string; reason: string }[];
}): FolioAgentApplyOperationsSummary => ({
  applied: result.applied.map((entry) => ({ id: entry.id })),
  skipped: result.skipped.map((entry) => ({ id: entry.id, reason: explainSkipReason(entry.reason) })),
});

const addComment = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!isPlainObject(args)) {
    return fail("add_comment expects an object with `blockId` and `text` strings.");
  }
  const blockId = args["blockId"];
  const quote = args["quote"];
  const text = args["text"];
  if (!isNonEmptyString(blockId)) {
    return fail("add_comment requires a non-empty string `blockId`.");
  }
  if (!isNonEmptyString(text)) {
    return fail("add_comment requires a non-empty string `text`.");
  }
  if (quote !== undefined && typeof quote !== "string") {
    return fail("add_comment's `quote` must be a string when provided.");
  }

  const comment: FolioAIComment = { text };
  const operation: FolioAIEditOperation = {
    id: "comment-1",
    type: "commentOnBlock",
    blockId,
    comment,
    ...(quote !== undefined ? { quote } : {}),
  };
  return ok(summarizeApplyResult(bridge.applyOperations([operation])));
};

const OPERATION_TYPES =
  "replaceInBlock, insertAfterBlock, insertBeforeBlock, replaceBlock, deleteBlock";

type SuggestedOperationType =
  | "replaceInBlock"
  | "insertAfterBlock"
  | "insertBeforeBlock"
  | "replaceBlock"
  | "deleteBlock";

const isOperationType = (value: unknown): value is SuggestedOperationType =>
  value === "replaceInBlock" ||
  value === "insertAfterBlock" ||
  value === "insertBeforeBlock" ||
  value === "replaceBlock" ||
  value === "deleteBlock";

/** Validate + map one `suggest_changes` operation, or return a plain-language error string. */
const buildSuggestedOperation = (raw: unknown, index: number): FolioAIEditOperation | string => {
  if (!isPlainObject(raw)) {
    return `operations[${index}] must be an object.`;
  }
  const type = raw["type"];
  const blockId = raw["blockId"];
  const id = raw["id"];
  const comment = raw["comment"];

  if (!isOperationType(type)) {
    return `operations[${index}].type must be one of ${OPERATION_TYPES}.`;
  }
  if (!isNonEmptyString(blockId)) {
    return `operations[${index}].blockId must be a non-empty string.`;
  }
  if (id !== undefined && !isNonEmptyString(id)) {
    return `operations[${index}].id must be a non-empty string when provided.`;
  }
  if (comment !== undefined && typeof comment !== "string") {
    return `operations[${index}].comment must be a string when provided.`;
  }
  const opId = isNonEmptyString(id) ? id : `op-${index + 1}`;
  const commentField = typeof comment === "string" ? { comment: { text: comment } } : {};

  if (type === "replaceInBlock") {
    const find = raw["find"];
    const replace = raw["replace"];
    if (!isNonEmptyString(find)) {
      return `operations[${index}] (replaceInBlock) requires a non-empty string \`find\`.`;
    }
    if (typeof replace !== "string") {
      return `operations[${index}] (replaceInBlock) requires a string \`replace\`.`;
    }
    return { id: opId, type, blockId, find, replace, ...commentField };
  }
  if (type === "insertAfterBlock" || type === "insertBeforeBlock") {
    const text = raw["text"];
    if (!isNonEmptyString(text)) {
      return `operations[${index}] (${type}) requires a non-empty string \`text\`.`;
    }
    return { id: opId, type, blockId, text, ...commentField };
  }
  if (type === "replaceBlock") {
    const text = raw["text"];
    if (!isNonEmptyString(text)) {
      return "operations[" + index + "] (replaceBlock) requires a non-empty string `text`.";
    }
    return { id: opId, type, blockId, text, ...commentField };
  }
  // deleteBlock
  return { id: opId, type, blockId, ...commentField };
};

const suggestChanges = (args: unknown, bridge: FolioAgentBridge): FolioToolCallResult => {
  if (!isPlainObject(args) || !Array.isArray(args["operations"])) {
    return fail("suggest_changes requires an `operations` array.");
  }
  const rawOperations = args["operations"];
  if (rawOperations.length === 0) {
    return fail("suggest_changes' `operations` array must not be empty.");
  }

  const operations: FolioAIEditOperation[] = [];
  for (const [index, raw] of rawOperations.entries()) {
    const built = buildSuggestedOperation(raw, index);
    if (typeof built === "string") {
      return fail(built);
    }
    operations.push(built);
  }

  return ok(summarizeApplyResult(bridge.applyOperations(operations)));
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
