import {
  createFolioAITextRangeHandle,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  getFolioDocumentOperationIssues,
  getFolioDocumentOutline,
  hashFolioAIBlockText,
  normalizeFolioAIBlockText,
  readFolioDocumentSection,
  type FolioAIBlock,
  type FolioAIEditNormalization,
  type FolioAITextRangeHandle,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatchPrecondition,
  type FolioDocumentStoryHandle,
} from "@stll/folio-core/server";

import type { FolioAgentBridge } from "./bridge";
import {
  decodeCommentId,
  decodeMainStoryTextRangeHandle,
  decodeSectionHandle,
  decodeStoryHandle,
} from "./codecs";
import { getDecodedCommentHandlers } from "./bridge";
import {
  explainTextTooLong,
  MAX_OPERATION_TEXT_LENGTH,
  parseAddCommentInput,
  prepareFolioAgentDocumentOperationBatch,
  parseSuggestChangesInput,
} from "./parse";
import type { FolioAgentToolOptions } from "./suggest-changes-options";
import type { FolioAgentToolInputByName, FolioToolCallResultFor } from "./tool-contract";
import { FOLIO_AGENT_TOOL_NAMES } from "./types";
import type {
  FolioAgentApplyOperationsSummary,
  FolioAgentCommentFilter,
  FolioAgentFindTextResult,
  FolioAgentInputNormalization,
  FolioAgentStoryFindTextResult,
  FolioAgentStoryTextMatch,
  FolioAgentToolName,
  FolioAgentTextMatch,
  FolioToolCallResult,
} from "./types";

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0;

const ok = <TResult>(result: TResult): FolioToolCallResult<TResult> => ({ ok: true, result });
const fail = (error: string): FolioToolCallResult<never> => ({ ok: false, error });

const VALID_TOOL_NAMES: readonly string[] = Object.values(FOLIO_AGENT_TOOL_NAMES);

/**
 * The same normalized-text hash the core snapshot/apply machinery uses for
 * `precondition.blockTextHash` (see `hashFolioAIBlockText` /
 * `normalizeFolioAIBlockText`). Computed straight from a block's `text`
 * rather than looked up from `snapshot.anchors` so every read path (main
 * document, a section, a find_text match) can attach it without threading
 * the anchors map around — it is defined to produce the exact same value.
 */
const blockTextHashOf = (text: string): string =>
  hashFolioAIBlockText(normalizeFolioAIBlockText(text));

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
export const executeFolioToolCallUntyped = (
  name: string,
  args: unknown,
  bridge: FolioAgentBridge,
  options: FolioAgentToolOptions = {},
): FolioToolCallResult => {
  if (!VALID_TOOL_NAMES.includes(name)) {
    return fail(`Unknown tool "${name}". Valid tools: ${VALID_TOOL_NAMES.join(", ")}.`);
  }

  try {
    return dispatch(name, args, bridge, options);
  } catch (error) {
    return fail(error instanceof Error ? error.message : String(error));
  }
};

export function executeFolioToolCall<TName extends FolioAgentToolName>(
  name: TName,
  args: FolioAgentToolInputByName[TName],
  bridge: FolioAgentBridge,
  options?: FolioAgentToolOptions,
): FolioToolCallResultFor<TName>;
export function executeFolioToolCall(
  name: string,
  args: unknown,
  bridge: FolioAgentBridge,
  options: FolioAgentToolOptions = {},
): FolioToolCallResult {
  return executeFolioToolCallUntyped(name, args, bridge, options);
}

const dispatch = (
  name: string,
  args: unknown,
  bridge: FolioAgentBridge,
  options: FolioAgentToolOptions,
): FolioToolCallResult => {
  if (!isFolioAgentToolName(name)) {
    return fail(`Unknown tool "${name}".`);
  }
  return TOOL_HANDLERS[name](args, bridge, options);
};

const readDocument = (
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readDocument> => {
  const { blocks } = bridge.snapshot();
  return ok(
    blocks.map((block) => ({
      blockId: block.id,
      kind: block.kind,
      text: block.text,
      blockTextHash: blockTextHashOf(block.text),
    })),
  );
};

const getDocumentOutline = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.getDocumentOutline> => {
  const maxDepth = isPlainObject(args) ? args["maxDepth"] : undefined;
  if (
    maxDepth !== undefined &&
    (typeof maxDepth !== "number" || !Number.isInteger(maxDepth) || maxDepth < 1 || maxDepth > 9)
  ) {
    return fail("get_document_outline's `maxDepth` must be an integer from 1 to 9.");
  }
  const outline = getFolioDocumentOutline(bridge.snapshot());
  const depth = maxDepth ?? 3;
  const sections = outline.sections
    .filter(({ level }) => level <= depth)
    .map((entry) => {
      const page =
        bridge.getTargetPage?.({
          type: "block",
          story: "main",
          blockId: entry.headingBlockId,
        }) ?? undefined;
      return page === undefined ? entry : Object.assign({}, entry, { page });
    });
  return ok({
    sections,
    totalSections: outline.sections.length,
    truncated: sections.length < outline.sections.length,
  });
};

const readSection = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readSection> => {
  if (!isPlainObject(args)) {
    return fail("read_section expects a section `handle` from get_document_outline.");
  }
  const handle = decodeSectionHandle(args["handle"]);
  if (handle === null) {
    return fail("read_section requires a valid `handle` from get_document_outline.");
  }
  const maxBlocks = args["maxBlocks"] ?? 100;
  if (
    typeof maxBlocks !== "number" ||
    !Number.isInteger(maxBlocks) ||
    maxBlocks < 1 ||
    maxBlocks > 200
  ) {
    return fail("read_section's `maxBlocks` must be an integer from 1 to 200.");
  }
  const afterBlockId = args["afterBlockId"];
  if (afterBlockId !== undefined && !isNonEmptyString(afterBlockId)) {
    return fail("read_section's `afterBlockId` must be a non-empty string when provided.");
  }

  const resolved = readFolioDocumentSection(bridge.snapshot(), handle);
  if (resolved.status === "missing") {
    return fail("The section no longer exists; run get_document_outline again.");
  }
  if (resolved.status === "stale") {
    return fail("The section heading changed; run get_document_outline again for a fresh handle.");
  }

  let startIndex = 0;
  if (afterBlockId !== undefined) {
    const cursorIndex = resolved.section.blocks.findIndex(({ id }) => id === afterBlockId);
    if (cursorIndex === -1) {
      return fail("read_section's `afterBlockId` is not part of this section.");
    }
    startIndex = cursorIndex + 1;
  }
  const selected = resolved.section.blocks.slice(startIndex, startIndex + maxBlocks);
  const hasMore = startIndex + selected.length < resolved.section.blocks.length;
  const lastBlockId = selected.at(-1)?.id;
  return ok({
    handle,
    heading: resolved.section.heading,
    blocks: selected.map(({ id, kind, text }) => ({
      blockId: id,
      kind,
      text,
      blockTextHash: blockTextHashOf(text),
    })),
    totalBlocks: resolved.section.blocks.length,
    truncated: hasMore,
    ...(hasMore && lastBlockId !== undefined && { nextAfterBlockId: lastBlockId }),
  });
};

const readStory = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readStory> => {
  if (!bridge.readStory) {
    return fail("This editor surface does not support story reads.");
  }
  const parsed = isPlainObject(args) ? decodeStoryHandle(args["handle"]) : null;
  if (parsed === null) {
    return fail("read_story requires a valid typed `handle` from list_stories.");
  }
  const story = bridge.readStory(parsed);
  return story ? ok(story) : fail("The requested story was not found; run list_stories again.");
};

const CONTEXT_RADIUS = 40;
const WORD_CHARACTER_AT_END = /[\p{L}\p{M}\p{N}_]$/u;
const WORD_CHARACTER_AT_START = /^[\p{L}\p{M}\p{N}_]/u;
/**
 * Window (UTF-16 code units) sliced on each side of a match for the
 * whole-word boundary check. The regexes above only ever test the single
 * grapheme adjacent to the match, but a surrogate pair or a base character
 * with stacked combining marks can span a few code units — this window is
 * generous enough to cover that while staying a constant, not
 * `block.text.length`. Without a bound, `.slice(0, at)` / `.slice(at + len)`
 * on every match makes a whole-word search O(n^2) in the block's length.
 */
const WORD_BOUNDARY_WINDOW = 8;

/** All matches for `query`, capped at {@link MAX_FIND_MATCHES}; `totalMatches` still counts every occurrence. */
type FindTextMatches<TMatch = FolioAgentTextMatch> = {
  matches: TMatch[];
  totalMatches: number;
};

const findTextMatches = (
  blocks: FolioAIBlock[],
  query: string,
  matchCase: boolean,
  wholeWord: boolean,
  getTargetPage?: (target: FolioAITextRangeHandle) => number | null,
  pageFilter?: number,
): FindTextMatches => {
  const matches: FolioAgentTextMatch[] = [];
  let totalMatches = 0;
  const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const expression = new RegExp(escapedQuery, matchCase ? "gu" : "giu");
  for (const block of blocks) {
    let occurrence = 0;
    let blockTextHash: string | undefined;
    for (const match of block.text.matchAll(expression)) {
      const at = match.index;
      const matchedText = match[0];
      if (
        wholeWord &&
        (WORD_CHARACTER_AT_END.test(block.text.slice(Math.max(0, at - WORD_BOUNDARY_WINDOW), at)) ||
          WORD_CHARACTER_AT_START.test(
            block.text.slice(
              at + matchedText.length,
              at + matchedText.length + WORD_BOUNDARY_WINDOW,
            ),
          ))
      ) {
        continue;
      }
      const range = createFolioAITextRangeHandle({
        blockId: block.id,
        text: block.text,
        startOffset: at,
        endOffset: at + matchedText.length,
      });
      if (range === null) {
        continue;
      }
      const page = getTargetPage?.(range) ?? undefined;
      if (pageFilter !== undefined && page !== pageFilter) {
        continue;
      }
      totalMatches += 1;
      if (matches.length < MAX_FIND_MATCHES) {
        const contextStart = Math.max(0, at - CONTEXT_RADIUS);
        const contextEnd = Math.min(block.text.length, at + matchedText.length + CONTEXT_RADIUS);
        blockTextHash ??= blockTextHashOf(block.text);
        matches.push({
          type: "main",
          story: { type: "main" },
          blockId: block.id,
          blockTextHash,
          range,
          occurrenceInBlock: occurrence,
          context: block.text.slice(contextStart, contextEnd),
          ...(page !== undefined && page !== null && { page }),
        });
      }
      occurrence += 1;
    }
  }
  return { matches, totalMatches };
};

const findStoryTextMatches = (
  text: string,
  story: Exclude<FolioDocumentStoryHandle, { type: "main" }>,
  query: string,
  matchCase: boolean,
  wholeWord: boolean,
): FindTextMatches<FolioAgentStoryTextMatch> => {
  const matches: FolioAgentStoryTextMatch[] = [];
  let totalMatches = 0;
  const escapedQuery = query.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
  const expression = new RegExp(escapedQuery, matchCase ? "gu" : "giu");
  for (const match of text.matchAll(expression)) {
    const at = match.index;
    const matchedText = match[0];
    if (
      wholeWord &&
      (WORD_CHARACTER_AT_END.test(text.slice(Math.max(0, at - WORD_BOUNDARY_WINDOW), at)) ||
        WORD_CHARACTER_AT_START.test(
          text.slice(at + matchedText.length, at + matchedText.length + WORD_BOUNDARY_WINDOW),
        ))
    ) {
      continue;
    }
    if (matches.length < MAX_FIND_MATCHES) {
      matches.push({
        type: "story",
        story,
        startOffset: at,
        endOffset: at + matchedText.length,
        occurrenceInStory: totalMatches,
        context: text.slice(
          Math.max(0, at - CONTEXT_RADIUS),
          Math.min(text.length, at + matchedText.length + CONTEXT_RADIUS),
        ),
      });
    }
    totalMatches += 1;
  }
  return { matches, totalMatches };
};

const findText = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.findText> => {
  if (!isPlainObject(args)) {
    return fail("find_text expects an object with a non-empty `query` string.");
  }
  const query = args["query"];
  const matchCase = args["matchCase"];
  const wholeWord = args["wholeWord"];
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
  if (wholeWord !== undefined && typeof wholeWord !== "boolean") {
    return fail("find_text's `wholeWord` must be a boolean when provided.");
  }

  const scope = args["scope"];
  let blocks = bridge.snapshot().blocks;
  let getTargetPage: ((target: FolioAITextRangeHandle) => number | null) | undefined;
  let pageFilter: number | undefined;
  if (scope !== undefined) {
    if (!isPlainObject(scope) || !isNonEmptyString(scope["type"])) {
      return fail("find_text's `scope` must be a document, section, page, or story scope.");
    }
    if (scope["type"] === "section") {
      const handle = decodeSectionHandle(scope["handle"]);
      if (handle === null) {
        return fail("find_text's section scope requires a handle from get_document_outline.");
      }
      const section = readFolioDocumentSection(bridge.snapshot(), handle);
      if (section.status !== "found") {
        return fail("The scoped section is stale or missing; run get_document_outline again.");
      }
      blocks = section.section.blocks;
    } else if (scope["type"] === "page") {
      const page = scope["page"];
      if (typeof page !== "number" || !Number.isInteger(page) || page < 1) {
        return fail("find_text's page scope requires an integer `page` >= 1.");
      }
      if (!bridge.getTargetPage) {
        return fail("This editor surface cannot search by page without live pagination.");
      }
      const pageCount = bridge.getPageCount?.();
      if (pageCount !== undefined && page > pageCount) {
        return fail(`find_text's page (${page}) exceeds the document's page count (${pageCount}).`);
      }
      getTargetPage = bridge.getTargetPage;
      pageFilter = page;
    } else if (scope["type"] === "story") {
      const handle = decodeStoryHandle(scope["handle"]);
      if (handle === null) {
        return fail("find_text's story scope requires a handle from list_stories.");
      }
      if (handle.type !== "main") {
        if (!bridge.readStory) {
          return fail("This editor surface does not support story search.");
        }
        const story = bridge.readStory(handle);
        if (story === null) {
          return fail("The scoped story was not found; run list_stories again.");
        }
        const found = findStoryTextMatches(
          story.text,
          handle,
          query,
          matchCase === true,
          wholeWord === true,
        );
        return ok({
          matches: found.matches,
          truncated: found.totalMatches > found.matches.length,
          totalMatches: found.totalMatches,
        } satisfies FolioAgentStoryFindTextResult);
      }
    } else if (scope["type"] !== "document") {
      return fail("find_text's `scope.type` must be document, section, page, or story.");
    }
  }
  const { matches, totalMatches } = findTextMatches(
    blocks,
    query,
    matchCase === true,
    wholeWord === true,
    getTargetPage,
    pageFilter,
  );
  const result: FolioAgentFindTextResult = {
    matches,
    truncated: totalMatches > matches.length,
    totalMatches,
  };
  return ok(result);
};

const isCommentFilter = (value: unknown): value is FolioAgentCommentFilter =>
  value === "all" || value === "open" || value === "resolved";

const readComments = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readComments> => {
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
  if (reason === "atomicBatchRejected") {
    return "another operation in this atomic batch could not be applied; no operations were committed.";
  }
  if (reason === "preconditionFailed") {
    return "the target block changed after the operation was prepared; re-read the document and retry with a fresh operation.";
  }
  if (reason === "staleRange") {
    return "the selected range changed or shifted; run find_text again and retry with the fresh range.";
  }
  if (reason === "emptyOperation") {
    return "this operation has no effect; nothing to apply.";
  }
  if (reason === "noopOperation") {
    return "this operation would not change the document (the text already matches what you asked for).";
  }
  if (reason === "documentVersionMismatch") {
    return "the document changed after these edits were proposed; re-read it and regenerate the edits against the current version.";
  }
  if (reason === "documentNotEditable") {
    return "the document is not open for editing right now; nothing was applied or queued. Ask the user to open it for editing, then retry.";
  }
  return reason;
};

/**
 * Turn one apply-time `FolioAIEditNormalization` (an automatic adjustment the
 * applier made to an operation's input, e.g. splitting a multi-line
 * `insertAfterBlock` / `insertBeforeBlock` `text` into several paragraphs)
 * into the same `{ path, message }` shape lenient decoding reports, so the
 * model sees every normalization in one list regardless of where it happened.
 */
const explainApplyNormalization = (
  normalization: FolioAIEditNormalization,
): FolioAgentInputNormalization => {
  if (normalization.code === "splitMultilineText") {
    return {
      path: `operations[id=${normalization.id}].text`,
      message:
        `\`text\` contained line breaks; split into ${normalization.paragraphCount} ` +
        "consecutive paragraphs (one per non-blank line) instead of one paragraph with " +
        "embedded newlines.",
    };
  }
  return {
    path: `operations[id=${normalization.id}]`,
    message: `input was normalized (${normalization.code}).`,
  };
};

type ApplyResultLike = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: { id: string }[];
  queued?: { id: string }[];
  skipped: { id: string; reason: string }[];
  issues?: FolioAgentApplyOperationsSummary["issues"];
  receipts?: FolioAgentApplyOperationsSummary["receipts"];
  normalizations?: readonly FolioAIEditNormalization[];
};

const summarizeApplyResult = (
  result: ApplyResultLike,
  normalizations: readonly FolioAgentInputNormalization[],
): FolioAgentApplyOperationsSummary => ({
  version: result.version,
  applied: result.applied.map((entry) => ({ id: entry.id })),
  queued: (result.queued ?? []).map((entry) => ({ id: entry.id })),
  skipped: result.skipped.map((entry) => ({
    id: entry.id,
    reason: explainSkipReason(entry.reason),
  })),
  issues: result.issues ?? [],
  receipts: result.receipts ?? [],
  normalizations: [
    ...normalizations,
    ...(result.normalizations ?? []).map(explainApplyNormalization),
  ],
});

/** Every operation skipped for one batch-wide reason, before anything reached the bridge. */
const rejectBatch = (
  operations: readonly FolioDocumentOperation[],
  reason: "documentVersionMismatch",
  normalizations: readonly FolioAgentInputNormalization[],
): FolioAgentApplyOperationsSummary => {
  const skipped = operations.map(({ id }) => ({ id, reason }));
  return summarizeApplyResult(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      applied: [],
      skipped,
      issues: getFolioDocumentOperationIssues(operations, skipped),
      receipts: [],
    },
    normalizations,
  );
};

/**
 * Attach a `precondition.blockTextHash` guard to every operation before
 * handing the batch to the bridge.
 *
 * When the caller (the model) already echoed a `precondition` on the
 * operation — sourced from a `blockTextHash` returned by an earlier
 * `read_document` / `read_section` / `find_text` call — that is honored
 * as-is: it is the only signal that can catch a document edit made
 * BETWEEN that read and this apply call.
 *
 * When no precondition was supplied, fall back to stamping one from this
 * call's own fresh `bridge.snapshot()`, matching the previous behavior.
 * This fallback cannot detect cross-call staleness — the snapshot it reads
 * from is taken right before applying, so it always matches the current
 * live document — but it still guards a later operation in the SAME batch
 * against an earlier operation in that batch touching the same block.
 */
type ApplyOperationsOptions = {
  operations: readonly FolioDocumentOperation[];
  precondition?: FolioDocumentOperationBatchPrecondition;
  normalizations?: readonly FolioAgentInputNormalization[];
};

const applyOperations = (
  bridge: FolioAgentBridge,
  { operations, precondition, normalizations = [] }: ApplyOperationsOptions,
): FolioAgentApplyOperationsSummary => {
  const snapshot = bridge.snapshot();
  const guardedOperations: FolioDocumentOperation[] = [];
  for (const operation of operations) {
    if (operation.precondition !== undefined) {
      guardedOperations.push(operation);
      continue;
    }
    const blockId =
      operation.type === "replaceRange" ||
      operation.type === "commentOnRange" ||
      operation.type === "formatRange"
        ? operation.range.blockId
        : operation.blockId;
    const blockTextHash = snapshot.anchors[blockId]?.textHash;
    guardedOperations.push({
      ...operation,
      ...(blockTextHash !== undefined && { precondition: { blockTextHash } }),
    });
  }
  return summarizeApplyResult(
    bridge.applyDocumentOperations(
      prepareFolioAgentDocumentOperationBatch({
        operations: guardedOperations,
        ...(bridge.documentOperationMode !== undefined && {
          mode: bridge.documentOperationMode,
        }),
        ...(precondition !== undefined && { precondition }),
      }),
    ),
    normalizations,
  );
};

const addComment = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.addComment> => {
  const parsed = parseAddCommentInput(args);
  if (!parsed.ok) {
    return fail(parsed.error);
  }
  return ok(applyOperations(bridge, { operations: [parsed.operation] }));
};

/**
 * A version-pinned batch is compared against the bridge's document version
 * BEFORE any operation is applied or queued: an approval that runs after
 * the document moved on skips as a whole instead of landing half of its
 * edits on a document the model never saw. A surface without a version
 * notion cannot honour the pin and refuses rather than guessing.
 */
const suggestChanges = (
  args: unknown,
  bridge: FolioAgentBridge,
  options: FolioAgentToolOptions,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.suggestChanges> => {
  const parsed = parseSuggestChangesInput(args, options.suggestChanges);
  if (!parsed.ok) {
    return fail(parsed.error);
  }
  if (parsed.precondition !== undefined) {
    if (!bridge.getDocumentVersion) {
      return fail(
        "This editor surface cannot verify `documentVersion`; it has no document version to compare against.",
      );
    }
    if (bridge.getDocumentVersion() !== parsed.precondition.documentVersion) {
      return ok(rejectBatch(parsed.operations, "documentVersionMismatch", parsed.normalizations));
    }
  }
  return ok(
    applyOperations(bridge, {
      operations: parsed.operations,
      ...(parsed.precondition !== undefined && { precondition: parsed.precondition }),
      normalizations: parsed.normalizations,
    }),
  );
};

const replyComment = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.replyComment> => {
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
  const decodedHandlers = getDecodedCommentHandlers(bridge);
  if (decodedHandlers === undefined) {
    const replied = bridge.replyToComment(commentId, text);
    return replied ? ok({ replied: true }) : fail(`No comment with id "${commentId}" was found.`);
  }
  const decodedCommentId = decodeCommentId(commentId);
  if (decodedCommentId === null) {
    return fail("reply_comment's `commentId` must be a comment id from `read_comments`.");
  }
  const replied = decodedHandlers.replyToComment(decodedCommentId, text);
  if (!replied) {
    return fail(`No comment with id "${commentId}" was found.`);
  }
  return ok({ replied: true });
};

const resolveComment = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.resolveComment> => {
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
  const decodedHandlers = getDecodedCommentHandlers(bridge);
  if (decodedHandlers === undefined) {
    const changed = bridge.resolveComment(commentId, resolved);
    return changed ? ok({ resolved }) : fail(`No comment with id "${commentId}" was found.`);
  }
  const decodedCommentId = decodeCommentId(commentId);
  if (decodedCommentId === null) {
    return fail("resolve_comment's `commentId` must be a comment id from `read_comments`.");
  }
  const changed = decodedHandlers.resolveComment(decodedCommentId, resolved);
  if (!changed) {
    return fail(`No comment with id "${commentId}" was found.`);
  }
  return ok({ resolved });
};

const readPage = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readPage> => {
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
  return ok({
    page,
    ...(pageCount !== undefined && { totalPages: pageCount }),
    text: bridge.getPageText(page),
  });
};

const readSelection = (
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.readSelection> => {
  if (!bridge.getSelectionText) {
    return fail("This editor surface does not support read_selection (no live selection).");
  }
  return ok({ text: bridge.getSelectionText() });
};

const scrollToBlock = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.scrollToBlock> => {
  if (!bridge.scrollToBlock) {
    return fail("This editor surface does not support scroll_to_block (no live editor view).");
  }
  const blockId = isPlainObject(args) ? args["blockId"] : undefined;
  if (!isNonEmptyString(blockId)) {
    return fail("scroll_to_block requires a non-empty string `blockId`.");
  }
  return ok({ scrolled: bridge.scrollToBlock(blockId) });
};

const showInDocument = (
  args: unknown,
  bridge: FolioAgentBridge,
): FolioToolCallResultFor<typeof FOLIO_AGENT_TOOL_NAMES.showInDocument> => {
  if (!bridge.showInDocument) {
    return fail("This editor surface does not support show_in_document (no live editor view).");
  }
  if (!isPlainObject(args)) {
    return fail("show_in_document expects exactly one of `blockId` or `range`.");
  }
  const blockId = args["blockId"];
  const rawRange = args["range"];
  if ((blockId === undefined) === (rawRange === undefined)) {
    return fail("show_in_document requires exactly one of `blockId` or `range`.");
  }
  if (blockId !== undefined) {
    if (!isNonEmptyString(blockId)) {
      return fail("show_in_document's `blockId` must be a non-empty string.");
    }
    return ok({
      shown: bridge.showInDocument({ type: "block", story: "main", blockId }),
    });
  }
  const range = decodeMainStoryTextRangeHandle(rawRange);
  if (range === null) {
    return fail("show_in_document's `range` must be copied from find_text.");
  }
  return ok({ shown: bridge.showInDocument(range) });
};

type FolioAgentToolHandlers = {
  [Name in FolioAgentToolName]: (
    args: unknown,
    bridge: FolioAgentBridge,
    options: FolioAgentToolOptions,
  ) => FolioToolCallResultFor<Name>;
};

const TOOL_HANDLERS = {
  [FOLIO_AGENT_TOOL_NAMES.readDocument]: (_args, bridge) => readDocument(bridge),
  [FOLIO_AGENT_TOOL_NAMES.getDocumentOutline]: getDocumentOutline,
  [FOLIO_AGENT_TOOL_NAMES.readSection]: readSection,
  [FOLIO_AGENT_TOOL_NAMES.listStories]: (_args, bridge) =>
    bridge.listStories
      ? ok(bridge.listStories())
      : fail("This editor surface does not support story discovery."),
  [FOLIO_AGENT_TOOL_NAMES.readStory]: readStory,
  [FOLIO_AGENT_TOOL_NAMES.findText]: findText,
  [FOLIO_AGENT_TOOL_NAMES.readComments]: readComments,
  [FOLIO_AGENT_TOOL_NAMES.readChanges]: (_args, bridge) => ok(bridge.getChanges()),
  [FOLIO_AGENT_TOOL_NAMES.addComment]: addComment,
  [FOLIO_AGENT_TOOL_NAMES.suggestChanges]: suggestChanges,
  [FOLIO_AGENT_TOOL_NAMES.replyComment]: replyComment,
  [FOLIO_AGENT_TOOL_NAMES.resolveComment]: resolveComment,
  [FOLIO_AGENT_TOOL_NAMES.readPage]: readPage,
  [FOLIO_AGENT_TOOL_NAMES.readSelection]: (_args, bridge) => readSelection(bridge),
  [FOLIO_AGENT_TOOL_NAMES.scrollToBlock]: scrollToBlock,
  [FOLIO_AGENT_TOOL_NAMES.showInDocument]: showInDocument,
} satisfies FolioAgentToolHandlers;

const isFolioAgentToolName = (name: string): name is FolioAgentToolName =>
  Object.hasOwn(TOOL_HANDLERS, name);
