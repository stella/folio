/**
 * Validation-only parsers for `suggest_changes` / `add_comment` tool-call
 * arguments — the canonical rules `execute.ts` runs before handing operations
 * to a {@link FolioAgentBridge}, factored out so a host with its own
 * review-queue UX (its own "propose an edit" path instead of
 * `executeFolioToolCall`) can validate a model's tool-call arguments with the
 * exact same rules, then route the resulting `FolioAIEditOperation`(s)
 * through that path instead of applying them immediately.
 *
 * `execute.ts` delegates to these two functions rather than duplicating the
 * rules, so there is exactly one place these caps and error messages live.
 */

import type {
  FolioAIComment,
  FolioAIEditOperation,
  FolioAITextRangeHandle,
} from "@stll/folio-core/server";

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isNonEmptyString = (value: unknown): value is string =>
  typeof value === "string" && value.length > 0;

/**
 * Hard caps on `suggest_changes` / `add_comment` / `reply_comment` input
 * size. Without these, a single tool call could ask the bridge to apply an
 * unbounded number of operations, or push an arbitrarily large string into
 * the tracked-changes engine, in one shot. `execute.ts` reuses
 * {@link MAX_OPERATION_TEXT_LENGTH} for `reply_comment`'s text cap too, so
 * the limit stays a single number shared across every text-bearing tool.
 */
export const MAX_OPERATIONS_PER_CALL = 50;
export const MAX_OPERATION_TEXT_LENGTH = 100_000;
const MAX_TABLE_INSERTION_CELL_TEXTS = 256;

/** Plain-language error for a string argument over {@link MAX_OPERATION_TEXT_LENGTH}. */
export const explainTextTooLong = (label: string, length: number): string =>
  `${label} is ${length.toLocaleString()} characters, over the ${MAX_OPERATION_TEXT_LENGTH.toLocaleString()}-character limit; shorten it or split it into multiple operations.`;

/** Result of {@link parseAddCommentInput}. */
export type ParseAddCommentResult =
  | { ok: true; operation: FolioAIEditOperation }
  | { ok: false; error: string };

/**
 * Validate `add_comment`'s raw tool-call arguments and build the
 * `commentOnBlock` {@link FolioAIEditOperation} it applies. Pure: does not
 * touch a bridge or document.
 */
export const parseAddCommentInput = (args: unknown): ParseAddCommentResult => {
  if (!isPlainObject(args)) {
    return { ok: false, error: "add_comment expects an object with `blockId` and `text` strings." };
  }
  const blockId = args["blockId"];
  const quote = args["quote"];
  const text = args["text"];
  if (!isNonEmptyString(blockId)) {
    return { ok: false, error: "add_comment requires a non-empty string `blockId`." };
  }
  if (!isNonEmptyString(text)) {
    return { ok: false, error: "add_comment requires a non-empty string `text`." };
  }
  if (text.length > MAX_OPERATION_TEXT_LENGTH) {
    return { ok: false, error: explainTextTooLong("add_comment's `text`", text.length) };
  }
  if (quote !== undefined && typeof quote !== "string") {
    return { ok: false, error: "add_comment's `quote` must be a string when provided." };
  }
  if (typeof quote === "string" && quote.length > MAX_OPERATION_TEXT_LENGTH) {
    return { ok: false, error: explainTextTooLong("add_comment's `quote`", quote.length) };
  }

  const comment: FolioAIComment = { text };
  const operation: FolioAIEditOperation = {
    id: "comment-1",
    type: "commentOnBlock",
    blockId,
    comment,
    ...(quote !== undefined ? { quote } : {}),
  };
  return { ok: true, operation };
};

const OPERATION_TYPES =
  "replaceInBlock, replaceRange, commentOnRange, formatRange, insertAfterBlock, insertBeforeBlock, replaceBlock, deleteBlock, insertTableRow, deleteTableRow, insertTableColumn, deleteTableColumn";

type SuggestedOperationType =
  | "replaceInBlock"
  | "replaceRange"
  | "commentOnRange"
  | "formatRange"
  | "insertAfterBlock"
  | "insertBeforeBlock"
  | "replaceBlock"
  | "deleteBlock"
  | "insertTableRow"
  | "deleteTableRow"
  | "insertTableColumn"
  | "deleteTableColumn";

const isOperationType = (value: unknown): value is SuggestedOperationType =>
  value === "replaceInBlock" ||
  value === "replaceRange" ||
  value === "commentOnRange" ||
  value === "formatRange" ||
  value === "insertAfterBlock" ||
  value === "insertBeforeBlock" ||
  value === "replaceBlock" ||
  value === "deleteBlock" ||
  value === "insertTableRow" ||
  value === "deleteTableRow" ||
  value === "insertTableColumn" ||
  value === "deleteTableColumn";

const readTextRange = (value: unknown, index: number): FolioAITextRangeHandle | string => {
  const path = `operations[${index}].range`;
  if (!isPlainObject(value)) {
    return `${path} must be an object copied from find_text.`;
  }
  const type = value["type"];
  const story = value["story"];
  const blockId = value["blockId"];
  const startOffset = value["startOffset"];
  const endOffset = value["endOffset"];
  const selectedTextHash = value["selectedTextHash"];
  if (type !== "textRange" || story !== "main") {
    return `${path} must be a main-story textRange copied from find_text.`;
  }
  if (!isNonEmptyString(blockId)) {
    return `${path}.blockId must be a non-empty string.`;
  }
  if (typeof startOffset !== "number" || !Number.isInteger(startOffset) || startOffset < 0) {
    return `${path}.startOffset must be a non-negative integer.`;
  }
  if (typeof endOffset !== "number" || !Number.isInteger(endOffset) || endOffset <= startOffset) {
    return `${path}.endOffset must be an integer greater than startOffset.`;
  }
  if (!isNonEmptyString(selectedTextHash) || !/^h[0-9a-z]+$/.test(selectedTextHash)) {
    return `${path}.selectedTextHash must be a normalized text hash.`;
  }
  return { type, story, blockId, startOffset, endOffset, selectedTextHash };
};

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
  if (id !== undefined && !isNonEmptyString(id)) {
    return `operations[${index}].id must be a non-empty string when provided.`;
  }
  if (comment !== undefined && typeof comment !== "string") {
    return `operations[${index}].comment must be a string when provided.`;
  }
  if (typeof comment === "string" && comment.length > MAX_OPERATION_TEXT_LENGTH) {
    return explainTextTooLong(`operations[${index}].comment`, comment.length);
  }
  const opId = isNonEmptyString(id) ? id : `op-${index + 1}`;
  const commentField = typeof comment === "string" ? { comment: { text: comment } } : {};

  if (type === "replaceRange" || type === "commentOnRange" || type === "formatRange") {
    const range = readTextRange(raw["range"], index);
    if (typeof range === "string") {
      return range;
    }
    if (type === "commentOnRange") {
      if (!isNonEmptyString(comment)) {
        return `operations[${index}] (commentOnRange) requires a non-empty string \`comment\`.`;
      }
      return { id: opId, type, range, comment: { text: comment } };
    }
    if (type === "formatRange") {
      const rawFormatting = raw["formatting"];
      if (!isPlainObject(rawFormatting)) {
        return `operations[${index}] (formatRange) requires a \`formatting\` object.`;
      }
      const formatting: { bold?: boolean; italic?: boolean; underline?: boolean } = {};
      for (const key of ["bold", "italic", "underline"] as const) {
        const value = rawFormatting[key];
        if (value === undefined) {
          continue;
        }
        if (typeof value !== "boolean") {
          return `operations[${index}].formatting.${key} must be a boolean when provided.`;
        }
        formatting[key] = value;
      }
      if (Object.keys(formatting).length === 0) {
        return `operations[${index}] (formatRange) requires at least one formatting property.`;
      }
      return { id: opId, type, range, formatting };
    }
    const replace = raw["replace"];
    if (typeof replace !== "string") {
      return `operations[${index}] (replaceRange) requires a string \`replace\`.`;
    }
    if (replace.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(`operations[${index}] (replaceRange) \`replace\``, replace.length);
    }
    return { id: opId, type, range, replace, ...commentField };
  }

  if (!isNonEmptyString(blockId)) {
    return `operations[${index}].blockId must be a non-empty string.`;
  }

  if (type === "insertTableRow" || type === "insertTableColumn") {
    const position = raw["position"];
    if (position !== undefined && position !== "after" && position !== "before") {
      return `operations[${index}] (${type}) \`position\` must be "after" or "before" when provided.`;
    }
    const rawCellTexts = raw["cellTexts"];
    if (rawCellTexts !== undefined && !Array.isArray(rawCellTexts)) {
      return `operations[${index}] (${type}) \`cellTexts\` must be an array of strings when provided.`;
    }
    if (rawCellTexts !== undefined && rawCellTexts.length > MAX_TABLE_INSERTION_CELL_TEXTS) {
      return `operations[${index}] (${type}) \`cellTexts\` has ${rawCellTexts.length.toLocaleString()} entries, over the ${MAX_TABLE_INSERTION_CELL_TEXTS.toLocaleString()}-cell limit.`;
    }
    const cellTexts: string[] = [];
    for (const [cellIndex, cellText] of (rawCellTexts ?? []).entries()) {
      if (typeof cellText !== "string") {
        return `operations[${index}] (${type}) \`cellTexts[${cellIndex}]\` must be a string.`;
      }
      if (cellText.length > MAX_OPERATION_TEXT_LENGTH) {
        return explainTextTooLong(
          `operations[${index}] (${type}) \`cellTexts[${cellIndex}]\``,
          cellText.length,
        );
      }
      cellTexts.push(cellText);
    }
    return {
      id: opId,
      type,
      blockId,
      ...(position !== undefined && { position }),
      ...(rawCellTexts !== undefined && { cellTexts }),
    };
  }

  if (type === "deleteTableRow" || type === "deleteTableColumn") {
    return { id: opId, type, blockId };
  }

  if (type === "replaceInBlock") {
    const find = raw["find"];
    const replace = raw["replace"];
    if (!isNonEmptyString(find)) {
      return `operations[${index}] (replaceInBlock) requires a non-empty string \`find\`.`;
    }
    if (find.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(`operations[${index}] (replaceInBlock) \`find\``, find.length);
    }
    if (typeof replace !== "string") {
      return `operations[${index}] (replaceInBlock) requires a string \`replace\`.`;
    }
    if (replace.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(
        `operations[${index}] (replaceInBlock) \`replace\``,
        replace.length,
      );
    }
    return { id: opId, type, blockId, find, replace, ...commentField };
  }
  if (type === "insertAfterBlock" || type === "insertBeforeBlock") {
    const text = raw["text"];
    if (!isNonEmptyString(text)) {
      return `operations[${index}] (${type}) requires a non-empty string \`text\`.`;
    }
    if (text.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(`operations[${index}] (${type}) \`text\``, text.length);
    }
    return { id: opId, type, blockId, text, ...commentField };
  }
  if (type === "replaceBlock") {
    const text = raw["text"];
    if (!isNonEmptyString(text)) {
      return "operations[" + index + "] (replaceBlock) requires a non-empty string `text`.";
    }
    if (text.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(`operations[${index}] (replaceBlock) \`text\``, text.length);
    }
    return { id: opId, type, blockId, text, ...commentField };
  }
  // deleteBlock
  return { id: opId, type, blockId, ...commentField };
};

/** Result of {@link parseSuggestChangesInput}. */
export type ParseSuggestChangesResult =
  | { ok: true; operations: FolioAIEditOperation[] }
  | { ok: false; error: string };

/**
 * Validate `suggest_changes`' raw tool-call arguments and build the
 * {@link FolioAIEditOperation}s it applies. Pure: does not touch a bridge or
 * document.
 */
export const parseSuggestChangesInput = (args: unknown): ParseSuggestChangesResult => {
  if (!isPlainObject(args) || !Array.isArray(args["operations"])) {
    return { ok: false, error: "suggest_changes requires an `operations` array." };
  }
  const rawOperations = args["operations"];
  if (rawOperations.length === 0) {
    return { ok: false, error: "suggest_changes' `operations` array must not be empty." };
  }
  if (rawOperations.length > MAX_OPERATIONS_PER_CALL) {
    return {
      ok: false,
      error: `suggest_changes' \`operations\` array has ${rawOperations.length.toLocaleString()} entries, over the ${MAX_OPERATIONS_PER_CALL}-operation limit; batch it across multiple suggest_changes calls.`,
    };
  }

  const operations: FolioAIEditOperation[] = [];
  for (const [index, raw] of rawOperations.entries()) {
    const built = buildSuggestedOperation(raw, index);
    if (typeof built === "string") {
      return { ok: false, error: built };
    }
    operations.push(built);
  }

  return { ok: true, operations };
};
