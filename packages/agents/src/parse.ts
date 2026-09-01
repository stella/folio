/**
 * Validation-only parsers for `suggest_changes` / `add_comment` tool-call
 * arguments — the canonical rules `execute.ts` runs before handing operations
 * to a {@link FolioAgentBridge}, factored out so a host can validate a
 * model's tool-call arguments with the exact same rules (a host review queue
 * should prefer a queue bridge, see the README, but the parsers stay pure).
 *
 * `parseSuggestChangesInput` is a front door over the contract parser in
 * `@stll/folio-core`, not a second parser: it decodes leniently (reporting
 * every normalisation), enforces the agent-layer caps and the host's
 * {@link FolioSuggestChangesOptions}, mints ids, wraps `comment` strings, and
 * then delegates every per-operation rule to `parseFolioDocumentOperationBatch`.
 */

import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE,
  InvalidFolioDocumentOperationBatchError,
  parseFolioDocumentOperationBatch,
  type FolioAIComment,
  type FolioAIEditOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationBatchPrecondition,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

import {
  resolveSuggestChangesOptions,
  type FolioSuggestChangesOptions,
  type ResolvedFolioSuggestChangesOptions,
} from "./suggest-changes-options";
import type { FolioAgentInputNormalization } from "./types";

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
 * The per-call operation count is a host option
 * (`FolioSuggestChangesOptions.maxOperations`), defaulting to
 * {@link DEFAULT_MAX_OPERATIONS_PER_CALL}.
 */
export const MAX_OPERATION_TEXT_LENGTH = 100_000;
const MAX_TABLE_INSERTION_CELL_TEXTS = 256;

/**
 * Aggregate cap on the SUM of every text-bearing field's length across one
 * `suggest_changes` call. Each field alone is bounded by
 * {@link MAX_OPERATION_TEXT_LENGTH} and each call by the operation cap, but
 * an operation can carry several capped fields (e.g. `comment` plus `text`),
 * and a table-insertion operation can carry up to
 * {@link MAX_TABLE_INSERTION_CELL_TEXTS} capped cell texts — so a
 * maximally-shaped batch could still push an unbounded total into the
 * tracked-changes engine in one call even though every per-field cap was
 * respected. This budget bounds the running total instead.
 */
export const MAX_TOTAL_OPERATION_TEXT_LENGTH = 2_000_000;

/** Plain-language error for a string argument over {@link MAX_OPERATION_TEXT_LENGTH}. */
export const explainTextTooLong = (label: string, length: number): string =>
  `${label} is ${length.toLocaleString()} characters, over the ${MAX_OPERATION_TEXT_LENGTH.toLocaleString()}-character limit; shorten it or split it into multiple operations.`;

/** Plain-language error when the running total across all operations exceeds {@link MAX_TOTAL_OPERATION_TEXT_LENGTH}. */
const explainAggregateTextTooLong = (index: number): string =>
  `operations[${index}] pushes suggest_changes' combined text over the ${MAX_TOTAL_OPERATION_TEXT_LENGTH.toLocaleString()}-character aggregate limit across all operations; split the edit across multiple suggest_changes calls.`;

/** Normalized text hashes (`hashFolioAIBlockText`) look like `h` followed by base-36 digits. */
const NORMALIZED_TEXT_HASH_PATTERN = /^h[0-9a-z]+$/;

/**
 * Parse an optional `precondition: { blockTextHash }` field on `add_comment`
 * arguments, present when the caller echoes a `blockTextHash` returned by an
 * earlier `read_document` / `read_section` / `find_text` call. Returns
 * `undefined` when omitted, the parsed precondition when valid, or a
 * plain-language error string otherwise.
 */
const readOperationPrecondition = (
  value: unknown,
): { blockTextHash: string } | undefined | string => {
  if (value === undefined) {
    return undefined;
  }
  if (!isPlainObject(value)) {
    return "`precondition` must be an object when provided.";
  }
  const blockTextHash = value["blockTextHash"];
  if (!isNonEmptyString(blockTextHash) || !NORMALIZED_TEXT_HASH_PATTERN.test(blockTextHash)) {
    return "`precondition.blockTextHash` must be a normalized block text hash, echoed from read_document / read_section / find_text.";
  }
  return { blockTextHash };
};

/** Mutable running budget threaded through the text caps for one `suggest_changes` call. */
type TextBudget = { remaining: number };

/** Decrement the shared budget by `length`; returns true once the aggregate cap is exceeded. */
const consumeTextBudget = (budget: TextBudget, length: number): boolean => {
  budget.remaining -= length;
  return budget.remaining < 0;
};

export type PrepareFolioAgentDocumentOperationBatchOptions = {
  operations: readonly FolioAIEditOperation[];
  mode?: FolioDocumentOperationMode;
  atomic?: boolean;
  dryRun?: boolean;
  precondition?: FolioDocumentOperationBatchPrecondition;
};

/**
 * Wrap agent-preprocessed operations in the versioned batch envelope and
 * delegate the canonical contract validation to core's parser. The returned
 * batch is safe to hand straight to `applyDocumentOperations`; core marks
 * parsed batches internally so the downstream apply path can skip reparsing.
 */
export const prepareFolioAgentDocumentOperationBatch = ({
  operations,
  mode,
  atomic,
  dryRun,
  precondition,
}: PrepareFolioAgentDocumentOperationBatchOptions): FolioDocumentOperationBatch =>
  parseFolioDocumentOperationBatch({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    operations,
    ...(mode !== undefined && { mode }),
    ...(atomic !== undefined && { atomic }),
    ...(dryRun !== undefined && { dryRun }),
    ...(precondition !== undefined && { precondition }),
  });

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
  const precondition = readOperationPrecondition(args["precondition"]);
  if (typeof precondition === "string") {
    return { ok: false, error: `add_comment's ${precondition}` };
  }

  const comment: FolioAIComment = { text };
  const operation: FolioAIEditOperation = {
    id: "comment-1",
    type: "commentOnBlock",
    blockId,
    comment,
    ...(quote !== undefined ? { quote } : {}),
    ...(precondition !== undefined ? { precondition } : {}),
  };
  return { ok: true, operation };
};

// ---------------------------------------------------------------------------
// suggest_changes
// ---------------------------------------------------------------------------

/** Result of {@link parseSuggestChangesInput}. */
export type ParseSuggestChangesResult =
  | {
      ok: true;
      operations: FolioAIEditOperation[];
      /** Present when the call pinned a host document version. */
      precondition?: FolioDocumentOperationBatchPrecondition;
      normalizations: FolioAgentInputNormalization[];
    }
  | { ok: false; error: string };

const SUGGEST_CHANGES_ARGUMENT_KEYS = ["operations", "documentVersion"] as const;

/** String-valued operation fields that count against the text caps. */
const TEXT_FIELDS = ["find", "replace", "text", "quote", "styleId", "area"] as const;

/**
 * Try to read a value the model serialised as a JSON string where an object
 * or array was expected. Returns `undefined` when it is not such a string.
 */
const parseEmbeddedJson = (value: unknown): unknown => {
  if (typeof value !== "string") {
    return undefined;
  }
  const trimmed = value.trim();
  if (!trimmed.startsWith("{") && !trimmed.startsWith("[")) {
    return undefined;
  }
  // Boundary decode of model output: JSON.parse throws on malformed text,
  // and "not JSON after all" is a normal outcome here, not a failure.
  try {
    return JSON.parse(trimmed);
  } catch {
    return undefined;
  }
};

type DecodedSuggestChangesInput = {
  operations: Record<string, unknown>[];
  documentVersion: unknown;
  normalizations: FolioAgentInputNormalization[];
};

/**
 * Lenient decode of the raw arguments: JSON-string envelopes and operations,
 * `kind` for `type`, stray keys. Every normalisation is recorded so the host
 * and the model can see what was tolerated. Shape errors that cannot be
 * normalised come back as a plain-language string.
 */
const decodeSuggestChangesInput = (
  rawArgs: unknown,
  options: ResolvedFolioSuggestChangesOptions,
): DecodedSuggestChangesInput | string => {
  const normalizations: FolioAgentInputNormalization[] = [];
  let args = rawArgs;
  const embeddedArgs = parseEmbeddedJson(args);
  if (isPlainObject(embeddedArgs)) {
    normalizations.push({ path: "$", message: "arguments were supplied as a JSON string" });
    args = embeddedArgs;
  }
  if (!isPlainObject(args)) {
    return "suggest_changes requires an `operations` array.";
  }
  for (const key of Object.keys(args)) {
    if (!(SUGGEST_CHANGES_ARGUMENT_KEYS as readonly string[]).includes(key)) {
      normalizations.push({ path: key, message: `unknown argument \`${key}\` was ignored` });
    }
  }

  let rawOperations = args["operations"];
  const embeddedOperations = parseEmbeddedJson(rawOperations);
  if (Array.isArray(embeddedOperations)) {
    normalizations.push({
      path: "operations",
      message: "`operations` was supplied as a JSON string",
    });
    rawOperations = embeddedOperations;
  }
  if (!Array.isArray(rawOperations)) {
    return "suggest_changes requires an `operations` array.";
  }
  if (rawOperations.length === 0) {
    return "suggest_changes' `operations` array must not be empty.";
  }
  if (rawOperations.length > options.maxOperations) {
    return `suggest_changes' \`operations\` array has ${rawOperations.length.toLocaleString()} entries, over the ${options.maxOperations}-operation limit; batch it across multiple suggest_changes calls.`;
  }

  const operations: Record<string, unknown>[] = [];
  for (const [index, rawOperation] of rawOperations.entries()) {
    const path = `operations[${index}]`;
    let operation = rawOperation;
    const embeddedOperation = parseEmbeddedJson(operation);
    if (isPlainObject(embeddedOperation)) {
      normalizations.push({ path, message: "operation was supplied as a JSON string" });
      operation = embeddedOperation;
    }
    if (!isPlainObject(operation)) {
      return `${path} must be an object.`;
    }
    let type = operation["type"];
    let typeKey = "type";
    if (type === undefined && typeof operation["kind"] === "string") {
      normalizations.push({ path: `${path}.kind`, message: "`kind` was read as `type`" });
      type = operation["kind"];
      typeKey = "kind";
    }
    if (!isAllowedOperationType(type, options)) {
      return `${path}.type must be one of ${options.operationTypes.join(", ")}.`;
    }
    const allowedKeys: readonly string[] = FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE[type];
    const decoded: Record<string, unknown> = { type };
    for (const [key, value] of Object.entries(operation)) {
      if (key === typeKey) {
        continue;
      }
      if (allowedKeys.includes(key)) {
        decoded[key] = value;
        continue;
      }
      normalizations.push({
        path: `${path}.${key}`,
        message: `\`${key}\` does not apply to ${type} and was ignored`,
      });
    }
    operations.push(decoded);
  }
  return { operations, documentVersion: args["documentVersion"], normalizations };
};

const isAllowedOperationType = (
  value: unknown,
  options: ResolvedFolioSuggestChangesOptions,
): value is FolioDocumentOperationType =>
  typeof value === "string" &&
  (options.operationTypes as readonly string[]).includes(value) &&
  Object.hasOwn(FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE, value);

/** Enforce the per-field and aggregate text caps on one decoded operation, or explain the breach. */
const checkTextCaps = (
  operation: Record<string, unknown>,
  index: number,
  budget: TextBudget,
): string | undefined => {
  const label = (field: string) => `operations[${index}].${field}`;
  const charge = (field: string, value: unknown): string | undefined => {
    if (typeof value !== "string") {
      return undefined;
    }
    if (value.length > MAX_OPERATION_TEXT_LENGTH) {
      return explainTextTooLong(label(field), value.length);
    }
    return consumeTextBudget(budget, value.length) ? explainAggregateTextTooLong(index) : undefined;
  };
  for (const field of TEXT_FIELDS) {
    const breach = charge(field, operation[field]);
    if (breach !== undefined) {
      return breach;
    }
  }
  const comment = operation["comment"];
  const commentBreach = charge("comment", isPlainObject(comment) ? comment["text"] : comment);
  if (commentBreach !== undefined) {
    return commentBreach;
  }
  const cellTexts = operation["cellTexts"];
  if (Array.isArray(cellTexts)) {
    if (cellTexts.length > MAX_TABLE_INSERTION_CELL_TEXTS) {
      return `${label("cellTexts")} has ${cellTexts.length.toLocaleString()} entries, over the ${MAX_TABLE_INSERTION_CELL_TEXTS.toLocaleString()}-cell limit.`;
    }
    for (const [cellIndex, cellText] of cellTexts.entries()) {
      const breach = charge(`cellTexts[${cellIndex}]`, cellText);
      if (breach !== undefined) {
        return breach;
      }
    }
  }
  const parties = operation["parties"];
  if (Array.isArray(parties)) {
    for (const [partyIndex, party] of parties.entries()) {
      if (!isPlainObject(party)) {
        continue;
      }
      for (const field of ["name", "signatory", "title"] as const) {
        const breach = charge(`parties[${partyIndex}].${field}`, party[field]);
        if (breach !== undefined) {
          return breach;
        }
      }
    }
  }
  return undefined;
};

/**
 * Ids minted for operations the model left unnamed. The per-call UUID keeps
 * them unique across calls so a host queue can key on them without
 * rewriting; the index keeps them readable in the model's own output.
 */
const mintOperationIdPrefix = (): string => `op-${crypto.randomUUID()}`;

/** Turn a contract-parser failure into the model-facing message shape (`operations[0].find: expected a string.`). */
const explainContractError = (error: unknown): string => {
  if (error instanceof InvalidFolioDocumentOperationBatchError) {
    return `${error.path.replace(/^\$\.?/, "")}: ${error.reason}.`;
  }
  return error instanceof Error ? error.message : String(error);
};

/**
 * Validate `suggest_changes`' raw tool-call arguments and build the
 * {@link FolioAIEditOperation}s it applies. Pure: does not touch a bridge or
 * document. `options` must match the {@link FolioSuggestChangesOptions} the
 * tool definition was built with.
 */
export const parseSuggestChangesInput = (
  args: unknown,
  options?: FolioSuggestChangesOptions,
): ParseSuggestChangesResult => {
  const resolved = resolveSuggestChangesOptions(options);
  const decoded = decodeSuggestChangesInput(args, resolved);
  if (typeof decoded === "string") {
    return { ok: false, error: decoded };
  }

  let precondition: FolioDocumentOperationBatchPrecondition | undefined;
  if (resolved.documentVersion !== null) {
    if (!isNonEmptyString(decoded.documentVersion)) {
      return {
        ok: false,
        error:
          "suggest_changes requires `documentVersion`: copy the current document version from the tool schema.",
      };
    }
    precondition = { documentVersion: decoded.documentVersion };
  } else if (decoded.documentVersion !== undefined) {
    decoded.normalizations.push({
      path: "documentVersion",
      message: "`documentVersion` is not pinned on this surface and was ignored",
    });
  }

  const budget: TextBudget = { remaining: MAX_TOTAL_OPERATION_TEXT_LENGTH };
  const idPrefix = mintOperationIdPrefix();
  const prepared: Record<string, unknown>[] = [];
  for (const [index, operation] of decoded.operations.entries()) {
    const breach = checkTextCaps(operation, index, budget);
    if (breach !== undefined) {
      return { ok: false, error: breach };
    }
    if (resolved.reviewMeta === "required") {
      if (operation["severity"] === undefined) {
        return {
          ok: false,
          error: `operations[${index}].severity is required on this surface ("low", "medium", or "high").`,
        };
      }
      if (!isNonEmptyString(operation["area"])) {
        return {
          ok: false,
          error: `operations[${index}].area is required on this surface: a short review area label.`,
        };
      }
    }
    const id = operation["id"];
    if (id !== undefined && !isNonEmptyString(id)) {
      return {
        ok: false,
        error: `operations[${index}].id must be a non-empty string when provided.`,
      };
    }
    const comment = operation["comment"];
    prepared.push({
      ...operation,
      id: id ?? `${idPrefix}-${index + 1}`,
      ...(typeof comment === "string" && { comment: { text: comment } }),
    });
  }

  // Boundary adapter over the intentionally throwing contract parser: an
  // invalid operation is an expected outcome here, reported to the model.
  try {
    const batch = parseFolioDocumentOperationBatch({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      operations: prepared,
      ...(precondition !== undefined && { precondition }),
    });
    return {
      ok: true,
      operations: [...batch.operations],
      ...(batch.precondition !== undefined && { precondition: batch.precondition }),
      normalizations: decoded.normalizations,
    };
  } catch (error) {
    return { ok: false, error: explainContractError(error) };
  }
};
