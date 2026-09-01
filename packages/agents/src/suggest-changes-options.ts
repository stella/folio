/**
 * Host-facing configuration of the `suggest_changes` tool. One resolved
 * options object drives the tool's JSON Schema (`tools.ts`), its parser
 * (`parse.ts`), and its capability description, so a host that needs a
 * narrower or stricter surface configures instead of re-deriving the
 * contract.
 */

import { TaggedError } from "better-result";
import {
  FOLIO_DOCUMENT_OPERATION_TYPES,
  type FolioDocumentOperationType,
} from "@stll/folio-core/server";

/** Whether every operation must carry `severity` and `area`. */
export type FolioSuggestChangesReviewMetaPolicy = "optional" | "required";

/**
 * Pin the batch to a host document version. The model must echo `current`
 * as the tool's top-level `documentVersion` argument; it lands on the batch
 * as `precondition.documentVersion`, and the executor skips the whole batch
 * with `documentVersionMismatch` when the bridge reports a different version
 * at apply time (an approval that ran after the document moved on).
 */
export type FolioSuggestChangesDocumentVersionOption = {
  current: string;
};

/**
 * Per-surface shape of the `suggest_changes` tool. Pass the same object to
 * `getFolioToolDefinitions` and `executeFolioToolCall` so the schema the
 * model sees and the parser that checks its calls agree.
 */
export type FolioSuggestChangesOptions = {
  /**
   * Operation types the model may emit, in any order. Defaults to every
   * contract type except `commentOnBlock` (covered by `add_comment`) and
   * `insertSignatureTable` (direct-only, so not a tracked change).
   */
  operationTypes?: readonly FolioDocumentOperationType[];
  /** Defaults to `"optional"`. */
  reviewMeta?: FolioSuggestChangesReviewMetaPolicy;
  /** Operations accepted per call, 1 to 200. Defaults to 50. */
  maxOperations?: number;
  documentVersion?: FolioSuggestChangesDocumentVersionOption;
};

/**
 * Per-surface configuration shared by `getFolioToolDefinitions` and
 * `executeFolioToolCall`; every tool but `suggest_changes` is fixed.
 */
export type FolioAgentToolOptions = {
  suggestChanges?: FolioSuggestChangesOptions;
};

export type ResolvedFolioSuggestChangesOptions = {
  /** Allowed types in contract order, deduplicated. */
  readonly operationTypes: readonly FolioDocumentOperationType[];
  readonly reviewMeta: FolioSuggestChangesReviewMetaPolicy;
  readonly maxOperations: number;
  readonly documentVersion: FolioSuggestChangesDocumentVersionOption | null;
};

const DEFAULT_EXCLUDED_OPERATION_TYPES: ReadonlySet<FolioDocumentOperationType> = new Set([
  "commentOnBlock",
  "insertSignatureTable",
]);

/** The operation types `suggest_changes` exposes when a host passes no `operationTypes`. */
export const DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES: readonly FolioDocumentOperationType[] =
  FOLIO_DOCUMENT_OPERATION_TYPES.filter((type) => !DEFAULT_EXCLUDED_OPERATION_TYPES.has(type));

export const DEFAULT_MAX_OPERATIONS_PER_CALL = 50;
export const MAX_OPERATIONS_PER_CALL_LIMIT = 200;

/** Programmer misuse of {@link FolioSuggestChangesOptions}; never reaches the model. */
export class InvalidFolioSuggestChangesOptionsError extends TaggedError(
  "InvalidFolioSuggestChangesOptionsError",
)<{
  message: string;
  option: keyof FolioSuggestChangesOptions;
}> {}

const invalidOption = (option: keyof FolioSuggestChangesOptions, message: string): never => {
  throw new InvalidFolioSuggestChangesOptionsError({
    message: `Invalid suggest_changes option \`${option}\`: ${message}.`,
    option,
  });
};

const isContractOperationType = (value: unknown): value is FolioDocumentOperationType =>
  typeof value === "string" &&
  (FOLIO_DOCUMENT_OPERATION_TYPES as readonly string[]).includes(value);

const resolveOperationTypes = (
  requested: readonly FolioDocumentOperationType[] | undefined,
): readonly FolioDocumentOperationType[] => {
  if (requested === undefined) {
    return DEFAULT_SUGGEST_CHANGES_OPERATION_TYPES;
  }
  if (!Array.isArray(requested)) {
    return invalidOption("operationTypes", "expected an array of operation types");
  }
  if (requested.length === 0) {
    return invalidOption("operationTypes", "expected at least one operation type");
  }
  const unknown = requested.find((type) => !isContractOperationType(type));
  if (unknown !== undefined) {
    return invalidOption("operationTypes", `"${String(unknown)}" is not a contract operation type`);
  }
  const wanted = new Set(requested);
  return FOLIO_DOCUMENT_OPERATION_TYPES.filter((type) => wanted.has(type));
};

const resolveMaxOperations = (requested: number | undefined): number => {
  if (requested === undefined) {
    return DEFAULT_MAX_OPERATIONS_PER_CALL;
  }
  if (!Number.isInteger(requested) || requested < 1 || requested > MAX_OPERATIONS_PER_CALL_LIMIT) {
    return invalidOption(
      "maxOperations",
      `expected an integer from 1 to ${MAX_OPERATIONS_PER_CALL_LIMIT}`,
    );
  }
  return requested;
};

const resolveDocumentVersion = (
  requested: FolioSuggestChangesDocumentVersionOption | undefined,
): FolioSuggestChangesDocumentVersionOption | null => {
  if (requested === undefined) {
    return null;
  }
  if (
    typeof requested !== "object" ||
    requested === null ||
    typeof requested.current !== "string" ||
    requested.current.length === 0
  ) {
    return invalidOption("documentVersion", "expected a non-empty `current` version token");
  }
  return { current: requested.current };
};

/** Validate and default a host's options; throws {@link InvalidFolioSuggestChangesOptionsError} on misuse. */
export const resolveSuggestChangesOptions = (
  options: FolioSuggestChangesOptions = {},
): ResolvedFolioSuggestChangesOptions => ({
  operationTypes: resolveOperationTypes(options.operationTypes),
  reviewMeta: options.reviewMeta ?? "optional",
  maxOperations: resolveMaxOperations(options.maxOperations),
  documentVersion: resolveDocumentVersion(options.documentVersion),
});
