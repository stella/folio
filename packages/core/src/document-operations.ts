import { TaggedError } from "better-result";

import { applyFolioAIEditOperations, type FolioAIEditView } from "./ai-edits/apply";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditOperation,
  FolioAIEditReviewMeta,
  FolioAIEditSeverity,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
} from "./ai-edits/types";

export const FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION = 1 as const;

export const FOLIO_DOCUMENT_OPERATION_TYPES = Object.freeze([
  "replaceInBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
  "replaceBlock",
  "deleteBlock",
  "commentOnBlock",
  "insertSignatureTable",
] as const satisfies readonly FolioAIEditOperation["type"][]);

export const FOLIO_DOCUMENT_OPERATION_MODES = Object.freeze([
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioAIEditApplyMode[]);

export const FOLIO_DOCUMENT_OPERATION_STORIES = Object.freeze(["main"] as const);

export type FolioDocumentOperation = FolioAIEditOperation;
export type FolioDocumentOperationMode = FolioAIEditApplyMode;
export type FolioDocumentOperationType = FolioDocumentOperation["type"];

const DIRECT_AND_TRACKED_MODES = FOLIO_DOCUMENT_OPERATION_MODES;
const DIRECT_ONLY_MODES = Object.freeze([
  "direct",
] as const satisfies readonly FolioDocumentOperationMode[]);

export const FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE = Object.freeze({
  replaceInBlock: DIRECT_AND_TRACKED_MODES,
  insertAfterBlock: DIRECT_AND_TRACKED_MODES,
  insertBeforeBlock: DIRECT_AND_TRACKED_MODES,
  replaceBlock: DIRECT_AND_TRACKED_MODES,
  deleteBlock: DIRECT_AND_TRACKED_MODES,
  commentOnBlock: DIRECT_AND_TRACKED_MODES,
  insertSignatureTable: DIRECT_ONLY_MODES,
} as const satisfies Readonly<
  Record<FolioDocumentOperationType, readonly FolioDocumentOperationMode[]>
>);

export type FolioDocumentOperationCapabilities = {
  readonly version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  readonly operationTypes: typeof FOLIO_DOCUMENT_OPERATION_TYPES;
  readonly modes: typeof FOLIO_DOCUMENT_OPERATION_MODES;
  readonly modesByOperationType: typeof FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE;
  readonly stories: typeof FOLIO_DOCUMENT_OPERATION_STORIES;
};

const DOCUMENT_OPERATION_CAPABILITIES = Object.freeze({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  operationTypes: FOLIO_DOCUMENT_OPERATION_TYPES,
  modes: FOLIO_DOCUMENT_OPERATION_MODES,
  modesByOperationType: FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  stories: FOLIO_DOCUMENT_OPERATION_STORIES,
} as const satisfies FolioDocumentOperationCapabilities);

export const getFolioDocumentOperationCapabilities = (): FolioDocumentOperationCapabilities =>
  DOCUMENT_OPERATION_CAPABILITIES;

const includesDocumentOperationMode = (
  supportedModes: readonly FolioDocumentOperationMode[],
  mode: FolioDocumentOperationMode,
): boolean => supportedModes.includes(mode);

export const isFolioDocumentOperationModeSupported = (
  operationType: FolioDocumentOperationType,
  mode: FolioDocumentOperationMode,
): boolean => {
  const supportedModes = FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE[operationType];
  if (supportedModes === undefined) {
    return false;
  }
  return includesDocumentOperationMode(supportedModes, mode);
};

export const isSupportedFolioDocumentOperationVersion = (
  value: unknown,
): value is typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION =>
  value === FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;

export class UnsupportedFolioDocumentOperationVersionError extends TaggedError(
  "UnsupportedFolioDocumentOperationVersionError",
)<{
  message: string;
  receivedVersion: unknown;
}>() {}

export class InvalidFolioDocumentOperationBatchError extends TaggedError(
  "InvalidFolioDocumentOperationBatchError",
)<{
  message: string;
  path: string;
  reason: string;
}>() {}

export const assertSupportedFolioDocumentOperationVersion = (
  value: unknown,
): typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION => {
  if (isSupportedFolioDocumentOperationVersion(value)) {
    return value;
  }
  throw new UnsupportedFolioDocumentOperationVersionError({
    message: "Unsupported document operation contract version.",
    receivedVersion: value,
  });
};

export type FolioDocumentOperationBatch = {
  readonly version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  operations: FolioDocumentOperation[];
  mode?: FolioDocumentOperationMode;
};

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const invalidBatch = (path: string, reason: string): never => {
  throw new InvalidFolioDocumentOperationBatchError({
    message: `Invalid document operation batch at ${path}: ${reason}.`,
    path,
    reason,
  });
};

const assertAllowedKeys = (
  value: Record<string, unknown>,
  path: string,
  allowedKeys: readonly string[],
): void => {
  const unexpected = Object.keys(value).find((key) => !allowedKeys.includes(key));
  if (unexpected !== undefined) {
    invalidBatch(`${path}.${unexpected}`, "unexpected property");
  }
};

const readString = (value: Record<string, unknown>, key: string, path: string): string => {
  const candidate = value[key];
  if (typeof candidate === "string") {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a string");
};

const readOptionalString = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): string | undefined => {
  const candidate = value[key];
  if (candidate === undefined) {
    return undefined;
  }
  if (typeof candidate === "string") {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a string when provided");
};

const readOptionalBoolean = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): boolean | undefined => {
  const candidate = value[key];
  if (candidate === undefined) {
    return undefined;
  }
  if (typeof candidate === "boolean") {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a boolean when provided");
};

const readOptionalComment = (
  value: Record<string, unknown>,
  path: string,
): { text: string } | undefined => {
  const candidate = value["comment"];
  if (candidate === undefined) {
    return undefined;
  }
  if (isPlainObject(candidate)) {
    assertAllowedKeys(candidate, `${path}.comment`, ["text"]);
    return { text: readString(candidate, "text", `${path}.comment`) };
  }
  return invalidBatch(`${path}.comment`, "expected an object when provided");
};

const isReviewSeverity = (value: unknown): value is FolioAIEditSeverity =>
  value === "low" || value === "medium" || value === "high";

const readReviewMeta = (value: Record<string, unknown>, path: string): FolioAIEditReviewMeta => {
  const severity = value["severity"];
  if (severity !== undefined && !isReviewSeverity(severity)) {
    return invalidBatch(`${path}.severity`, 'expected "low", "medium", or "high" when provided');
  }
  const area = readOptionalString(value, "area", path);
  return {
    ...(severity !== undefined && { severity }),
    ...(area !== undefined && { area }),
  };
};

const COMMON_OPERATION_KEYS = ["id", "type", "blockId", "severity", "area"] as const;

const parseSignatureParties = (
  value: Record<string, unknown>,
  path: string,
): { name: string; signatory?: string; title?: string }[] => {
  const parties = value["parties"];
  if (!Array.isArray(parties)) {
    return invalidBatch(`${path}.parties`, "expected an array");
  }
  return parties.map((party, index) => {
    const partyPath = `${path}.parties[${index}]`;
    if (!isPlainObject(party)) {
      return invalidBatch(partyPath, "expected an object");
    }
    assertAllowedKeys(party, partyPath, ["name", "signatory", "title"]);
    const signatory = readOptionalString(party, "signatory", partyPath);
    const title = readOptionalString(party, "title", partyPath);
    return {
      name: readString(party, "name", partyPath),
      ...(signatory !== undefined && { signatory }),
      ...(title !== undefined && { title }),
    };
  });
};

const parseDocumentOperation = (value: unknown, index: number): FolioDocumentOperation => {
  const path = `$.operations[${index}]`;
  if (!isPlainObject(value)) {
    return invalidBatch(path, "expected an object");
  }

  const id = readString(value, "id", path);
  const type = readString(value, "type", path);
  const blockId = readString(value, "blockId", path);
  const reviewMeta = readReviewMeta(value, path);
  const comment = readOptionalComment(value, path);

  if (type === "replaceInBlock") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "find", "replace", "comment"]);
    return {
      ...reviewMeta,
      id,
      type,
      blockId,
      find: readString(value, "find", path),
      replace: readString(value, "replace", path),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "insertAfterBlock" || type === "insertBeforeBlock") {
    assertAllowedKeys(value, path, [
      ...COMMON_OPERATION_KEYS,
      "text",
      "inheritFormatting",
      "pageBreakBefore",
      "styleId",
      "comment",
    ]);
    const inheritFormatting = readOptionalBoolean(value, "inheritFormatting", path);
    const pageBreakBefore = readOptionalBoolean(value, "pageBreakBefore", path);
    const styleId = readOptionalString(value, "styleId", path);
    return {
      ...reviewMeta,
      id,
      type,
      blockId,
      text: readString(value, "text", path),
      ...(inheritFormatting !== undefined && { inheritFormatting }),
      ...(pageBreakBefore !== undefined && { pageBreakBefore }),
      ...(styleId !== undefined && { styleId }),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "replaceBlock") {
    assertAllowedKeys(value, path, [
      ...COMMON_OPERATION_KEYS,
      "text",
      "preserveFormatting",
      "styleId",
      "comment",
    ]);
    const preserveFormatting = readOptionalBoolean(value, "preserveFormatting", path);
    const styleId = readOptionalString(value, "styleId", path);
    return {
      ...reviewMeta,
      id,
      type,
      blockId,
      text: readString(value, "text", path),
      ...(preserveFormatting !== undefined && { preserveFormatting }),
      ...(styleId !== undefined && { styleId }),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "deleteBlock") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "comment"]);
    return { ...reviewMeta, id, type, blockId, ...(comment !== undefined && { comment }) };
  }

  if (type === "commentOnBlock") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "quote", "comment"]);
    if (comment === undefined) {
      return invalidBatch(`${path}.comment`, "expected an object");
    }
    const quote = readOptionalString(value, "quote", path);
    return {
      ...reviewMeta,
      id,
      type,
      blockId,
      ...(quote !== undefined && { quote }),
      comment,
    };
  }

  if (type === "insertSignatureTable") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "position", "parties", "comment"]);
    const position = value["position"];
    if (position !== undefined && position !== "after" && position !== "before") {
      return invalidBatch(`${path}.position`, 'expected "after" or "before" when provided');
    }
    return {
      ...reviewMeta,
      id,
      type,
      blockId,
      ...(position !== undefined && { position }),
      parties: parseSignatureParties(value, path),
      ...(comment !== undefined && { comment }),
    };
  }

  return invalidBatch(`${path}.type`, `unsupported operation type "${type}"`);
};

export const parseFolioDocumentOperationBatch = (value: unknown): FolioDocumentOperationBatch => {
  if (!isPlainObject(value)) {
    return invalidBatch("$", "expected an object");
  }
  assertAllowedKeys(value, "$", ["version", "operations", "mode"]);
  const version = assertSupportedFolioDocumentOperationVersion(value["version"]);
  const operations = value["operations"];
  if (!Array.isArray(operations)) {
    return invalidBatch("$.operations", "expected an array");
  }
  const mode = value["mode"];
  if (mode !== undefined && mode !== "direct" && mode !== "tracked-changes") {
    return invalidBatch("$.mode", 'expected "direct" or "tracked-changes" when provided');
  }
  return {
    version,
    operations: operations.map(parseDocumentOperation),
    ...(mode !== undefined && { mode }),
  };
};

export type FolioDocumentOperationResult = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
};

export type ApplyFolioDocumentOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  batch: FolioDocumentOperationBatch;
  author?: string;
  createCommentId?: (text: string) => number;
};

export const applyFolioDocumentOperations = ({
  view,
  snapshot,
  batch,
  author,
  createCommentId,
}: ApplyFolioDocumentOperationsOptions): FolioDocumentOperationResult => {
  const parsedBatch = parseFolioDocumentOperationBatch(batch);
  return {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    ...applyFolioAIEditOperations({
      view,
      snapshot,
      operations: parsedBatch.operations,
      mode: parsedBatch.mode ?? "tracked-changes",
      ...(author !== undefined && { author }),
      ...(createCommentId !== undefined && { createCommentId }),
    }),
  };
};
