import { TaggedError } from "better-result";

import {
  applyFolioAIEditOperations,
  type FolioAIEditView,
  previewFolioAIEditOperations,
} from "./ai-edits/apply";
import type {
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditOperation,
  FolioAIEditPrecondition,
  FolioAIEditReviewMeta,
  FolioAIEditSeverity,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
  FolioAITextRangeHandle,
} from "./ai-edits/types";

export const FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION = 1 as const;

export const FOLIO_DOCUMENT_OPERATION_TYPES = Object.freeze([
  "replaceInBlock",
  "replaceRange",
  "commentOnRange",
  "formatRange",
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
export const FOLIO_DOCUMENT_OPERATION_PRECONDITIONS = Object.freeze(["blockTextHash"] as const);
export const FOLIO_DOCUMENT_OPERATION_BATCH_MODES = Object.freeze([
  "best-effort",
  "atomic",
] as const);

export type FolioDocumentOperation = FolioAIEditOperation;
export type FolioDocumentOperationMode = FolioAIEditApplyMode;
export type FolioDocumentOperationPrecondition = FolioAIEditPrecondition;
export type FolioDocumentOperationType = FolioDocumentOperation["type"];

const DIRECT_AND_TRACKED_MODES = FOLIO_DOCUMENT_OPERATION_MODES;
const DIRECT_ONLY_MODES = Object.freeze([
  "direct",
] as const satisfies readonly FolioDocumentOperationMode[]);

export const FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE = Object.freeze({
  replaceInBlock: DIRECT_AND_TRACKED_MODES,
  replaceRange: DIRECT_AND_TRACKED_MODES,
  commentOnRange: DIRECT_AND_TRACKED_MODES,
  formatRange: DIRECT_ONLY_MODES,
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
  readonly batchModes: typeof FOLIO_DOCUMENT_OPERATION_BATCH_MODES;
  readonly dryRun: true;
  readonly preconditions: typeof FOLIO_DOCUMENT_OPERATION_PRECONDITIONS;
  readonly stories: typeof FOLIO_DOCUMENT_OPERATION_STORIES;
};

const DOCUMENT_OPERATION_CAPABILITIES = Object.freeze({
  version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  operationTypes: FOLIO_DOCUMENT_OPERATION_TYPES,
  modes: FOLIO_DOCUMENT_OPERATION_MODES,
  modesByOperationType: FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE,
  batchModes: FOLIO_DOCUMENT_OPERATION_BATCH_MODES,
  dryRun: true,
  preconditions: FOLIO_DOCUMENT_OPERATION_PRECONDITIONS,
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
  atomic?: boolean;
  dryRun?: boolean;
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

const readNonNegativeInteger = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): number => {
  const candidate = value[key];
  if (typeof candidate === "number" && Number.isInteger(candidate) && candidate >= 0) {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a non-negative integer");
};

const readTextRange = (value: Record<string, unknown>, path: string): FolioAITextRangeHandle => {
  const candidate = value["range"];
  const rangePath = `${path}.range`;
  if (!isPlainObject(candidate)) {
    return invalidBatch(rangePath, "expected an object");
  }
  assertAllowedKeys(candidate, rangePath, [
    "type",
    "story",
    "blockId",
    "startOffset",
    "endOffset",
    "selectedTextHash",
  ]);
  if (candidate["type"] !== "textRange") {
    return invalidBatch(`${rangePath}.type`, 'expected "textRange"');
  }
  if (candidate["story"] !== "main") {
    return invalidBatch(`${rangePath}.story`, 'expected "main"');
  }
  const blockId = readString(candidate, "blockId", rangePath);
  if (blockId.length === 0) {
    return invalidBatch(`${rangePath}.blockId`, "expected a non-empty string");
  }
  const startOffset = readNonNegativeInteger(candidate, "startOffset", rangePath);
  const endOffset = readNonNegativeInteger(candidate, "endOffset", rangePath);
  if (endOffset <= startOffset) {
    return invalidBatch(`${rangePath}.endOffset`, "expected a value greater than startOffset");
  }
  const selectedTextHash = readString(candidate, "selectedTextHash", rangePath);
  if (!/^h[0-9a-z]+$/.test(selectedTextHash)) {
    return invalidBatch(`${rangePath}.selectedTextHash`, "expected a normalized text hash");
  }
  return {
    type: "textRange",
    story: "main",
    blockId,
    startOffset,
    endOffset,
    selectedTextHash,
  };
};

const readInlineFormatting = (
  value: Record<string, unknown>,
  path: string,
): { bold?: boolean; italic?: boolean; underline?: boolean } => {
  const candidate = value["formatting"];
  const formattingPath = `${path}.formatting`;
  if (!isPlainObject(candidate)) {
    return invalidBatch(formattingPath, "expected an object");
  }
  assertAllowedKeys(candidate, formattingPath, ["bold", "italic", "underline"]);
  const bold = readOptionalBoolean(candidate, "bold", formattingPath);
  const italic = readOptionalBoolean(candidate, "italic", formattingPath);
  const underline = readOptionalBoolean(candidate, "underline", formattingPath);
  if (bold === undefined && italic === undefined && underline === undefined) {
    return invalidBatch(formattingPath, "expected at least one formatting property");
  }
  return {
    ...(bold !== undefined && { bold }),
    ...(italic !== undefined && { italic }),
    ...(underline !== undefined && { underline }),
  };
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

const readOptionalPrecondition = (
  value: Record<string, unknown>,
  path: string,
): FolioAIEditPrecondition | undefined => {
  const candidate = value["precondition"];
  if (candidate === undefined) {
    return undefined;
  }
  if (!isPlainObject(candidate)) {
    return invalidBatch(`${path}.precondition`, "expected an object when provided");
  }
  const preconditionPath = `${path}.precondition`;
  assertAllowedKeys(candidate, preconditionPath, ["blockTextHash"]);
  const blockTextHash = readString(candidate, "blockTextHash", preconditionPath);
  if (!/^h[0-9a-z]+$/.test(blockTextHash)) {
    return invalidBatch(
      `${preconditionPath}.blockTextHash`,
      "expected a normalized block text hash",
    );
  }
  return { blockTextHash };
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

const COMMON_OPERATION_KEYS = [
  "id",
  "type",
  "blockId",
  "severity",
  "area",
  "precondition",
] as const;

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
  const reviewMeta = readReviewMeta(value, path);
  const precondition = readOptionalPrecondition(value, path);
  const comment = readOptionalComment(value, path);
  const operationMeta = {
    ...reviewMeta,
    ...(precondition !== undefined && { precondition }),
  };

  if (type === "replaceRange") {
    assertAllowedKeys(value, path, [
      "id",
      "type",
      "range",
      "replace",
      "comment",
      "severity",
      "area",
      "precondition",
    ]);
    return {
      ...operationMeta,
      id,
      type,
      range: readTextRange(value, path),
      replace: readString(value, "replace", path),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "commentOnRange") {
    assertAllowedKeys(value, path, [
      "id",
      "type",
      "range",
      "comment",
      "severity",
      "area",
      "precondition",
    ]);
    if (comment === undefined) {
      return invalidBatch(`${path}.comment`, "expected an object");
    }
    return { ...operationMeta, id, type, range: readTextRange(value, path), comment };
  }

  if (type === "formatRange") {
    assertAllowedKeys(value, path, [
      "id",
      "type",
      "range",
      "formatting",
      "severity",
      "area",
      "precondition",
    ]);
    return {
      ...operationMeta,
      id,
      type,
      range: readTextRange(value, path),
      formatting: readInlineFormatting(value, path),
    };
  }

  const blockId = readString(value, "blockId", path);

  if (type === "replaceInBlock") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "find", "replace", "comment"]);
    return {
      ...operationMeta,
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
      ...operationMeta,
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
      ...operationMeta,
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
    return { ...operationMeta, id, type, blockId, ...(comment !== undefined && { comment }) };
  }

  if (type === "commentOnBlock") {
    assertAllowedKeys(value, path, [...COMMON_OPERATION_KEYS, "quote", "comment"]);
    if (comment === undefined) {
      return invalidBatch(`${path}.comment`, "expected an object");
    }
    const quote = readOptionalString(value, "quote", path);
    return {
      ...operationMeta,
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
      ...operationMeta,
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
  assertAllowedKeys(value, "$", ["version", "operations", "mode", "atomic", "dryRun"]);
  const version = assertSupportedFolioDocumentOperationVersion(value["version"]);
  const operations = value["operations"];
  if (!Array.isArray(operations)) {
    return invalidBatch("$.operations", "expected an array");
  }
  const mode = value["mode"];
  if (mode !== undefined && mode !== "direct" && mode !== "tracked-changes") {
    return invalidBatch("$.mode", 'expected "direct" or "tracked-changes" when provided');
  }
  const atomic = readOptionalBoolean(value, "atomic", "$");
  const dryRun = readOptionalBoolean(value, "dryRun", "$");
  const parsedOperations = operations.map(parseDocumentOperation);
  const operationIds = new Set<string>();
  for (const [index, operation] of parsedOperations.entries()) {
    if (operationIds.has(operation.id)) {
      return invalidBatch(`$.operations[${index}].id`, "expected a unique operation id");
    }
    operationIds.add(operation.id);
  }
  return {
    version,
    operations: parsedOperations,
    ...(mode !== undefined && { mode }),
    ...(atomic !== undefined && { atomic }),
    ...(dryRun !== undefined && { dryRun }),
  };
};

export type FolioDocumentOperationStatus = "committed" | "previewed" | "rejected";

export type FolioDocumentOperationResult = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  status: FolioDocumentOperationStatus;
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

type ApplyParsedDocumentOperationBatchOptions = {
  targetView: FolioAIEditView;
  targetCreateCommentId?: (text: string) => number;
  preview?: boolean;
};

export const applyFolioDocumentOperations = ({
  view,
  snapshot,
  batch,
  author,
  createCommentId,
}: ApplyFolioDocumentOperationsOptions): FolioDocumentOperationResult => {
  const parsedBatch = parseFolioDocumentOperationBatch(batch);
  const apply = ({
    targetView,
    targetCreateCommentId = createCommentId,
    preview = false,
  }: ApplyParsedDocumentOperationBatchOptions) => {
    const applyOperations = preview ? previewFolioAIEditOperations : applyFolioAIEditOperations;
    return applyOperations({
      view: targetView,
      snapshot,
      operations: parsedBatch.operations,
      mode: parsedBatch.mode ?? "tracked-changes",
      ...(author !== undefined && { author }),
      ...(targetCreateCommentId !== undefined && { createCommentId: targetCreateCommentId }),
    });
  };

  const preview = () => apply({ targetView: view, preview: true });

  const atomicResult = (
    previewResult: FolioAIEditApplyResult,
    status: "previewed" | "rejected",
  ): FolioDocumentOperationResult => {
    const skippedById = new Map(
      previewResult.skipped.map((operation) => [operation.id, operation]),
    );
    return {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status,
      applied: [],
      skipped: parsedBatch.operations.map(
        ({ id }): FolioAIEditSkippedOperation =>
          skippedById.get(id) ?? { id, reason: "atomicBatchRejected" },
      ),
    };
  };

  if (parsedBatch.dryRun === true) {
    const previewResult = preview();
    if (parsedBatch.atomic === true && previewResult.skipped.length > 0) {
      return atomicResult(previewResult, "previewed");
    }
    return {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status: "previewed",
      applied: previewResult.applied.map(({ id }) => ({ id })),
      skipped: previewResult.skipped,
    };
  }

  if (parsedBatch.atomic === true) {
    const previewResult = preview();
    if (previewResult.skipped.length > 0) {
      return atomicResult(previewResult, "rejected");
    }
  }

  return {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    status: "committed",
    ...apply({ targetView: view }),
  };
};
