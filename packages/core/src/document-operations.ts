import { TaggedError } from "better-result";

import type { FolioTableTemplates } from "./ai-edits/table-template";
import { paragraphMarkFormattingTemplatesOf } from "./ai-edits/paragraph-mark-template";
import {
  applyFolioAIEditOperations,
  type FolioAIEditApplyOutcome,
  type FolioAIEditView,
  type FolioWordDiffOptions,
  type FolioRevisionStamp,
  previewFolioAIEditOperations,
} from "./ai-edits/apply";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditNormalization,
  FolioAIEditOperation,
  FolioAIEditPrecondition,
  FolioAIEditReviewMeta,
  FolioAIEditSeverity,
  FolioAIEditSkippedOperation,
  FolioAIEditSnapshot,
  FolioAIInlineFormatting,
  FolioAIParagraphSpacing,
  FolioAITextRangeHandle,
} from "./ai-edits/types";
import type { LineSpacingRule, ParagraphAlignment } from "./types/document";
import { LINE_SPACING_RULE_VALUES, PARAGRAPH_ALIGNMENT_VALUES } from "./types/documentEnumValues";

export const FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION = 1 as const;

/** Direct paragraph-alignment values accepted by the operation contract. */
export const FOLIO_PARAGRAPH_ALIGNMENT_VALUES = Object.freeze([...PARAGRAPH_ALIGNMENT_VALUES]);

/** `w:spacing/@w:lineRule` values accepted by the operation contract. */
export const FOLIO_LINE_SPACING_RULE_VALUES = Object.freeze([...LINE_SPACING_RULE_VALUES]);

export const FOLIO_DOCUMENT_OPERATION_TYPES = Object.freeze([
  "replaceInBlock",
  "replaceRange",
  "commentOnRange",
  "formatRange",
  "insertAfterBlock",
  "insertBeforeBlock",
  "replaceBlock",
  "deleteBlock",
  "splitBlock",
  "mergeBlockWithNext",
  "setBlockParagraphProperties",
  "insertTable",
  "deleteTable",
  "commentOnBlock",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
  "insertTableColumn",
  "deleteTableColumn",
  "mergeTableCells",
  "splitTableCell",
] as const satisfies readonly FolioAIEditOperation["type"][]);

export const FOLIO_DOCUMENT_OPERATION_MODES = Object.freeze([
  "direct",
  "tracked-changes",
  "suggested",
] as const satisfies readonly FolioAIEditApplyMode[]);

export const FOLIO_DOCUMENT_OPERATION_STORIES = Object.freeze([
  "main",
  "header",
  "footer",
  "footnote",
  "endnote",
] as const);
export const FOLIO_DOCUMENT_OPERATION_PRECONDITIONS = Object.freeze(["blockTextHash"] as const);
/**
 * Batch-level preconditions. `documentVersion` is an opaque host token (an
 * entity version id, a collaboration checkpoint) pinning the document the
 * batch was authored against; whoever knows the live version (an agent
 * bridge, a host applying the batch itself) compares it and skips the whole
 * batch with `documentVersionMismatch` when it differs. The core apply path
 * preserves the token but has no version to compare it to.
 */
export const FOLIO_DOCUMENT_OPERATION_BATCH_PRECONDITIONS = Object.freeze([
  "documentVersion",
] as const);
export const FOLIO_DOCUMENT_OPERATION_BATCH_MODES = Object.freeze([
  "best-effort",
  "atomic",
] as const);

export type FolioDocumentOperation = FolioAIEditOperation;
export type FolioDocumentOperationMode = FolioAIEditApplyMode;
export type FolioDocumentOperationPrecondition = FolioAIEditPrecondition;
export type FolioDocumentOperationType = FolioDocumentOperation["type"];

const parsedFolioDocumentOperationBatches = new WeakSet<object>();

// Suggested mode covers inline text/format edits and block/table structural
// operations whose revisions the serialization strip removes (see apply.ts).
// Comment ops (not tracked changes) and cell merge/split stay direct-and-tracked.
const DIRECT_AND_TRACKED_MODES = Object.freeze([
  "direct",
  "tracked-changes",
] as const satisfies readonly FolioDocumentOperationMode[]);
const DIRECT_TRACKED_AND_SUGGESTED_MODES = FOLIO_DOCUMENT_OPERATION_MODES;
// A whole inserted table has no OOXML tracked representation, so it supports
// direct and suggested (accept applies it directly) but not tracked-changes.
const DIRECT_AND_SUGGESTED_MODES = Object.freeze([
  "direct",
  "suggested",
] as const satisfies readonly FolioDocumentOperationMode[]);

export const FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE = Object.freeze({
  replaceInBlock: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  replaceRange: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  commentOnRange: DIRECT_AND_TRACKED_MODES,
  formatRange: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  insertAfterBlock: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  insertBeforeBlock: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  replaceBlock: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  deleteBlock: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  splitBlock: DIRECT_AND_TRACKED_MODES,
  mergeBlockWithNext: DIRECT_AND_TRACKED_MODES,
  setBlockParagraphProperties: DIRECT_AND_TRACKED_MODES,
  insertTable: DIRECT_AND_TRACKED_MODES,
  deleteTable: DIRECT_AND_TRACKED_MODES,
  commentOnBlock: DIRECT_AND_TRACKED_MODES,
  insertSignatureTable: DIRECT_AND_SUGGESTED_MODES,
  insertTableRow: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  deleteTableRow: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  insertTableColumn: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  deleteTableColumn: DIRECT_TRACKED_AND_SUGGESTED_MODES,
  mergeTableCells: DIRECT_AND_TRACKED_MODES,
  splitTableCell: DIRECT_AND_TRACKED_MODES,
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
  readonly batchPreconditions: typeof FOLIO_DOCUMENT_OPERATION_BATCH_PRECONDITIONS;
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
  batchPreconditions: FOLIO_DOCUMENT_OPERATION_BATCH_PRECONDITIONS,
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
  // A plain indexed lookup resolves "__proto__" / "constructor" / "toString"
  // to an inherited Object.prototype member instead of undefined — an
  // untrusted `operationType` string must be checked as an own property
  // first, or the `.includes` call below throws on that non-array value.
  if (!Object.hasOwn(FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE, operationType)) {
    return false;
  }
  const supportedModes = FOLIO_DOCUMENT_OPERATION_MODES_BY_TYPE[operationType];
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
}> {}

export class InvalidFolioDocumentOperationBatchError extends TaggedError(
  "InvalidFolioDocumentOperationBatchError",
)<{
  message: string;
  path: string;
  reason: string;
}> {}

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

/** Batch-level guard; see {@link FOLIO_DOCUMENT_OPERATION_BATCH_PRECONDITIONS}. */
export type FolioDocumentOperationBatchPrecondition = {
  readonly documentVersion: string;
};

export type FolioDocumentOperationBatch = {
  readonly version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  readonly operations: readonly FolioDocumentOperation[];
  readonly mode?: FolioDocumentOperationMode;
  readonly atomic?: boolean;
  readonly dryRun?: boolean;
  readonly precondition?: FolioDocumentOperationBatchPrecondition;
};

const isPlainObject = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);

const isParsedFolioDocumentOperationBatch = (
  value: unknown,
): value is FolioDocumentOperationBatch =>
  typeof value === "object" && value !== null && parsedFolioDocumentOperationBatches.has(value);

const freezeParsedValue = (value: unknown): void => {
  if (Array.isArray(value)) {
    for (const entry of value) {
      freezeParsedValue(entry);
    }
    Object.freeze(value);
    return;
  }
  if (!isPlainObject(value)) {
    return;
  }
  for (const entry of Object.values(value)) {
    freezeParsedValue(entry);
  }
  Object.freeze(value);
};

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

const readClearableNonEmptyString = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): string | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined || candidate === null) {
    return candidate;
  }
  if (typeof candidate === "string" && candidate.trim().length > 0) {
    return candidate.trim();
  }
  return invalidBatch(`${path}.${key}`, "expected a non-empty string or null when provided");
};

const readClearableFontSize = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): number | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined || candidate === null) {
    return candidate;
  }
  if (typeof candidate === "number" && candidate > 0 && Number.isInteger(candidate * 2)) {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a positive half-point value or null");
};

const readClearableRgbColor = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): string | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined || candidate === null) {
    return candidate;
  }
  if (typeof candidate === "string" && /^#?[0-9a-fA-F]{6}$/u.test(candidate)) {
    return candidate.replace(/^#/u, "").toUpperCase();
  }
  return invalidBatch(`${path}.${key}`, "expected a six-digit RGB color or null");
};

const readOptionalStringArray = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): string[] | undefined => {
  const candidate = value[key];
  if (candidate === undefined) {
    return undefined;
  }
  if (!Array.isArray(candidate)) {
    return invalidBatch(`${path}.${key}`, "expected an array when provided");
  }
  return candidate.map((item, index) => {
    if (typeof item === "string") {
      return item;
    }
    return invalidBatch(`${path}.${key}[${index}]`, "expected a string");
  });
};

const readNonNegativeInteger = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): number => {
  const candidate = value[key];
  if (typeof candidate === "number" && Number.isSafeInteger(candidate) && candidate >= 0) {
    return candidate;
  }
  return invalidBatch(`${path}.${key}`, "expected a non-negative safe integer");
};

/**
 * A non-negative integer that may also be cleared. `undefined` is "say
 * nothing"; `null` is "there is none", which is a different instruction and
 * the one absence alone cannot give.
 */
const readClearableNonNegativeInteger = (
  value: Record<string, unknown>,
  key: string,
  path: string,
): number | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined) {
    return undefined;
  }
  return candidate === null ? null : readNonNegativeInteger(value, key, path);
};

type ReadClearableParagraphAlignmentParams = {
  value: Record<string, unknown>;
  key: string;
  path: string;
};

const readClearableParagraphAlignment = ({
  value,
  key,
  path,
}: ReadClearableParagraphAlignmentParams): ParagraphAlignment | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined || candidate === null) {
    return candidate;
  }
  for (const alignment of FOLIO_PARAGRAPH_ALIGNMENT_VALUES) {
    if (candidate === alignment) {
      return alignment;
    }
  }
  return invalidBatch(
    `${path}.${key}`,
    `expected one of ${FOLIO_PARAGRAPH_ALIGNMENT_VALUES.join(", ")} or null when provided`,
  );
};

const PARAGRAPH_SPACING_KEYS = [
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
] as const satisfies readonly (keyof FolioAIParagraphSpacing)[];

const isLineSpacingRule = (value: unknown): value is LineSpacingRule =>
  FOLIO_LINE_SPACING_RULE_VALUES.some((rule) => value === rule);

type ReadClearableParagraphSpacingParams = {
  value: Record<string, unknown>;
  key: string;
  path: string;
};

/**
 * Read one complete direct `w:spacing` cluster. Individual fields are
 * optional so explicit zero and false survive; null removes the child.
 */
const readClearableParagraphSpacing = ({
  value,
  key,
  path,
}: ReadClearableParagraphSpacingParams): FolioAIParagraphSpacing | null | undefined => {
  const candidate = value[key];
  if (candidate === undefined || candidate === null) {
    return candidate;
  }
  const spacingPath = `${path}.${key}`;
  if (!isPlainObject(candidate)) {
    return invalidBatch(spacingPath, "expected an object or null when provided");
  }
  assertAllowedKeys(candidate, spacingPath, PARAGRAPH_SPACING_KEYS);
  if (Object.keys(candidate).length === 0) {
    return invalidBatch(spacingPath, "expected at least one spacing property");
  }

  const spacing: FolioAIParagraphSpacing = {};
  for (const unsignedKey of ["spaceBefore", "spaceAfter"] as const) {
    if (candidate[unsignedKey] !== undefined) {
      spacing[unsignedKey] = readNonNegativeInteger(candidate, unsignedKey, spacingPath);
    }
  }
  const lineSpacing = candidate["lineSpacing"];
  if (lineSpacing !== undefined) {
    if (typeof lineSpacing !== "number" || !Number.isSafeInteger(lineSpacing)) {
      return invalidBatch(
        `${spacingPath}.lineSpacing`,
        "expected a safe signed integer when provided",
      );
    }
    spacing.lineSpacing = lineSpacing;
  }
  const lineSpacingRule = candidate["lineSpacingRule"];
  if (lineSpacingRule !== undefined) {
    if (!isLineSpacingRule(lineSpacingRule)) {
      return invalidBatch(
        `${spacingPath}.lineSpacingRule`,
        `expected one of ${FOLIO_LINE_SPACING_RULE_VALUES.join(", ")} when provided`,
      );
    }
    spacing.lineSpacingRule = lineSpacingRule;
  }
  for (const booleanKey of ["beforeAutospacing", "afterAutospacing"] as const) {
    const booleanValue = candidate[booleanKey];
    if (booleanValue === undefined) {
      continue;
    }
    if (typeof booleanValue !== "boolean") {
      return invalidBatch(`${spacingPath}.${booleanKey}`, "expected a boolean when provided");
    }
    spacing[booleanKey] = booleanValue;
  }
  return spacing;
};

/**
 * A rectangular grid of cell texts. Rectangular because a table whose rows
 * hold different cell counts is not a table any consumer can lay out, and the
 * batch is the last place to catch that.
 */
const readTableRows = (
  value: Record<string, unknown>,
  path: string,
): readonly (readonly string[])[] => {
  const rows = value["rows"];
  const rowsPath = `${path}.rows`;
  if (!Array.isArray(rows) || rows.length === 0) {
    return invalidBatch(rowsPath, "expected a non-empty array");
  }
  const parsed = rows.map((row, index) => {
    if (!Array.isArray(row) || row.length === 0) {
      return invalidBatch(`${rowsPath}[${String(index)}]`, "expected a non-empty array");
    }
    return row.map((cell, cellIndex) => {
      if (typeof cell !== "string") {
        return invalidBatch(
          `${rowsPath}[${String(index)}][${String(cellIndex)}]`,
          "expected a string",
        );
      }
      return cell;
    });
  });
  const width = parsed[0]?.length ?? 0;
  if (parsed.some((row) => row.length !== width)) {
    return invalidBatch(rowsPath, "expected every row to hold the same number of cells");
  }
  return parsed;
};

const readParagraphProperties = (
  value: Record<string, unknown>,
  path: string,
): FolioAIBlockParagraphProperties => {
  const candidate = value["properties"];
  const propertiesPath = `${path}.properties`;
  if (!isPlainObject(candidate)) {
    return invalidBatch(propertiesPath, "expected an object");
  }
  assertAllowedKeys(candidate, propertiesPath, ["styleId", "listLevel", "alignment", "spacing"]);
  const rawStyleId = candidate["styleId"];
  const styleId =
    rawStyleId === null ? null : readOptionalString(candidate, "styleId", propertiesPath);
  const listLevel = readClearableNonNegativeInteger(candidate, "listLevel", propertiesPath);
  const alignment = readClearableParagraphAlignment({
    value: candidate,
    key: "alignment",
    path: propertiesPath,
  });
  const spacing = readClearableParagraphSpacing({
    value: candidate,
    key: "spacing",
    path: propertiesPath,
  });
  if (
    styleId === undefined &&
    listLevel === undefined &&
    alignment === undefined &&
    spacing === undefined
  ) {
    return invalidBatch(propertiesPath, "expected at least one property to set");
  }
  return {
    ...(styleId !== undefined && { styleId }),
    ...(listLevel !== undefined && { listLevel }),
    ...(alignment !== undefined && { alignment }),
    ...(spacing !== undefined && { spacing }),
  };
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
): FolioAIInlineFormatting => {
  const candidate = value["formatting"];
  const formattingPath = `${path}.formatting`;
  if (!isPlainObject(candidate)) {
    return invalidBatch(formattingPath, "expected an object");
  }
  assertAllowedKeys(candidate, formattingPath, [
    "bold",
    "italic",
    "underline",
    "strike",
    "fontFamily",
    "fontSizePt",
    "color",
  ]);
  const bold = readOptionalBoolean(candidate, "bold", formattingPath);
  const italic = readOptionalBoolean(candidate, "italic", formattingPath);
  const underline = readOptionalBoolean(candidate, "underline", formattingPath);
  const strike = readOptionalBoolean(candidate, "strike", formattingPath);
  const fontFamily = readClearableNonEmptyString(candidate, "fontFamily", formattingPath);
  const fontSizePt = readClearableFontSize(candidate, "fontSizePt", formattingPath);
  const color = readClearableRgbColor(candidate, "color", formattingPath);
  if (
    bold === undefined &&
    italic === undefined &&
    underline === undefined &&
    strike === undefined &&
    fontFamily === undefined &&
    fontSizePt === undefined &&
    color === undefined
  ) {
    return invalidBatch(formattingPath, "expected at least one formatting property");
  }
  return {
    ...(bold !== undefined && { bold }),
    ...(italic !== undefined && { italic }),
    ...(underline !== undefined && { underline }),
    ...(strike !== undefined && { strike }),
    ...(fontFamily !== undefined && { fontFamily }),
    ...(fontSizePt !== undefined && { fontSizePt }),
    ...(color !== undefined && { color }),
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
  "suggestionId",
] as const;

const RANGE_OPERATION_KEYS = ["id", "type", "range", "severity", "area", "precondition"] as const;

/**
 * Every property each operation type accepts on the wire; the parser
 * rejects any other key. Exported so a lenient front door (an LLM tool
 * decoder) can strip stray keys with the same knowledge instead of
 * guessing, and so tool schemas can be derived from the contract.
 */
export const FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE = Object.freeze({
  replaceInBlock: [...COMMON_OPERATION_KEYS, "find", "replace", "comment"],
  replaceRange: [...RANGE_OPERATION_KEYS, "suggestionId", "replace", "comment"],
  commentOnRange: [...RANGE_OPERATION_KEYS, "comment"],
  formatRange: [...RANGE_OPERATION_KEYS, "suggestionId", "formatting"],
  insertAfterBlock: [
    ...COMMON_OPERATION_KEYS,
    "text",
    "inheritFormatting",
    "alignment",
    "spacing",
    "listLevel",
    "moveId",
    "pageBreakBefore",
    "styleId",
    "comment",
  ],
  insertBeforeBlock: [
    ...COMMON_OPERATION_KEYS,
    "text",
    "inheritFormatting",
    "alignment",
    "spacing",
    "listLevel",
    "moveId",
    "pageBreakBefore",
    "styleId",
    "comment",
  ],
  replaceBlock: [...COMMON_OPERATION_KEYS, "text", "preserveFormatting", "styleId", "comment"],
  deleteBlock: [...COMMON_OPERATION_KEYS, "moveId", "comment"],
  splitBlock: [...COMMON_OPERATION_KEYS, "offset", "separator"],
  mergeBlockWithNext: [...COMMON_OPERATION_KEYS, "separator"],
  setBlockParagraphProperties: [...COMMON_OPERATION_KEYS, "properties"],
  insertTable: [...COMMON_OPERATION_KEYS, "position", "rows"],
  deleteTable: COMMON_OPERATION_KEYS,
  commentOnBlock: [...COMMON_OPERATION_KEYS, "quote", "comment"],
  insertSignatureTable: [...COMMON_OPERATION_KEYS, "position", "parties", "comment"],
  insertTableRow: [...COMMON_OPERATION_KEYS, "position", "cellTexts"],
  deleteTableRow: COMMON_OPERATION_KEYS,
  insertTableColumn: [...COMMON_OPERATION_KEYS, "position", "cellTexts"],
  deleteTableColumn: COMMON_OPERATION_KEYS,
  mergeTableCells: [...COMMON_OPERATION_KEYS, "endBlockId", "rowCount"],
  splitTableCell: COMMON_OPERATION_KEYS,
} as const satisfies Readonly<Record<FolioDocumentOperationType, readonly string[]>>);

const isFolioDocumentOperationType = (value: string): value is FolioDocumentOperationType =>
  Object.hasOwn(FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE, value);

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
    const parsedParty: { name: string; signatory?: string; title?: string } = {
      name: readString(party, "name", partyPath),
    };
    if (signatory !== undefined) {
      parsedParty.signatory = signatory;
    }
    if (title !== undefined) {
      parsedParty.title = title;
    }
    return parsedParty;
  });
};

const parseDocumentOperation = (value: unknown, index: number): FolioDocumentOperation => {
  const path = `$.operations[${index}]`;
  if (!isPlainObject(value)) {
    return invalidBatch(path, "expected an object");
  }

  const id = readString(value, "id", path);
  const type = readString(value, "type", path);
  if (!isFolioDocumentOperationType(type)) {
    return invalidBatch(`${path}.type`, `unsupported operation type "${type}"`);
  }
  assertAllowedKeys(value, path, FOLIO_DOCUMENT_OPERATION_KEYS_BY_TYPE[type]);
  const reviewMeta = readReviewMeta(value, path);
  const precondition = readOptionalPrecondition(value, path);
  const suggestionId = readOptionalString(value, "suggestionId", path);
  const comment = readOptionalComment(value, path);
  const operationMeta = {
    ...reviewMeta,
    ...(precondition !== undefined && { precondition }),
    ...(suggestionId !== undefined && { suggestionId }),
  };

  if (type === "replaceRange") {
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
    if (comment === undefined) {
      return invalidBatch(`${path}.comment`, "expected an object");
    }
    return { ...operationMeta, id, type, range: readTextRange(value, path), comment };
  }

  if (type === "formatRange") {
    return {
      ...operationMeta,
      id,
      type,
      range: readTextRange(value, path),
      formatting: readInlineFormatting(value, path),
    };
  }

  const blockId = readString(value, "blockId", path);

  if (type === "insertTable") {
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
      rows: readTableRows(value, path),
    };
  }

  if (type === "deleteTable") {
    return { ...operationMeta, id, type, blockId };
  }

  if (type === "setBlockParagraphProperties") {
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      properties: readParagraphProperties(value, path),
    };
  }

  if (type === "splitBlock" || type === "mergeBlockWithNext") {
    const separator = readOptionalString(value, "separator", path);
    if (type === "mergeBlockWithNext") {
      return { ...operationMeta, id, type, blockId, ...(separator !== undefined && { separator }) };
    }
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      offset: readNonNegativeInteger(value, "offset", path),
      ...(separator !== undefined && { separator }),
    };
  }

  if (type === "replaceInBlock") {
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
    const inheritFormatting = readOptionalBoolean(value, "inheritFormatting", path);
    const pageBreakBefore = readOptionalBoolean(value, "pageBreakBefore", path);
    const styleId = value["styleId"] === null ? null : readOptionalString(value, "styleId", path);
    const moveId = readOptionalString(value, "moveId", path);
    const listLevel = readClearableNonNegativeInteger(value, "listLevel", path);
    const alignment = readClearableParagraphAlignment({ value, key: "alignment", path });
    const spacing = readClearableParagraphSpacing({ value, key: "spacing", path });
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      text: readString(value, "text", path),
      ...(inheritFormatting !== undefined && { inheritFormatting }),
      ...(alignment !== undefined && { alignment }),
      ...(spacing !== undefined && { spacing }),
      ...(listLevel !== undefined && { listLevel }),
      ...(moveId !== undefined && { moveId }),
      ...(pageBreakBefore !== undefined && { pageBreakBefore }),
      ...(styleId !== undefined && { styleId }),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "replaceBlock") {
    const preserveFormatting = readOptionalBoolean(value, "preserveFormatting", path);
    const styleId = value["styleId"] === null ? null : readOptionalString(value, "styleId", path);
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
    const moveId = readOptionalString(value, "moveId", path);
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      ...(moveId !== undefined && { moveId }),
      ...(comment !== undefined && { comment }),
    };
  }

  if (type === "commentOnBlock") {
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

  if (type === "insertTableRow") {
    const position = value["position"];
    if (position !== undefined && position !== "after" && position !== "before") {
      return invalidBatch(`${path}.position`, 'expected "after" or "before" when provided');
    }
    const cellTexts = readOptionalStringArray(value, "cellTexts", path);
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      ...(position !== undefined && { position }),
      ...(cellTexts !== undefined && { cellTexts }),
    };
  }

  if (type === "deleteTableRow") {
    return { ...operationMeta, id, type, blockId };
  }

  if (type === "insertTableColumn") {
    const position = value["position"];
    if (position !== undefined && position !== "after" && position !== "before") {
      return invalidBatch(`${path}.position`, 'expected "after" or "before" when provided');
    }
    const cellTexts = readOptionalStringArray(value, "cellTexts", path);
    return {
      ...operationMeta,
      id,
      type,
      blockId,
      ...(position !== undefined && { position }),
      ...(cellTexts !== undefined && { cellTexts }),
    };
  }

  if (type === "deleteTableColumn") {
    return { ...operationMeta, id, type, blockId };
  }

  if (type === "mergeTableCells") {
    const endBlockId = readOptionalString(value, "endBlockId", path);
    const rawRowCount = value["rowCount"];
    if ((endBlockId === undefined) === (rawRowCount === undefined)) {
      return invalidBatch(path, "expected exactly one of endBlockId or rowCount");
    }
    if (endBlockId !== undefined) {
      if (endBlockId.length === 0) {
        return invalidBatch(`${path}.endBlockId`, "expected a non-empty string");
      }
      return { ...operationMeta, id, type, blockId, endBlockId };
    }
    const rowCount = readNonNegativeInteger(value, "rowCount", path);
    if (rowCount < 2) {
      return invalidBatch(`${path}.rowCount`, "expected an integer greater than or equal to 2");
    }
    return { ...operationMeta, id, type, blockId, rowCount };
  }

  // splitTableCell: the type guard above admits nothing else.
  return { ...operationMeta, id, type, blockId };
};

const readOptionalBatchPrecondition = (
  value: Record<string, unknown>,
): FolioDocumentOperationBatchPrecondition | undefined => {
  const candidate = value["precondition"];
  if (candidate === undefined) {
    return undefined;
  }
  if (!isPlainObject(candidate)) {
    return invalidBatch("$.precondition", "expected an object when provided");
  }
  assertAllowedKeys(candidate, "$.precondition", FOLIO_DOCUMENT_OPERATION_BATCH_PRECONDITIONS);
  const documentVersion = readString(candidate, "documentVersion", "$.precondition");
  if (documentVersion.length === 0) {
    return invalidBatch("$.precondition.documentVersion", "expected a non-empty string");
  }
  return { documentVersion };
};

export const parseFolioDocumentOperationBatch = (value: unknown): FolioDocumentOperationBatch => {
  if (isParsedFolioDocumentOperationBatch(value)) {
    return value;
  }
  if (!isPlainObject(value)) {
    return invalidBatch("$", "expected an object");
  }
  assertAllowedKeys(value, "$", [
    "version",
    "operations",
    "mode",
    "atomic",
    "dryRun",
    "precondition",
  ]);
  const version = assertSupportedFolioDocumentOperationVersion(value["version"]);
  const operations = value["operations"];
  if (!Array.isArray(operations)) {
    return invalidBatch("$.operations", "expected an array");
  }
  const mode = value["mode"];
  if (
    mode !== undefined &&
    mode !== "direct" &&
    mode !== "tracked-changes" &&
    mode !== "suggested"
  ) {
    return invalidBatch(
      "$.mode",
      'expected "direct", "tracked-changes", or "suggested" when provided',
    );
  }
  const atomic = readOptionalBoolean(value, "atomic", "$");
  const dryRun = readOptionalBoolean(value, "dryRun", "$");
  const precondition = readOptionalBatchPrecondition(value);
  const parsedOperations = operations.map(parseDocumentOperation);
  const operationIds = new Set<string>();
  for (const [index, operation] of parsedOperations.entries()) {
    if (operationIds.has(operation.id)) {
      return invalidBatch(`$.operations[${index}].id`, "expected a unique operation id");
    }
    operationIds.add(operation.id);
  }
  const parsedBatch = {
    version,
    operations: parsedOperations,
    ...(mode !== undefined && { mode }),
    ...(atomic !== undefined && { atomic }),
    ...(dryRun !== undefined && { dryRun }),
    ...(precondition !== undefined && { precondition }),
  } satisfies FolioDocumentOperationBatch;
  freezeParsedValue(parsedBatch);
  parsedFolioDocumentOperationBatches.add(parsedBatch);
  return parsedBatch;
};

/**
 * `queued` is reported by a surface that routes the batch into a host-owned
 * review queue instead of applying it; the core apply path never produces it.
 */
export type FolioDocumentOperationStatus = "committed" | "previewed" | "rejected" | "queued";

export type FolioDocumentOperationRecovery =
  | "refreshDocument"
  | "narrowMatch"
  | "changeMode"
  | "changeTarget"
  | "removeOperation"
  | "inspectBatch"
  | "resolveTrackedChange"
  | "retryLater";

export type FolioDocumentOperationIssue = {
  operationId: string;
  operationIndex: number;
  path: `$.operations[${number}]`;
  code: FolioAIEditSkippedOperation["reason"];
  retryable: boolean;
  recovery: FolioDocumentOperationRecovery;
};

export type FolioDocumentOperationStory =
  | "main"
  | { type: "header"; relationshipId: string }
  | { type: "footer"; relationshipId: string }
  | { type: "footnote"; noteId: number }
  | { type: "endnote"; noteId: number };

/** One typed target affected by a successfully applied document operation. */
export type FolioDocumentOperationAffectedTarget =
  | {
      type: "block";
      story: FolioDocumentOperationStory;
      blockId: string;
      effect: "updated" | "deleted" | "commented";
    }
  | {
      type: "textRange";
      range: FolioAITextRangeHandle;
      effect: "formatted" | "commented";
      story?: Exclude<FolioDocumentOperationStory, "main">;
    }
  | {
      type: "insertion";
      story: FolioDocumentOperationStory;
      anchorBlockId: string;
      position: "before" | "after";
      content: "block" | "signatureTable" | "table" | "tableRow" | "tableColumn";
    }
  | {
      type: "comment";
      commentId: number;
    }
  | {
      type: "tableRow";
      story: FolioDocumentOperationStory;
      anchorBlockId: string;
      effect: "deleted";
    }
  | {
      type: "tableColumn";
      story: FolioDocumentOperationStory;
      anchorBlockId: string;
      effect: "deleted";
    }
  | ({
      type: "tableCells";
      story: FolioDocumentOperationStory;
      anchorBlockId: string;
      effect: "merged";
    } & (
      | { endAnchorBlockId: string; rowCount?: never }
      | { rowCount: number; endAnchorBlockId?: never }
    ))
  | {
      type: "tableCell";
      story: FolioDocumentOperationStory;
      anchorBlockId: string;
      effect: "split";
    };

/** Input-ordered effect receipt for one successfully applied operation. */
export type FolioDocumentOperationReceipt = {
  operationId: string;
  operationIndex: number;
  affected: FolioDocumentOperationAffectedTarget[];
};

/** Opaque handle for undoing one committed document-operation batch. */
export type FolioDocumentOperationUndoHandle = {
  type: "documentOperationUndo";
  id: string;
};

export type FolioDocumentOperationUndoFailureReason =
  | "unknownHandle"
  | "notLatest"
  | "documentChanged";

export type FolioDocumentOperationUndoResult =
  | {
      status: "undone";
      undoHandle: FolioDocumentOperationUndoHandle;
    }
  | {
      status: "rejected";
      undoHandle: FolioDocumentOperationUndoHandle;
      reason: FolioDocumentOperationUndoFailureReason;
    };

/** One operation a host surface accepted into its own review queue rather than applying. */
export type FolioDocumentOperationQueuedOperation = {
  id: string;
};

/**
 * What every operation result carries, whichever status it reports. Exported
 * so the public API report shows these fields rather than a name it cannot
 * resolve.
 */
export type FolioDocumentOperationResultBase = {
  version: typeof FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION;
  applied: FolioAIEditAppliedOperation[];
  skipped: FolioAIEditSkippedOperation[];
  issues: FolioDocumentOperationIssue[];
  /** Successful effects in input-operation order; skipped operations are omitted. */
  receipts: FolioDocumentOperationReceipt[];
  /** Present when at least one operation triggered an automatic input normalization. */
  normalizations?: FolioAIEditNormalization[];
  /** Present when the execution surface can undo this committed batch. */
  undoHandle: FolioDocumentOperationUndoHandle | null;
  /**
   * First revision id a following batch may allocate against this document:
   * one past the last id this batch used, or the seed it started from when it
   * allocated none. A caller writing several batches into one package — one
   * per document story — must seed each from the previous batch's value, or
   * two stories claim the same `w:id`.
   *
   * The in-process applier always reports it. A host bridge that delegates to
   * an editor it does not control omits it rather than guessing a number a
   * caller would seed the next batch from.
   */
  nextRevisionId?: number;
};

/**
 * Discriminated on `status`: only a host review-queue surface reports
 * `"queued"`, and only that branch carries the `queued` list, so a consumer
 * narrowing on the status can never meet a contradictory payload.
 */
export type FolioDocumentOperationResult =
  | (FolioDocumentOperationResultBase & {
      status: "queued";
      /** Operations parked in the host review queue instead of applied. */
      queued: FolioDocumentOperationQueuedOperation[];
    })
  | (FolioDocumentOperationResultBase & {
      status: Exclude<FolioDocumentOperationStatus, "queued">;
      queued?: never;
    });

const recoveryByReason = {
  missingBlock: "refreshDocument",
  changedBlock: "refreshDocument",
  ambiguousFind: "narrowMatch",
  missingFind: "refreshDocument",
  unsupportedBlock: "changeTarget",
  unsupportedMode: "changeMode",
  atomicBatchRejected: "inspectBatch",
  preconditionFailed: "refreshDocument",
  staleRange: "refreshDocument",
  emptyOperation: "removeOperation",
  pendingParagraphPropertyChange: "resolveTrackedChange",
  noopOperation: "removeOperation",
  documentVersionMismatch: "refreshDocument",
  documentNotEditable: "retryLater",
} as const satisfies Record<FolioAIEditSkippedOperation["reason"], FolioDocumentOperationRecovery>;

export const getFolioDocumentOperationIssues = (
  operations: readonly FolioDocumentOperation[],
  skipped: readonly FolioAIEditSkippedOperation[],
): FolioDocumentOperationIssue[] => {
  const indexById = new Map<string, number>();
  operations.forEach(({ id }, index) => {
    indexById.set(id, index);
  });
  return skipped.map(({ id, reason }) => {
    const operationIndex = indexById.get(id) ?? -1;
    return {
      operationId: id,
      operationIndex,
      path: `$.operations[${operationIndex}]`,
      code: reason,
      retryable:
        reason !== "emptyOperation" &&
        reason !== "noopOperation" &&
        reason !== "pendingParagraphPropertyChange",
      recovery: recoveryByReason[reason],
    };
  });
};

const getPrimaryAffectedTarget = (
  operation: FolioDocumentOperation,
  story: FolioDocumentOperationStory,
): FolioDocumentOperationAffectedTarget => {
  switch (operation.type) {
    case "replaceInBlock":
    case "replaceBlock":
      return {
        type: "block",
        story,
        blockId: operation.blockId,
        effect: "updated",
      };
    case "replaceRange":
      return {
        type: "block",
        story,
        blockId: operation.range.blockId,
        effect: "updated",
      };
    case "commentOnRange":
      return {
        type: "textRange",
        range: operation.range,
        effect: "commented",
        ...(story !== "main" && { story }),
      };
    case "formatRange":
      return {
        type: "textRange",
        range: operation.range,
        effect: "formatted",
        ...(story !== "main" && { story }),
      };
    case "insertAfterBlock":
    case "insertBeforeBlock":
      return {
        type: "insertion",
        story,
        anchorBlockId: operation.blockId,
        position: operation.type === "insertBeforeBlock" ? "before" : "after",
        content: "block",
      };
    case "deleteBlock":
      return {
        type: "block",
        story,
        blockId: operation.blockId,
        effect: "deleted",
      };
    case "setBlockParagraphProperties":
    case "splitBlock":
    case "mergeBlockWithNext":
      // The paragraph mark moved, and it belongs to this block; the block the
      // break was taken from or given to is the one a caller navigates to.
      return {
        type: "block",
        story,
        blockId: operation.blockId,
        effect: "updated",
      };
    case "commentOnBlock":
      return {
        type: "block",
        story,
        blockId: operation.blockId,
        effect: "commented",
      };
    case "insertSignatureTable":
      return {
        type: "insertion",
        story,
        anchorBlockId: operation.blockId,
        position: operation.position ?? "after",
        content: "signatureTable",
      };
    case "insertTable":
      return {
        type: "insertion",
        story,
        anchorBlockId: operation.blockId,
        position: operation.position ?? "after",
        content: "table",
      };
    case "deleteTable":
      return {
        type: "block",
        story,
        blockId: operation.blockId,
        effect: "deleted",
      };
    case "insertTableRow":
      return {
        type: "insertion",
        story,
        anchorBlockId: operation.blockId,
        position: operation.position ?? "after",
        content: "tableRow",
      };
    case "deleteTableRow":
      return {
        type: "tableRow",
        story,
        anchorBlockId: operation.blockId,
        effect: "deleted",
      };
    case "insertTableColumn":
      return {
        type: "insertion",
        story,
        anchorBlockId: operation.blockId,
        position: operation.position ?? "after",
        content: "tableColumn",
      };
    case "deleteTableColumn":
      return {
        type: "tableColumn",
        story,
        anchorBlockId: operation.blockId,
        effect: "deleted",
      };
    case "mergeTableCells":
      return operation.rowCount !== undefined
        ? {
            type: "tableCells",
            story,
            anchorBlockId: operation.blockId,
            rowCount: operation.rowCount,
            effect: "merged",
          }
        : {
            type: "tableCells",
            story,
            anchorBlockId: operation.blockId,
            endAnchorBlockId: operation.endBlockId,
            effect: "merged",
          };
    case "splitTableCell":
      return {
        type: "tableCell",
        story,
        anchorBlockId: operation.blockId,
        effect: "split",
      };
  }
};

/** Build deterministic affected-target receipts from operations and their applied entries. */
export const getFolioDocumentOperationReceipts = (
  operations: readonly FolioDocumentOperation[],
  applied: readonly FolioAIEditAppliedOperation[],
): FolioDocumentOperationReceipt[] =>
  getFolioDocumentOperationReceiptsForStory({ operations, applied, story: "main" });

type GetFolioDocumentOperationReceiptsForStoryOptions = {
  operations: readonly FolioDocumentOperation[];
  applied: readonly FolioAIEditAppliedOperation[];
  story: FolioDocumentOperationStory;
};

const getFolioDocumentOperationReceiptsForStory = ({
  operations,
  applied,
  story,
}: GetFolioDocumentOperationReceiptsForStoryOptions): FolioDocumentOperationReceipt[] => {
  const appliedById = new Map<string, FolioAIEditAppliedOperation>();
  applied.forEach((operation) => {
    appliedById.set(operation.id, operation);
  });
  const receipts: FolioDocumentOperationReceipt[] = [];
  operations.forEach((operation, operationIndex) => {
    const appliedOperation = appliedById.get(operation.id);
    if (!appliedOperation) {
      return;
    }
    const affected: FolioDocumentOperationAffectedTarget[] = [
      getPrimaryAffectedTarget(operation, story),
    ];
    if (appliedOperation.commentId !== undefined) {
      affected.push({ type: "comment", commentId: appliedOperation.commentId });
    }
    receipts.push({ operationId: operation.id, operationIndex, affected });
  });
  return receipts;
};

export type ApplyFolioDocumentOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  batch: FolioDocumentOperationBatch;
  story?: FolioDocumentOperationStory;
  author?: string;
  createCommentId?: (text: string) => number;
  createUndoHandle?: () => FolioDocumentOperationUndoHandle;
  /** Omit to stamp revisions from the wall clock and the shared id cursor. */
  revisionStamp?: FolioRevisionStamp;
  /** Token size a replacement's redline is cut at. Word-level by default. */
  wordDiff?: FolioWordDiffOptions;
  /**
   * Tables and rows an `insertTable` / `insertTableRow` in the batch places
   * verbatim, by operation id. Outside the serialized batch because a table
   * node is not JSON: the contract still describes a table by its cell texts,
   * and an in-process caller that already holds the table — a comparison
   * copying the target document's — hands the whole thing over instead.
   */
  tableTemplates?: FolioTableTemplates;
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
  story = "main",
  author,
  createCommentId,
  createUndoHandle,
  revisionStamp,
  wordDiff,
  tableTemplates,
}: ApplyFolioDocumentOperationsOptions): FolioDocumentOperationResult => {
  const parsedBatch = parseFolioDocumentOperationBatch(batch);
  const paragraphMarkFormattingTemplates = paragraphMarkFormattingTemplatesOf(batch);
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
      ...(revisionStamp !== undefined && { revisionStamp }),
      ...(wordDiff !== undefined && { wordDiff }),
      ...(tableTemplates !== undefined && { tableTemplates }),
      ...(paragraphMarkFormattingTemplates !== undefined && {
        paragraphMarkFormattingTemplates,
      }),
    });
  };

  const preview = () => apply({ targetView: view, preview: true });

  const atomicResult = (
    previewResult: FolioAIEditApplyOutcome,
    status: "previewed" | "rejected",
  ): FolioDocumentOperationResult => {
    const skippedById = new Map(
      previewResult.skipped.map((operation) => [operation.id, operation]),
    );
    const skipped = parsedBatch.operations.map(
      ({ id }): FolioAIEditSkippedOperation =>
        skippedById.get(id) ?? { id, reason: "atomicBatchRejected" },
    );
    return {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      status,
      applied: [],
      skipped,
      issues: getFolioDocumentOperationIssues(parsedBatch.operations, skipped),
      receipts: [],
      ...(previewResult.normalizations !== undefined && {
        normalizations: previewResult.normalizations,
      }),
      undoHandle: null,
      nextRevisionId: previewResult.nextRevisionId,
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
      issues: getFolioDocumentOperationIssues(parsedBatch.operations, previewResult.skipped),
      receipts: getFolioDocumentOperationReceiptsForStory({
        operations: parsedBatch.operations,
        applied: previewResult.applied,
        story,
      }),
      ...(previewResult.normalizations !== undefined && {
        normalizations: previewResult.normalizations,
      }),
      undoHandle: null,
      nextRevisionId: previewResult.nextRevisionId,
    };
  }

  if (parsedBatch.atomic === true) {
    const previewResult = preview();
    if (previewResult.skipped.length > 0) {
      return atomicResult(previewResult, "rejected");
    }
  }

  const result = apply({ targetView: view });
  return {
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    status: "committed",
    ...result,
    issues: getFolioDocumentOperationIssues(parsedBatch.operations, result.skipped),
    receipts: getFolioDocumentOperationReceiptsForStory({
      operations: parsedBatch.operations,
      applied: result.applied,
      story,
    }),
    undoHandle:
      result.applied.length > 0 && createUndoHandle !== undefined ? createUndoHandle() : null,
  };
};
