import { cloneModel } from "../modelClone";
import { Result, TaggedError } from "better-result";
import type { ParagraphFormatting, TextFormatting } from "../../model/document";
import { PARAGRAPH_ALIGNMENTS } from "../../model/ooxmlEnumerations.gen";
import { isOpStory } from "./address";
import { captureDocumentOp, restoreDocumentOp } from "../wire";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentOp,
  type DocumentOpType,
  type TableEditOp,
  type SetTableOp,
  type InsertTableOp,
  type DeleteTableOp,
  type InsertRowOp,
  type DeleteRowOp,
  type SetTableRowsOp,
  type SetContainerBlocksOp,
  type SplitHalf,
  type TextPosition,
} from "../types";

/** A batch applies atomically, in array order, against its stated revision. */
export type DocumentBatch = {
  schema: typeof DOCUMENT_OP_SCHEMA_VERSION;
  opId: string;
  /** Opaque actor identity; authentication belongs to the caller. */
  actor: string;
  baseRev: number;
  ops: readonly DocumentOp[];
  revision?: number;
};

/** An operation with no position or structural effect. */
export const NO_SEQUENCED_EFFECT = "none";

/** Document-dependent facts captured when an operation is sequenced. */
export type SequencedOpEffect =
  | { type: "splitBlock"; newHalf: SplitHalf }
  | { type: "joinBlocks"; firstLength: number }
  | { type: "touchedBlocks"; blockIds: readonly string[] }
  | { type: typeof NO_SEQUENCED_EFFECT };

export type SequencedBatch = DocumentBatch & {
  revision: number;
  effects?: readonly SequencedOpEffect[];
};

export const BATCH_REJECTION_REASONS = {
  INVALID_BATCH: "invalidBatch",
  UNSUPPORTED_SCHEMA: "unsupportedSchema",
  UNSUPPORTED_PAIR: "unsupportedPair",
  CONFLICT: "conflict",
  STALE_BASE: "staleBase",
  INVALID_OPERATION: "invalidOperation",
  TABLE_REQUIRES_EXCLUSIVE_EDIT: "tableRequiresExclusiveEdit",
} as const;

export type BatchRejectionReason =
  (typeof BATCH_REJECTION_REASONS)[keyof typeof BATCH_REJECTION_REASONS];

export class BatchRejection extends TaggedError("BatchRejection")<{
  message: string;
  reason: BatchRejectionReason;
  opId?: string;
  opType?: DocumentOpType;
  overType?: DocumentOpType;
}> {}

/** Table edits and their structural inverses require exclusive editing until transforms exist. */
export const TABLE_EXCLUSIVE_OP_TYPES = {
  insertTable: true,
  deleteTable: true,
  setContainerBlocks: true,
  insertRow: true,
  deleteRow: true,
  setTableRows: true,
  insertColumn: true,
  deleteColumn: true,
  mergeCells: true,
  splitCell: true,
  setTableGrid: true,
  setCellProps: true,
  setRowProps: true,
  setTableProps: true,
  setTable: true,
} as const satisfies Record<
  (
    | TableEditOp
    | SetTableOp
    | InsertTableOp
    | DeleteTableOp
    | InsertRowOp
    | DeleteRowOp
    | SetTableRowsOp
    | SetContainerBlocksOp
  )["type"],
  true
>;

type Validator = (value: unknown) => boolean;
type Fields = Readonly<Record<string, Validator>>;
type CommonOperationFields = Pick<DocumentOp, "undefinedFields">;
type OperationFields<Type extends DocumentOpType> = {
  [Key in Exclude<
    keyof Extract<DocumentOp, { type: Type }>,
    keyof CommonOperationFields
  >]-?: Validator;
};

const isRecord = (value: unknown): value is Record<string, unknown> =>
  typeof value === "object" && value !== null && !Array.isArray(value);
const string: Validator = (value) => typeof value === "string";
const nonemptyString: Validator = (value) => typeof value === "string" && value.length > 0;
const boolean: Validator = (value) => typeof value === "boolean";
const number: Validator = (value) => typeof value === "number" && Number.isFinite(value);
const natural: Validator = (value) =>
  typeof value === "number" && Number.isSafeInteger(value) && value >= 0;
const blockId: Validator = (value) => typeof value === "string" && /^[0-9A-F]{8}$/u.test(value);
const absent: Validator = (value) => value === undefined;
const optional =
  (check: Validator): Validator =>
  (value) =>
    value === undefined || check(value);
const nullable =
  (check: Validator): Validator =>
  (value) =>
    value === null || check(value);
const literal =
  (...values: readonly unknown[]): Validator =>
  (value) =>
    values.includes(value);
const array =
  (check: Validator): Validator =>
  (value) =>
    Array.isArray(value) && value.every(check);
const commonOperationFields = {
  undefinedFields: optional(array(array(nonemptyString))),
} satisfies Record<keyof CommonOperationFields, Validator>;
const object =
  (fields: Fields): Validator =>
  (value) =>
    isRecord(value) &&
    Object.keys(value).every((key) => Object.hasOwn(fields, key)) &&
    Object.entries(fields).every(([key, check]) => check(value[key]));
const properties = (fields: Fields, patch = false): Validator =>
  object(
    Object.fromEntries(
      Object.entries(fields).map(([key, check]) => [
        key,
        optional(patch ? nullable(check) : check),
      ]),
    ),
  );

const runPropertyFields = {
  bold: boolean,
  boldCs: boolean,
  italic: boolean,
  italicCs: boolean,
  strike: boolean,
  doubleStrike: boolean,
  smallCaps: boolean,
  allCaps: boolean,
  hidden: boolean,
  noProof: boolean,
  emboss: boolean,
  imprint: boolean,
  outline: boolean,
  shadow: boolean,
  rtl: boolean,
  cs: boolean,
  fontSize: number,
  fontSizeCs: number,
  spacing: number,
  position: number,
  scale: number,
  kerning: number,
  styleId: string,
  vertAlign: literal("baseline", "superscript", "subscript"),
  language: object({ val: optional(string), eastAsia: optional(string), bidi: optional(string) }),
} satisfies Partial<Record<keyof TextFormatting, Validator>>;
const paragraphPropertyFields = {
  alignment: literal(...PARAGRAPH_ALIGNMENTS),
  bidi: boolean,
  kinsoku: boolean,
  overflowPunctuation: boolean,
  spaceBefore: number,
  spaceAfter: number,
  lineSpacing: number,
  lineSpacingRule: literal("auto", "exact", "atLeast"),
  snapToGrid: boolean,
  beforeAutospacing: boolean,
  afterAutospacing: boolean,
  indentLeft: number,
  indentRight: number,
  indentFirstLine: number,
  hangingIndent: boolean,
  keepNext: boolean,
  keepLines: boolean,
  widowControl: boolean,
  pageBreakBefore: boolean,
  contextualSpacing: boolean,
  styleId: string,
} satisfies Partial<Record<keyof ParagraphFormatting, Validator>>;
const positionFields = {
  story: isOpStory,
  blockId,
  offset: natural,
  zeroWidthBefore: optional(natural),
} satisfies Record<keyof TextPosition, Validator>;
const position = object(positionFields);
const revision = optional(
  object({ id: natural, author: string, date: string, initials: optional(string) }),
);
const newIds = optional(
  object({ revision: optional(array(natural)), control: optional(array(natural)) }),
);
const half = optional(literal("first", "second"));
const whenEmpty = optional(literal("omit", "keep"));
const blockAnchor = object({ type: literal("before", "after"), blockId });
const runProps = properties(runPropertyFields);
const runPatch = properties(runPropertyFields, true);
const paragraphProps = properties(paragraphPropertyFields);
const paragraphPatch = properties(paragraphPropertyFields, true);
const runContent: Validator = (value) => {
  if (!isRecord(value)) return false;
  switch (value["type"]) {
    case "text":
      return object({ type: literal("text"), text: string })(value);
    case "tab":
      return object({ type: literal("tab") })(value);
    case "softHyphen":
    case "noBreakHyphen":
    case "renderedPageBreak":
      return object({ type: literal(value["type"]) })(value);
    case "break":
      return object({
        type: literal("break"),
        breakType: optional(literal("page", "column", "textWrapping")),
        clear: optional(literal("none", "left", "right", "all")),
        sourceElement: optional(literal("br")),
      })(value);
    default:
      return false;
  }
};
const inline = object({
  type: literal("run"),
  formatting: optional(runProps),
  content: array(runContent),
});
const slice = object({ content: array(inline), openStart: natural, openEnd: natural });
const paragraph = object({
  type: literal("paragraph"),
  paraId: blockId,
  textId: optional(blockId),
  formatting: optional(paragraphProps),
  content: array(inline),
});

/**
 * The unknown-input decoder accepts these operation kinds with scalar
 * formatting, language properties, plain runs and basic inline atoms.
 * Other embedded model records and inverse preconditions are refused;
 * table operations require exclusive editing in both typed and wire batches.
 */
const operationFields = {
  insertText: {
    type: literal("insertText"),
    at: position,
    text: string,
    runProps: (value) => value === "inherit" || runProps(value),
    newIds,
    revision,
  } satisfies OperationFields<"insertText">,
  deleteRange: {
    type: literal("deleteRange"),
    from: position,
    to: position,
    join: optional(natural),
    expected: optional(slice),
    newIds,
    revision,
  } satisfies OperationFields<"deleteRange">,
  setRunProps: {
    type: literal("setRunProps"),
    from: position,
    to: position,
    patch: runPatch,
    expected: optional(runPatch),
    whenEmpty,
    joinStart: optional(natural),
    joinEnd: optional(natural),
    newIds,
    revision,
    propertyReview: absent,
  } satisfies OperationFields<"setRunProps">,
  setParagraphProps: {
    type: literal("setParagraphProps"),
    story: isOpStory,
    blockId,
    patch: paragraphPatch,
    expected: optional(paragraphPatch),
    whenEmpty,
    revision,
  } satisfies OperationFields<"setParagraphProps">,
  splitBlock: {
    type: literal("splitBlock"),
    at: position,
    newBlockId: blockId,
    newHalf: half,
    newParagraph: absent,
    firstMark: absent,
    firstSectionProperties: absent,
    sectionView: absent,
    newIds,
    revision,
  } satisfies OperationFields<"splitBlock">,
  joinBlocks: {
    type: literal("joinBlocks"),
    story: isOpStory,
    blockId,
    nextBlockId: blockId,
    depth: optional(natural),
    survivor: half,
    expectedRetired: absent,
    expectedSurvivor: absent,
    sectionBoundary: absent,
    sectionView: absent,
    newIds,
    revision,
  } satisfies OperationFields<"joinBlocks">,
  insertBlocks: {
    type: literal("insertBlocks"),
    story: isOpStory,
    at: blockAnchor,
    blocks: array(paragraph),
    newIds,
    revision,
  } satisfies OperationFields<"insertBlocks">,
  deleteBlocks: {
    type: literal("deleteBlocks"),
    story: isOpStory,
    blockIds: array(blockId),
    newIds,
    revision,
  } satisfies OperationFields<"deleteBlocks">,
  resolveRevision: {
    type: literal("resolveRevision"),
    story: isOpStory,
    revisionIds: array(natural),
    decision: literal("accept", "reject"),
  } satisfies OperationFields<"resolveRevision">,
  insertContent: {
    type: literal("insertContent"),
    at: position,
    slice,
    newIds,
    revision,
    seamPolicy: absent,
  } satisfies OperationFields<"insertContent">,
  splitInline: undefined,
  joinInline: undefined,
  replaceBlocks: undefined,
  setParagraphReview: undefined,
  replaceInline: undefined,
  insertTable: undefined,
  deleteTable: undefined,
  setContainerBlocks: undefined,
  insertRow: undefined,
  deleteRow: undefined,
  setTableRows: undefined,
  createNumberingInstance: undefined,
  deleteNumberingInstance: undefined,
  setSectionEndpoint: undefined,
  addNote: undefined,
  createHeaderFooter: undefined,
  removeHeaderFooter: undefined,
  removeNote: undefined,
  restoreStoryParts: undefined,
  setSectionProps: undefined,
  insertColumn: undefined,
  deleteColumn: undefined,
  mergeCells: undefined,
  splitCell: undefined,
  setTableGrid: undefined,
  setCellProps: undefined,
  setRowProps: undefined,
  setTableProps: undefined,
  setTable: undefined,
} satisfies Record<DocumentOpType, Fields | undefined>;

export const BATCH_WIRE_OP_TYPES = Object.freeze(
  Object.entries(operationFields)
    .filter(([, fields]) => fields !== undefined)
    .map(([type]) => type),
);

const isDocumentOp = (value: unknown): value is DocumentOp => {
  if (!isRecord(value) || typeof value["type"] !== "string") return false;
  const entry = Object.entries(operationFields).find(([type]) => type === value["type"]);
  const fields = entry?.at(1);
  return (
    typeof fields === "object" &&
    fields !== null &&
    object({ ...fields, ...commonOperationFields })(value)
  );
};

/** Bound recursion and refuse values except JSON data and absent object fields. */
export const MAX_BATCH_WIRE_BYTES = 256 * 1024;

type JsonDataOptions = {
  value: unknown;
  ancestors: Set<object>;
  depth: number;
  budget: { used: number };
};
const isJsonData = ({ value, ancestors, depth, budget }: JsonDataOptions): boolean => {
  budget.used += typeof value === "string" ? value.length + 2 : 1;
  if (depth > 64 || budget.used > MAX_BATCH_WIRE_BYTES) return false;
  if (value === null || typeof value === "string" || typeof value === "boolean") return true;
  if (typeof value === "number") return Number.isFinite(value) && !Object.is(value, -0);
  if (typeof value !== "object") return false;
  if (ancestors.has(value)) return false;
  const prototype: unknown = Object.getPrototypeOf(value);
  if (prototype !== Object.prototype && prototype !== Array.prototype && prototype !== null)
    return false;
  if (Object.getOwnPropertySymbols(value).length > 0) return false;
  ancestors.add(value);
  const entries = Object.entries(Object.getOwnPropertyDescriptors(value));
  const valid =
    entries.every(([key, descriptor]) => {
      if (Array.isArray(value) && key === "length") return true;
      if (!descriptor.enumerable || !Object.hasOwn(descriptor, "value")) return false;
      if (key === "__proto__" || key === "constructor" || key === "prototype") return false;
      const child: unknown = descriptor.value;
      if (child === undefined && !Array.isArray(value)) return true;
      budget.used += key.length;
      return isJsonData({ value: child, ancestors, depth: depth + 1, budget });
    }) &&
    (!Array.isArray(value) || Object.keys(value).length === value.length);
  ancestors.delete(value);
  return valid;
};

const isDocumentBatch = (value: unknown): value is DocumentBatch =>
  object({
    schema: literal(DOCUMENT_OP_SCHEMA_VERSION),
    opId: nonemptyString,
    actor: nonemptyString,
    baseRev: natural,
    ops: array(isDocumentOp),
    revision: optional(natural),
    effects: optional(
      array((effect) => {
        if (!isRecord(effect)) return false;
        switch (effect["type"]) {
          case NO_SEQUENCED_EFFECT:
            return object({ type: literal(NO_SEQUENCED_EFFECT) })(effect);
          case "splitBlock":
            return object({ type: literal("splitBlock"), newHalf: literal("first", "second") })(
              effect,
            );
          case "joinBlocks":
            return object({ type: literal("joinBlocks"), firstLength: natural })(effect);
          case "touchedBlocks":
            return object({ type: literal("touchedBlocks"), blockIds: array(blockId) })(effect);
          default:
            return false;
        }
      }),
    ),
  })(value);

/** Decode a JSON-shaped batch, refusing unsupported payloads before apply. */
export const validateDocumentBatch = (value: unknown): Result<DocumentBatch, BatchRejection> => {
  const jsonData = Result.try(() =>
    isJsonData({ value, ancestors: new Set(), depth: 0, budget: { used: 0 } }),
  );
  if (jsonData.isErr() || !jsonData.value || !isRecord(value)) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Batch must contain plain JSON data",
      }),
    );
  }
  const encoded = Result.try(() => JSON.stringify(value));
  if (
    encoded.isErr() ||
    new TextEncoder().encode(encoded.value).byteLength > MAX_BATCH_WIRE_BYTES
  ) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Batch exceeds its JSON wire size limit",
      }),
    );
  }
  // The decoder also admits typed in-process batches. JSON normalization here
  // would erase their own undefined fields before the wire boundary captures them.
  const normalized = Result.try((): unknown => cloneModel(value));
  if (normalized.isErr() || !isRecord(normalized.value)) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Batch must be a JSON object",
      }),
    );
  }
  const decoded = decodeDocumentBatch(normalized.value);
  if (decoded.isErr()) return decoded;
  if (new TextEncoder().encode(JSON.stringify(decoded.value)).byteLength > MAX_BATCH_WIRE_BYTES) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Captured batch exceeds its JSON wire size limit",
      }),
    );
  }
  return decoded;
};

const decodeDocumentBatch = (
  value: Record<string, unknown>,
): Result<DocumentBatch, BatchRejection> => {
  if (value["schema"] !== DOCUMENT_OP_SCHEMA_VERSION) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.UNSUPPORTED_SCHEMA,
        message: "Unsupported document operation schema",
      }),
    );
  }
  if (
    Array.isArray(value["ops"]) &&
    value["ops"].some(
      (op: unknown) =>
        isRecord(op) &&
        typeof op["type"] === "string" &&
        Object.hasOwn(TABLE_EXCLUSIVE_OP_TYPES, op["type"]),
    )
  ) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.TABLE_REQUIRES_EXCLUSIVE_EDIT,
        message:
          "Table operations require exclusive editing until sequencing transforms are available",
      }),
    );
  }
  if (!Array.isArray(value["ops"]) || !value["ops"].every(isDocumentOp)) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
        message: "Batch contains an invalid or unsupported wire operation",
      }),
    );
  }
  if (!isDocumentBatch(value)) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Invalid batch identity or revision",
      }),
    );
  }
  if ("effects" in value) {
    const effects: unknown = value.effects;
    if (
      !natural(value.revision) ||
      !Array.isArray(effects) ||
      effects.length !== value.ops.length ||
      !effects.every((effect: unknown, index: number) => {
        if (!isRecord(effect)) return false;
        const op = value.ops.at(index);
        if (effect["type"] === NO_SEQUENCED_EFFECT)
          return (
            op?.type !== "splitBlock" &&
            op?.type !== "joinBlocks" &&
            op?.type !== "deleteBlocks" &&
            op?.type !== "resolveRevision"
          );
        if (effect["type"] === "touchedBlocks")
          return op?.type === "deleteBlocks" || op?.type === "resolveRevision";
        return effect["type"] === op?.type;
      })
    ) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
          message: "Sequenced effects must match their operation indexes",
        }),
      );
    }
  }
  const ops: DocumentOp[] = [];
  for (const op of value.ops) {
    const restored = restoreDocumentOp(op);
    if (restored.isErr() || !isDocumentOp(restored.value)) {
      return Result.err(
        new BatchRejection({
          reason: BATCH_REJECTION_REASONS.INVALID_OPERATION,
          message: restored.isErr()
            ? restored.error.message
            : "Restored operation is outside the supported wire payload",
          opType: op.type,
        }),
      );
    }
    ops.push(captureDocumentOp(restored.value));
  }
  return Result.ok({ ...value, ops });
};

const isSequencedBatch = (value: DocumentBatch): value is SequencedBatch =>
  typeof value.revision === "number";

/** Validate an assigned revision and any document-dependent effect metadata. */
export const validateSequencedBatch = (value: unknown): Result<SequencedBatch, BatchRejection> => {
  const batch = validateDocumentBatch(value);
  if (batch.isErr()) return batch;
  if (!isSequencedBatch(batch.value)) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Sequenced batch requires a revision",
      }),
    );
  }
  return Result.ok(batch.value);
};

/** Parse the wire representation without propagating parser exceptions. */
export const parseDocumentBatch = (json: string): Result<DocumentBatch, BatchRejection> => {
  if (new TextEncoder().encode(json).byteLength > MAX_BATCH_WIRE_BYTES) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Batch exceeds its JSON wire size limit",
      }),
    );
  }
  const parsed = Result.try((): unknown => JSON.parse(json));
  if (parsed.isErr()) {
    return Result.err(
      new BatchRejection({
        reason: BATCH_REJECTION_REASONS.INVALID_BATCH,
        message: "Batch is not valid JSON",
      }),
    );
  }
  return validateDocumentBatch(parsed.value);
};
