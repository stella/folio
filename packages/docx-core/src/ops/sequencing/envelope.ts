import { Result, TaggedError, type TaggedErrorClass } from "better-result";

import { encodeUtf8 } from "@stll/utf8-boundaries";

import { INSERTION_SEAM_POLICIES } from "../../model/content";
import type { IdLeaseKind } from "../idLease";
import { cloneModel } from "../modelClone";
import {
  type CommentState,
  type DeleteRowOp,
  type DeleteTableOp,
  DOCUMENT_OP_SCHEMA_VERSION,
  type DocumentOp,
  type DocumentOpType,
  type InsertRowOp,
  type InsertTableOp,
  type SetContainerBlocksOp,
  type SetTableOp,
  type SetTableRowsOp,
  type SplitHalf,
  type StoryParts,
  type TableEditOp,
  type TextPosition,
} from "../types";
import { captureDocumentOp, restoreDocumentOp } from "../wire";
import { headerFooterStory, isOpStory, noteStory } from "./address";
import {
  abstractNumbering,
  array,
  blockContent,
  blockId,
  boolean,
  boundedJson,
  comment,
  documentSettings,
  type Fields,
  headerFooter,
  inlineContent,
  isRecord,
  literal,
  mapEntries,
  natural,
  newComment,
  nonemptyString,
  note,
  nullable,
  number,
  numberingDefinitions,
  numberingInstance,
  object,
  optional,
  packageResourceMedia,
  paragraph,
  paragraphMarkChange,
  paragraphPropertyFields,
  paragraphReviewFields,
  presence,
  properties,
  relationship,
  runPropertyFields,
  sectionProperties,
  sectionPropertyFields,
  sectionViewChange,
  sectionViewState,
  splitParagraphFields,
  string,
  styleDefinitions,
  table,
  tableCellFormattingFields,
  tableFormattingFields,
  tableRow,
  tableRowFormattingFields,
  trackedChangeInfo,
  type Validator,
} from "./wireRecords";

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
  | {
      type: "insertContent";
      /**
       * Whether the insertion cut a record carrying an id (a revision or a
       * content control) in two, its second half taking a new id.
       */
      cutIdentified: boolean;
    }
  | {
      type: "touchedBlocks";
      /** Paragraphs the operation changed, inserted or removed. */
      blockIds: readonly string[];
      /**
       * For a table structure edit, every paragraph of the table before and
       * after it: another structure edit of that table conflicts with it.
       */
      tableBlockIds?: readonly string[];
      /**
       * For a comment operation, the comments it added, removed or changed,
       * replies of a deleted thread among them, and whether it removed a
       * package relationship of the comment parts.
       */
      comments?: { ids: readonly number[]; relationshipRemoved: boolean };
    }
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
  /** A concurrent operation restructured the table this one restructures. */
  TABLE_REQUIRES_EXCLUSIVE_EDIT: "tableRequiresExclusiveEdit",
  /** A concurrent operation changed a paragraph this one addresses. */
  BLOCK_CONFLICT: "blockConflict",
  /** A concurrent operation wrote a package part this one uses (`part`). */
  PART_CONFLICT: "partConflict",
  /** The batch introduces an id its actor does not lease. */
  OUTSIDE_ID_LEASE: "outsideIdLease",
} as const;

export type BatchRejectionReason =
  (typeof BATCH_REJECTION_REASONS)[keyof typeof BATCH_REJECTION_REASONS];

/**
 * Package parts an operation reads or writes besides the paragraphs it
 * addresses. Concurrent operations conflict on a part when one writes it and
 * both name the same entry, or one names the whole part.
 */
export const EXCLUSIVE_PARTS: Readonly<{
  readonly NUMBERING: "numbering";
  readonly STYLES: "styles";
  readonly RELATIONSHIPS: "relationships";
  readonly MEDIA: "media";
  readonly SECTIONS: "sections";
  readonly HEADERS_FOOTERS: "headersFooters";
  readonly NOTES: "notes";
  readonly SETTINGS: "settings";
  readonly REVISIONS: "revisions";
  readonly COMMENTS: "comments";
}> = Object.freeze({
  NUMBERING: "numbering",
  STYLES: "styles",
  RELATIONSHIPS: "relationships",
  MEDIA: "media",
  SECTIONS: "sections",
  HEADERS_FOOTERS: "headersFooters",
  NOTES: "notes",
  SETTINGS: "settings",
  REVISIONS: "revisions",
  COMMENTS: "comments",
} as const);

/** One of {@link EXCLUSIVE_PARTS}. */
export type ExclusivePart =
  (typeof EXCLUSIVE_PARTS)[keyof typeof EXCLUSIVE_PARTS];

const BatchRejectionBase: TaggedErrorClass<"BatchRejection"> =
  TaggedError("BatchRejection");

export class BatchRejection extends BatchRejectionBase<{
  message: string;
  reason: BatchRejectionReason;
  opId?: string;
  opType?: DocumentOpType;
  overType?: DocumentOpType;
  /** The kind of the id an out-of-lease batch introduces. */
  idKind?: IdLeaseKind;
  /** The package part a part conflict names. */
  part?: ExclusivePart;
}> {}

/**
 * Table structure edits and their structural inverses. Two of them that
 * restructure the same table concurrently conflict; one of them passes
 * concurrent edits of other tables and of paragraphs it does not address.
 */
export const TABLE_EXCLUSIVE_OP_TYPES: {
  readonly insertTable: true;
  readonly deleteTable: true;
  readonly setContainerBlocks: true;
  readonly insertRow: true;
  readonly deleteRow: true;
  readonly setTableRows: true;
  readonly insertColumn: true;
  readonly deleteColumn: true;
  readonly mergeCells: true;
  readonly splitCell: true;
  readonly setTableGrid: true;
  readonly setCellProps: true;
  readonly setRowProps: true;
  readonly setTableProps: true;
  readonly setTable: true;
} = {
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

/**
 * The effect the sequencer records for each kind: none for an operation the
 * position transform rebases by its own fields, the split half or join length
 * a position needs, or the paragraphs a block-scoped operation or a comment's
 * anchors touched.
 */
export const SEQUENCED_EFFECT_TYPES: Readonly<
  Record<DocumentOpType, SequencedOpEffect["type"]>
> = Object.freeze({
  insertText: NO_SEQUENCED_EFFECT,
  deleteRange: NO_SEQUENCED_EFFECT,
  setRunProps: NO_SEQUENCED_EFFECT,
  setParagraphProps: NO_SEQUENCED_EFFECT,
  insertBlocks: NO_SEQUENCED_EFFECT,
  splitBlock: "splitBlock",
  joinBlocks: "joinBlocks",
  resolveRevision: "touchedBlocks",
  deleteBlocks: "touchedBlocks",
  insertContent: "insertContent",
  splitInline: NO_SEQUENCED_EFFECT,
  joinInline: NO_SEQUENCED_EFFECT,
  replaceBlocks: "touchedBlocks",
  setParagraphReview: "touchedBlocks",
  setContentControl: "touchedBlocks",
  replaceInline: "touchedBlocks",
  setRevisionInfo: "touchedBlocks",
  insertTable: "touchedBlocks",
  deleteTable: "touchedBlocks",
  setContainerBlocks: "touchedBlocks",
  insertRow: "touchedBlocks",
  deleteRow: "touchedBlocks",
  setTableRows: "touchedBlocks",
  insertColumn: "touchedBlocks",
  deleteColumn: "touchedBlocks",
  mergeCells: "touchedBlocks",
  splitCell: "touchedBlocks",
  setTableGrid: "touchedBlocks",
  setCellProps: "touchedBlocks",
  setRowProps: "touchedBlocks",
  setTableProps: "touchedBlocks",
  setTable: "touchedBlocks",
  createNumberingInstance: "touchedBlocks",
  deleteNumberingInstance: "touchedBlocks",
  setSectionEndpoint: "touchedBlocks",
  setSectionProps: "touchedBlocks",
  setPackageResources: "touchedBlocks",
  createHeaderFooter: "touchedBlocks",
  removeHeaderFooter: "touchedBlocks",
  restoreStoryParts: "touchedBlocks",
  addNote: "touchedBlocks",
  removeNote: "touchedBlocks",
  createComment: "touchedBlocks",
  updateCommentContent: "touchedBlocks",
  setCommentResolution: "touchedBlocks",
  deleteComment: "touchedBlocks",
  restoreCommentState: "touchedBlocks",
});

/** The effect type a sequenced operation of this kind must record. */
export const sequencedEffectType = (
  type: DocumentOpType,
): SequencedOpEffect["type"] | undefined =>
  Object.entries(SEQUENCED_EFFECT_TYPES).find(([kind]) => kind === type)?.[1];

type CommonOperationFields = Pick<DocumentOp, "undefinedFields">;
type OperationFields<Type extends DocumentOpType> = {
  [
    Key in Exclude<
      keyof Extract<DocumentOp, { type: Type }>,
      keyof CommonOperationFields
    >
  ]-?: Validator;
};

const commonOperationFields = {
  undefinedFields: optional(array(array(nonemptyString))),
} satisfies Record<keyof CommonOperationFields, Validator>;
const positionFields = {
  story: isOpStory,
  blockId,
  offset: natural,
  zeroWidthBefore: optional(natural),
} satisfies Record<keyof TextPosition, Validator>;
const position = object(positionFields);
const revision = optional(
  object({
    id: natural,
    author: string,
    date: string,
    initials: optional(string),
    suggestionId: optional(string),
  }),
);
const idSet = object({
  revision: optional(array(natural)),
  control: optional(array(natural)),
});
const newIds = optional(idSet);
const half = optional(literal("first", "second"));
const whenEmpty = optional(literal("omit", "keep"));
const blockAnchor = object({ type: literal("before", "after"), blockId });
const runProps = properties(runPropertyFields);
const runPatch = properties(runPropertyFields, true);
const paragraphPatch = properties(paragraphPropertyFields, true);
const slice = object({
  content: array(inlineContent),
  openStart: natural,
  openEnd: natural,
});
const headerFooterType = literal("default", "first", "even");
const tableTarget = { story: isOpStory, blockId, revision, newIds };
const numberingPartState = presence("definitions", numberingDefinitions);
const resourceEntry =
  (value: Validator): Validator =>
  (entry) =>
    object({ type: literal("absent") })(entry) ||
    object({ type: literal("present"), value })(entry);
const resourceMapChange = (value: Validator): Validator =>
  object({
    expected: literal("omitted", "undefined", "present"),
    next: literal("omitted", "undefined", "present"),
    entries: array(
      object({
        key: nonemptyString,
        expected: resourceEntry(value),
        next: resourceEntry(value),
      }),
    ),
  });
const packageResources = object({
  styles: presence("present", styleDefinitions),
  numbering: presence("present", numberingDefinitions),
});
const sectionMap = nullable(mapEntries(headerFooterType, headerFooter));
const partMap = nullable(mapEntries(nonemptyString, headerFooter));
const storyParts: Validator = object({
  undefinedFields: optional(
    array(
      object({
        target: literal("body", "package", "section"),
        index: optional(natural),
        keys: array(nonemptyString),
      }),
    ),
  ),
  body: optional(
    object({
      background: boundedJson,
      content: optional(array(blockContent)),
      finalSectionProperties: optional(nullable(sectionProperties)),
      comments: boundedJson,
    }),
  ),
  sections: optional(
    array(
      object({
        index: natural,
        properties: optional(sectionProperties),
        headers: optional(sectionMap),
        footers: optional(sectionMap),
      }),
    ),
  ),
  headers: optional(partMap),
  footers: optional(partMap),
  footnotes: optional(nullable(array(note))),
  endnotes: optional(nullable(array(note))),
  settings: optional(nullable(documentSettings)),
} satisfies Record<keyof StoryParts, Validator>);

const commentAnchor: Validator = (value) =>
  object({ kind: literal("range"), from: position, to: position })(value) ||
  object({ kind: literal("point"), at: position })(value) ||
  object({ kind: literal("reply"), parentId: natural })(value) ||
  object({
    kind: literal("revision"),
    story: isOpStory,
    revisionId: natural,
  })(value);
const commentState: Validator = object({
  relationshipPresence: literal("absent", "undefined", "present"),
  relationships: array(
    object({ index: natural, key: nonemptyString, relationship }),
  ),
  listPresence: literal("absent", "undefined", "present"),
  records: array(object({ index: natural, comment })),
  anchors: array(
    object({ story: isOpStory, blockId, content: array(inlineContent) }),
  ),
} satisfies Record<keyof CommentState, Validator>);

/**
 * The unknown-input decoder accepts every operation kind. Embedded model
 * records are checked by `wireRecords.ts`: their kinds, their declared
 * fields, paragraph ids and scalar formatting values, within the decoder's
 * JSON bounds. A comment's author and date are the operation's own: no
 * receiver stamps them.
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
    propertyReview: optional(literal("append")),
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
    newParagraph: optional(splitParagraphFields),
    firstMark: optional(paragraphMarkChange),
    firstSectionProperties: optional(sectionProperties),
    sectionView: optional(sectionViewChange),
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
    expectedRetired: optional(splitParagraphFields),
    expectedSurvivor: optional(paragraphReviewFields),
    sectionBoundary: optional(literal("remove")),
    sectionView: optional(sectionViewChange),
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
    move: optional(literal(true)),
  } satisfies OperationFields<"insertBlocks">,
  deleteBlocks: {
    type: literal("deleteBlocks"),
    story: isOpStory,
    blockIds: array(blockId),
    newIds,
    revision,
    move: optional(literal(true)),
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
    seamPolicy: optional(literal(INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS)),
  } satisfies OperationFields<"insertContent">,
  splitInline: {
    type: literal("splitInline"),
    at: position,
    depth: natural,
    newIds,
  } satisfies OperationFields<"splitInline">,
  joinInline: {
    type: literal("joinInline"),
    at: position,
    depth: natural,
    ifAlike: optional(literal(true)),
  } satisfies OperationFields<"joinInline">,
  replaceBlocks: {
    type: literal("replaceBlocks"),
    story: isOpStory,
    expected: array(paragraph),
    blocks: array(paragraph),
    sectionBoundaries: optional(literal("replace")),
    sectionView: optional(sectionViewChange),
  } satisfies OperationFields<"replaceBlocks">,
  setParagraphReview: {
    type: literal("setParagraphReview"),
    story: isOpStory,
    blockId,
    expected: paragraphReviewFields,
    review: paragraphReviewFields,
  } satisfies OperationFields<"setParagraphReview">,
  replaceInline: {
    type: literal("replaceInline"),
    story: isOpStory,
    blockId,
    expected: array(inlineContent),
    content: array(inlineContent),
  } satisfies OperationFields<"replaceInline">,
  setRevisionInfo: {
    type: literal("setRevisionInfo"),
    story: isOpStory,
    infos: array(trackedChangeInfo),
  } satisfies OperationFields<"setRevisionInfo">,
  insertTable: {
    type: literal("insertTable"),
    story: isOpStory,
    at: blockAnchor,
    table,
    revision,
    newIds,
    terminal: optional(object({ beforeBlockId: blockId })),
  } satisfies OperationFields<"insertTable">,
  deleteTable: {
    type: literal("deleteTable"),
    ...tableTarget,
    expected: optional(table),
  } satisfies OperationFields<"deleteTable">,
  setContainerBlocks: {
    type: literal("setContainerBlocks"),
    story: isOpStory,
    blockId,
    expected: array(blockContent),
    blocks: array(blockContent),
  } satisfies OperationFields<"setContainerBlocks">,
  insertRow: {
    type: literal("insertRow"),
    ...tableTarget,
    at: natural,
    row: tableRow,
  } satisfies OperationFields<"insertRow">,
  deleteRow: {
    type: literal("deleteRow"),
    ...tableTarget,
    expected: optional(tableRow),
  } satisfies OperationFields<"deleteRow">,
  setTableRows: {
    type: literal("setTableRows"),
    story: isOpStory,
    blockId,
    expected: array(tableRow),
    rows: array(tableRow),
  } satisfies OperationFields<"setTableRows">,
  insertColumn: {
    type: literal("insertColumn"),
    ...tableTarget,
    column: natural,
    width: number,
    newBlockIds: array(blockId),
  } satisfies OperationFields<"insertColumn">,
  deleteColumn: {
    type: literal("deleteColumn"),
    ...tableTarget,
    column: natural,
  } satisfies OperationFields<"deleteColumn">,
  mergeCells: {
    type: literal("mergeCells"),
    ...tableTarget,
    top: natural,
    bottom: natural,
    left: natural,
    right: natural,
    newBlockIds: array(blockId),
  } satisfies OperationFields<"mergeCells">,
  splitCell: {
    type: literal("splitCell"),
    ...tableTarget,
    newBlockIds: array(blockId),
  } satisfies OperationFields<"splitCell">,
  setTableGrid: {
    type: literal("setTableGrid"),
    ...tableTarget,
    columnWidths: array(number),
  } satisfies OperationFields<"setTableGrid">,
  setCellProps: {
    type: literal("setCellProps"),
    ...tableTarget,
    patch: properties(tableCellFormattingFields, true),
  } satisfies OperationFields<"setCellProps">,
  setRowProps: {
    type: literal("setRowProps"),
    ...tableTarget,
    patch: properties(tableRowFormattingFields, true),
  } satisfies OperationFields<"setRowProps">,
  setTableProps: {
    type: literal("setTableProps"),
    ...tableTarget,
    patch: properties(tableFormattingFields, true),
  } satisfies OperationFields<"setTableProps">,
  setTable: {
    type: literal("setTable"),
    story: isOpStory,
    blockId,
    expected: table,
    table,
  } satisfies OperationFields<"setTable">,
  createNumberingInstance: {
    type: literal("createNumberingInstance"),
    num: numberingInstance,
    abstractNum: optional(abstractNumbering),
    expected: optional(numberingPartState),
    restore: optional(numberingPartState),
  } satisfies OperationFields<"createNumberingInstance">,
  deleteNumberingInstance: {
    type: literal("deleteNumberingInstance"),
    num: numberingInstance,
    abstractNum: optional(abstractNumbering),
    expected: optional(numberingPartState),
    restore: optional(numberingPartState),
  } satisfies OperationFields<"deleteNumberingInstance">,
  setSectionEndpoint: {
    type: literal("setSectionEndpoint"),
    endpoint: (value) =>
      object({ type: literal("paragraph"), blockId })(value) ||
      object({ type: literal("final") })(value),
    expected: optional(presence("present", sectionProperties)),
    properties: presence("present", sectionProperties),
    expectedSectionMetadata: optional(sectionViewState),
    sectionMetadata: optional(sectionViewState),
  } satisfies OperationFields<"setSectionEndpoint">,
  setSectionProps: {
    type: literal("setSectionProps"),
    sectionIndex: natural,
    patch: properties(sectionPropertyFields, true),
  } satisfies OperationFields<"setSectionProps">,
  setContentControl: {
    type: literal("setContentControl"),
    story: isOpStory,
    controlId: natural,
    blockIds: array(blockId),
    expected: isRecord,
    properties: isRecord,
  } satisfies OperationFields<"setContentControl">,
  setPackageResources: {
    type: literal("setPackageResources"),
    expected: packageResources,
    resources: packageResources,
    relationships: resourceMapChange(relationship),
    media: resourceMapChange(packageResourceMedia),
  } satisfies OperationFields<"setPackageResources">,
  addNote: {
    type: literal("addNote"),
    at: position,
    note,
    revision,
    newIds,
  } satisfies OperationFields<"addNote">,
  removeNote: {
    type: literal("removeNote"),
    at: position,
    story: noteStory,
    revision,
    newIds,
  } satisfies OperationFields<"removeNote">,
  createHeaderFooter: {
    type: literal("createHeaderFooter"),
    sectionIndex: natural,
    story: headerFooterStory,
    referenceType: headerFooterType,
    content: array(blockContent),
  } satisfies OperationFields<"createHeaderFooter">,
  removeHeaderFooter: {
    type: literal("removeHeaderFooter"),
    sectionIndex: natural,
    story: headerFooterStory,
    referenceType: headerFooterType,
  } satisfies OperationFields<"removeHeaderFooter">,
  restoreStoryParts: {
    type: literal("restoreStoryParts"),
    expected: storyParts,
    parts: storyParts,
  } satisfies OperationFields<"restoreStoryParts">,
  createComment: {
    type: literal("createComment"),
    comment: newComment,
    anchor: commentAnchor,
    newIds,
  } satisfies OperationFields<"createComment">,
  updateCommentContent: {
    type: literal("updateCommentContent"),
    id: natural,
    content: array(paragraph),
    patch: optional(
      properties(
        {
          author: string,
          initials: string,
          date: string,
          suggestionId: string,
        },
        true,
      ),
    ),
  } satisfies OperationFields<"updateCommentContent">,
  setCommentResolution: {
    type: literal("setCommentResolution"),
    id: natural,
    status: literal("open", "resolved"),
  } satisfies OperationFields<"setCommentResolution">,
  deleteComment: {
    type: literal("deleteComment"),
    id: natural,
    scope: literal("thread", "reply"),
  } satisfies OperationFields<"deleteComment">,
  restoreCommentState: {
    type: literal("restoreCommentState"),
    ids: array(natural),
    scaffoldIds: idSet,
    expected: commentState,
    state: commentState,
  } satisfies OperationFields<"restoreCommentState">,
} satisfies Record<DocumentOpType, Fields>;

export const BATCH_WIRE_OP_TYPES: readonly string[] = Object.freeze(
  Object.keys(operationFields),
);

const isDocumentOp = (value: unknown): value is DocumentOp => {
  if (!isRecord(value) || typeof value["type"] !== "string") return false;
  const entry = Object.entries(operationFields).find(
    ([type]) => type === value["type"],
  );
  const fields = entry?.at(1);
  return (
    typeof fields === "object" &&
    object({ ...fields, ...commonOperationFields })(value)
  );
};

/** Bound recursion and refuse values except JSON data and absent object fields. */
export const MAX_BATCH_WIRE_BYTES: number = 256 * 1024;

type JsonDataOptions = {
  value: unknown;
  ancestors: Set<object>;
  depth: number;
  budget: { used: number };
};
const isJsonData = ({
  value,
  ancestors,
  depth,
  budget,
}: JsonDataOptions): boolean => {
  budget.used += typeof value === "string" ? value.length + 2 : 1;
  if (depth > 64 || budget.used > MAX_BATCH_WIRE_BYTES) return false;
  if (value === null || typeof value === "string" || typeof value === "boolean")
    return true;
  if (typeof value === "number")
    return Number.isFinite(value) && !Object.is(value, -0);
  if (typeof value !== "object") return false;
  if (ancestors.has(value)) return false;
  const prototype: unknown = Object.getPrototypeOf(value);
  if (
    prototype !== Object.prototype &&
    prototype !== Array.prototype &&
    prototype !== null
  )
    return false;
  if (Object.getOwnPropertySymbols(value).length > 0) return false;
  ancestors.add(value);
  const entries = Object.entries(Object.getOwnPropertyDescriptors(value));
  const valid =
    entries.every(([key, descriptor]) => {
      if (Array.isArray(value) && key === "length") return true;
      if (!descriptor.enumerable || !Object.hasOwn(descriptor, "value"))
        return false;
      if (key === "__proto__" || key === "constructor" || key === "prototype")
        return false;
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
            return object({
              type: literal("splitBlock"),
              newHalf: literal("first", "second"),
            })(effect);
          case "joinBlocks":
            return object({
              type: literal("joinBlocks"),
              firstLength: natural,
            })(effect);
          case "insertContent":
            return object({
              type: literal("insertContent"),
              cutIdentified: boolean,
            })(effect);
          case "touchedBlocks":
            return object({
              type: literal("touchedBlocks"),
              blockIds: array(blockId),
              tableBlockIds: optional(array(blockId)),
              comments: optional(
                object({ ids: array(natural), relationshipRemoved: boolean }),
              ),
            })(effect);
          default:
            return false;
        }
      }),
    ),
  })(value);

/** Decode a JSON-shaped batch, refusing unsupported payloads before apply. */
export const validateDocumentBatch = (
  value: unknown,
): Result<DocumentBatch, BatchRejection> => {
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
    encodeUtf8(encoded.value).byteLength > MAX_BATCH_WIRE_BYTES
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
  if (
    encodeUtf8(JSON.stringify(decoded.value)).byteLength > MAX_BATCH_WIRE_BYTES
  ) {
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
        return (
          op !== undefined && effect["type"] === sequencedEffectType(op.type)
        );
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
export const validateSequencedBatch = (
  value: unknown,
): Result<SequencedBatch, BatchRejection> => {
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
export const parseDocumentBatch = (
  json: string,
): Result<DocumentBatch, BatchRejection> => {
  if (encodeUtf8(json).byteLength > MAX_BATCH_WIRE_BYTES) {
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
