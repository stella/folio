import type { Mark, Node as PMNode, Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";
import { canJoin, canSplit } from "prosemirror-transform";
import { panic } from "better-result";

import {
  expectParagraphAttrs,
  expectRunFormattingOverrideMarkAttrs,
  expectRunPropertyChangeMarkAttrs,
} from "../prosemirror/attrs";
import { PPR_CHANGE_SCOPED_ATTR_KEYS } from "../prosemirror/commands/propertyChangeScope";
import { directParagraphAlignment } from "../prosemirror/paragraphAlignment";
import { getDocumentStyleResolver } from "../prosemirror/plugins/documentStyles";
import type { ParagraphPropertyChangeAttrs } from "../prosemirror/schema/nodes";
import { marksToTextFormatting } from "../prosemirror/conversion/fromProseDoc";
import { markStructuralChange } from "../prosemirror/extensions/features/ParagraphChangeTrackerExtension";
import { requestDeterministicParaIds } from "../prosemirror/extensions/features/ParaIdAllocatorExtension";
import {
  addedBreakCarrierBefore,
  finalParagraphsOf,
  paragraphEndsItsContainer,
} from "../prosemirror/containerFinalParagraph";
import { isZeroWidthAnchor } from "../prosemirror/zeroWidthAnchors";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import type { ParagraphAlignment, ParagraphFormatting, RunPropertyChange } from "../types/document";
import { stripBlockIdentityAttrs } from "./block-identity";
import { buildCleanBlockText } from "./clean-text";
import {
  hasInlineEmphasis,
  parseInlineEmphasisRuns,
  stripInlineEmphasisMarkers,
} from "./inline-emphasis";
import { hashFolioAIBlockText, isHiddenTableRow, normalizeFolioAIBlockText } from "./snapshot";
import {
  mergeTableRectangle,
  mergeTrackedVerticalTableCells,
  splitTableRectangle,
  splitTrackedVerticalTableCell,
} from "./table-cell-mutations";
import {
  planTableMutations,
  type TableMutationPlanTarget,
  type TableRectangle,
} from "./table-mutation-plan";
import {
  applyTableColumnDeletion,
  applyTableColumnInsertion,
  applyTableRowDeletion,
  applyTableRowInsertion,
  findTableColumnInsertion,
  findTableRowInsertion,
  getTableColumnCoordinateKey,
  splitCellParagraphTexts,
  type TableColumnDeletion,
  type TableColumnInsertion,
  type TableRowDeletion,
  type TableRowInsertion,
  markTableRowContent,
} from "./table-row-column-mutations";
import {
  tableFromTemplate,
  type FolioTableTemplates,
  type TableStructureRevision,
} from "./table-template";
import {
  findOutermostTableBoundary,
  findEnclosingTableCell,
  findEnclosingTableRow,
  tableRectangleCutsMergedCell,
} from "./table-targets";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditAppliedOperation,
  FolioAIEditApplyMode,
  FolioAIEditApplyResult,
  FolioAIEditNormalization,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
  FolioAIInlineFormatting,
  FolioAIEditSkipReason,
  FolioAIEditSkippedOperation,
  FolioAISignatureParty,
} from "./types";
import { diffWordSegments, type WordDiffGranularity } from "./word-diff";

/**
 * The only editor surface the apply logic touches: a current `state`
 * and a `dispatch` that swaps in the next one. A live `EditorView`
 * satisfies this structurally, and so does a headless seam
 * (`{ state, dispatch: (tr) => { state = state.apply(tr); } }`) — the
 * apply path never reaches for anything DOM-bound on the view, which is
 * what lets the same operation applier drive both the React editor and
 * the server-side reviewer in `./headless`.
 */
export type FolioAIEditView = {
  state: EditorState;
  dispatch: (transaction: Transaction) => void;
};

/**
 * Fixed revision provenance, for callers whose output must be reproducible.
 * Both halves of the stamp travel together because either one left to the
 * ambient clock (`new Date()` for the date, a `Date.now()`-seeded cursor for
 * the ids) makes the produced package differ between two runs over the same
 * inputs. `idSeed` starts a contiguous range the batch allocates from, so it
 * must sit above every revision id the target document already carries.
 */
export type FolioRevisionStamp = {
  /** ISO-8601 date stamped on every `w:ins` / `w:del` this batch produces. */
  date: string;
  /** First revision id the batch may allocate. */
  idSeed: number;
};

type ApplyFolioAIEditOperationsOptions = {
  view: FolioAIEditView;
  snapshot: FolioAIEditSnapshot;
  operations: readonly FolioAIEditOperation[];
  mode?: FolioAIEditApplyMode;
  author?: string;
  /** Optional author initials (w:initials) stamped alongside the author. */
  initials?: string;
  createCommentId?: (text: string) => number;
  /** Omit to stamp revisions from the wall clock and the shared id cursor. */
  revisionStamp?: FolioRevisionStamp;
  /** How a replacement's redline is cut. */
  wordDiff?: FolioWordDiffOptions;
  /**
   * Tables and rows an `insertTable` / `insertTableRow` operation should place
   * verbatim, by operation id.
   *
   * The operations describe their content as cell texts, which is what a
   * caller writing a table from nothing has. A caller copying one it already
   * holds — a comparison placing the target document's table — has the whole
   * thing: `w:tblPr`, the `w:tblGrid` widths, `w:trPr`, `w:tcPr` with its
   * spans and merges, and cell paragraphs with their own properties. Rebuilt
   * from text, none of that survives. Kept out of the operation payload
   * because it is a document node rather than JSON: the serialized contract
   * still describes a table by its cell texts.
   */
  tableTemplates?: FolioTableTemplates;
};

type ApplyFolioAIEditOperationsInternalOptions = ApplyFolioAIEditOperationsOptions & {
  revisionIdSeed?: number;
};

/**
 * How an apply batch cuts the redline for a replacement. Word-level by
 * default; character granularity marks the changed letters inside a token,
 * which reads well for a reference or a date. Normalization is deliberately
 * absent: a batch that leaves a difference unmarked does not accept back to
 * the text the caller asked for.
 */
export type FolioWordDiffOptions = { granularity?: WordDiffGranularity };

/**
 * An apply result plus where the batch left the revision-id counter.
 *
 * A caller writing several batches into one package — a comparison walking
 * the main story, then each header, footer and note — has to give every batch
 * a seed above the last id the previous one used, or two stories claim the
 * same `w:id` and a consumer resolving one revision resolves the other with
 * it. The batch is the only thing that knows how many ids it took, so it says
 * so rather than making the caller guess a stride.
 */
export type FolioAIEditApplyOutcome = FolioAIEditApplyResult & {
  /**
   * First revision id a following batch may allocate: one past the last id
   * this batch used, or the seed it started from when it allocated none.
   */
  nextRevisionId: number;
};

/**
 * Operation types applied in `"suggested"` mode. Every produced revision — an
 * inline mark, a whole-node `_suggestedInsert` marker, or a suggested
 * `trIns`/`trDel`/`cellMarker` — is stripped from serialized DOCX until
 * accepted. Cell merge/split (not row/column ops) and comment ops stay
 * `unsupportedMode`.
 */
const SUGGESTED_SUPPORTED_OPERATION_TYPES: ReadonlySet<FolioAIEditOperation["type"]> = new Set([
  "replaceInBlock",
  "replaceRange",
  "formatRange",
  "setBlockParagraphProperties",
  "replaceBlock",
  "deleteBlock",
  "insertAfterBlock",
  "insertBeforeBlock",
  "insertSignatureTable",
  "insertTableRow",
  "deleteTableRow",
  "insertTableColumn",
  "deleteTableColumn",
]);

/**
 * The attrs that render a list label, cleared together whenever a paragraph
 * stops being a list item. `w:numPr` alone is not enough: the marker attrs the
 * editor caches would keep drawing a label on a paragraph that is no longer in
 * the list. One frozen record rather than a list repeated at each call site,
 * because a list repeated is a list that drifts.
 */
const CLEARED_LIST_MARKER_ATTRS = Object.freeze({
  listMarker: null,
  listMarkerHidden: null,
  listLevelNumFmts: null,
  listLevelStarts: null,
  listAbstractNumId: null,
  listStartOverride: null,
});

type ResolvedOperationFields = {
  operation: FolioAIEditOperation;
  from: number;
  to: number;
  blockFrom: number;
  blockTo: number;
  blockNode: PMNode;
  /**
   * One entry per paragraph the insert produces. Usually a single entry;
   * more than one when the operation's `text` contained a line break and
   * was split into consecutive paragraphs (see {@link splitInsertParagraphTexts}).
   */
  insertTexts?: readonly string[];
  tableRowInsertion?: TableRowInsertion;
  tableRowDeletion?: TableRowDeletion;
  tableColumnInsertion?: TableColumnInsertion;
  tableColumnDeletion?: TableColumnDeletion;
  tableCellMerge?: TableCellMerge;
  tableCellSplit?: TableCellSplit;
  /** The table `deleteTable` removes, as it stood before the batch. */
  deletedTable?: { position: number; node: PMNode };
  commentId?: number;
  /**
   * Position in the input `operations` array, used as a secondary
   * sort key so same-position operations preserve the AI's logical
   * ordering when applied bottom-up.
   */
  originalIndex: number;
};

const REPLACE_BLOCK_IMPACT = {
  text: { changesStyle: false, changesText: true },
  style: { changesStyle: true, changesText: false },
  textAndStyle: { changesStyle: true, changesText: true },
} as const;

type ReplaceBlockImpact = keyof typeof REPLACE_BLOCK_IMPACT;
type ReplaceBlockOperation = Extract<FolioAIEditOperation, { type: "replaceBlock" }>;
type NonReplaceBlockOperation = Exclude<FolioAIEditOperation, ReplaceBlockOperation>;

type ResolvedReplaceBlockOperation = Omit<ResolvedOperationFields, "operation"> & {
  operation: ReplaceBlockOperation;
  replaceBlockImpact: ReplaceBlockImpact;
};

type ResolvedOperation =
  | ResolvedReplaceBlockOperation
  | (Omit<ResolvedOperationFields, "operation"> & {
      operation: NonReplaceBlockOperation;
      replaceBlockImpact?: never;
    });

const isResolvedReplaceBlockOperation = (
  item: ResolvedOperation,
): item is ResolvedReplaceBlockOperation => item.operation.type === "replaceBlock";

type ReplaceBlockImpactOptions = {
  changesText: boolean;
  changesStyle: boolean;
};

const resolveReplaceBlockImpact = ({
  changesText,
  changesStyle,
}: ReplaceBlockImpactOptions): ReplaceBlockImpact => {
  if (changesText) {
    return changesStyle ? "textAndStyle" : "text";
  }
  if (changesStyle) {
    return "style";
  }
  return panic("Cannot resolve a replaceBlock that changes neither text nor style");
};

/**
 * The attrs one `setBlockParagraphProperties` writes, or `null` when the
 * block already holds them. `styleId: null` clears the style; `listLevel`
 * moves `w:numPr/w:ilvl` and leaves `w:numId` alone, because a demoted item
 * stays in the same list; `listLevel: null` drops `w:numPr` entirely, which
 * is a paragraph that stopped being a list item.
 */
type ParagraphPropertiesPatchOptions = {
  node: PMNode;
  properties: FolioAIBlockParagraphProperties;
  resolvedAlignmentFromStyle: ParagraphAlignment | undefined;
};

const paragraphPropertiesPatch = ({
  node,
  properties,
  resolvedAlignmentFromStyle,
}: ParagraphPropertiesPatchOptions): Record<string, unknown> | null => {
  const attrs = expectParagraphAttrs(node);
  const currentDirectAlignment = directParagraphAlignment(attrs);
  const patch: Record<string, unknown> = {};
  let originalFormatting =
    attrs._originalFormatting === undefined || attrs._originalFormatting === null
      ? undefined
      : { ...attrs._originalFormatting };
  let originalFormattingChanged = false;
  const nextStyleId = properties.styleId;
  const styleChanged = nextStyleId !== undefined && (attrs.styleId ?? null) !== nextStyleId;
  if (styleChanged) {
    patch["styleId"] = nextStyleId;
    patch["alignmentFromStyle"] = resolvedAlignmentFromStyle;
    if (properties.alignment === undefined) {
      patch["alignment"] = currentDirectAlignment ?? resolvedAlignmentFromStyle ?? null;
    }
    originalFormatting ??= {};
    if (nextStyleId === null) {
      Reflect.deleteProperty(originalFormatting, "styleId");
    } else {
      originalFormatting.styleId = nextStyleId;
    }
    if (currentDirectAlignment !== undefined) {
      originalFormatting.alignment = currentDirectAlignment;
    }
    originalFormattingChanged = true;
  }
  if (properties.listLevel !== undefined) {
    const numPr: unknown = node.attrs["numPr"];
    const current =
      typeof numPr === "object" && numPr !== null && "ilvl" in numPr ? numPr.ilvl : undefined;
    if (properties.listLevel === null) {
      if (numPr !== null && numPr !== undefined) {
        patch["numPr"] = null;
        Object.assign(patch, CLEARED_LIST_MARKER_ATTRS);
      }
    } else if (current !== properties.listLevel) {
      const numId =
        typeof numPr === "object" && numPr !== null && "numId" in numPr ? numPr.numId : undefined;
      patch["numPr"] = {
        ...(typeof numId === "number" && { numId }),
        ilvl: properties.listLevel,
      };
    }
  }
  if (
    properties.alignment !== undefined &&
    (styleChanged || (currentDirectAlignment ?? null) !== properties.alignment)
  ) {
    originalFormatting ??= {};
    if (properties.alignment === null) {
      Reflect.deleteProperty(originalFormatting, "alignment");
    } else {
      originalFormatting.alignment = properties.alignment;
    }
    const alignmentFromStyle =
      properties.styleId === undefined ? attrs.alignmentFromStyle : resolvedAlignmentFromStyle;
    patch["alignment"] = properties.alignment ?? alignmentFromStyle ?? null;
    originalFormattingChanged = true;
  }
  if (originalFormattingChanged && originalFormatting !== undefined) {
    patch["_originalFormatting"] =
      Object.keys(originalFormatting).length > 0 ? originalFormatting : null;
  }
  return Object.keys(patch).length > 0 ? patch : null;
};

const REVISION_MARK_NAMES: ReadonlySet<string> = new Set(["insertion", "deletion"]);

/**
 * Strip insertion and deletion marks from the zero-width anchors the batch
 * touched.
 *
 * An operation marks a RANGE, and a zero-width anchor between two words is
 * inside it. The operation did not create or delete that anchor, so the anchor
 * must keep its existing revision ownership. Marking is by range everywhere;
 * the guard lives here once rather than at every call site.
 */
const withoutRevisionsOnZeroWidthAnchors = (
  tr: Transaction,
  batchRevisionIds: ReadonlySet<number>,
): Transaction => {
  const anchors: { from: number; to: number; marks: readonly Mark[] }[] = [];
  tr.doc.descendants((node, pos) => {
    if (!isZeroWidthAnchor(node)) {
      return true;
    }
    const marks = node.marks.filter(
      ({ attrs, type }) =>
        REVISION_MARK_NAMES.has(type.name) && batchRevisionIds.has(attrs["revisionId"]),
    );
    if (marks.length > 0) {
      anchors.push({ from: pos, to: pos + node.nodeSize, marks });
    }
    return false;
  });
  for (const { from, to, marks } of anchors) {
    for (const mark of marks) {
      tr = tr.removeMark(from, to, mark);
    }
  }
  return tr;
};

/**
 * The block's non-text inline children a block deletion still has to mark, as
 * positions in `tr.doc`.
 *
 * Skips what is already deleted, so an earlier revision's run is not marked
 * twice, and skips zero-width anchors because the operation did not create or
 * delete them; their existing revision ownership must remain unchanged.
 */
const undeletedContentAtomRanges = (
  tr: Transaction,
  blockFrom: number,
): { from: number; to: number }[] => {
  const position = tr.mapping.map(blockFrom);
  const block = tr.doc.nodeAt(position);
  if (!block?.isTextblock) {
    return [];
  }
  const ranges: { from: number; to: number }[] = [];
  block.forEach((child, offset) => {
    if (
      child.isText ||
      isZeroWidthAnchor(child) ||
      child.marks.some(({ type }) => type.name === "deletion")
    ) {
      return;
    }
    const from = position + 1 + offset;
    ranges.push({ from, to: from + child.nodeSize });
  });
  return ranges;
};

/** The in-scope paragraph properties as they stand, for a `w:pPrChange` record. */
const paragraphPropertiesSnapshot = (
  node: PMNode,
): NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]> => {
  const attrs = expectParagraphAttrs(node);
  const snapshot: Record<string, unknown> = {};
  for (const key of PPR_CHANGE_SCOPED_ATTR_KEYS) {
    if (key === "alignment") {
      continue;
    }
    const value: unknown = attrs[key];
    if (
      key === "hangingIndent" &&
      value === false &&
      attrs._originalFormatting?.hangingIndent === undefined
    ) {
      continue;
    }
    if (value !== null && value !== undefined) {
      snapshot[key] = value;
    }
  }
  const directAlignment = directParagraphAlignment(attrs);
  if (directAlignment !== undefined) {
    snapshot["alignment"] = directAlignment;
  }
  return snapshot;
};

/** The carrier pPr before this batch changed it, for final-mark rotation. */
const paragraphPropertiesBeforeBatch = (
  node: PMNode,
  batchRevisionIds: ReadonlySet<number>,
): NonNullable<ParagraphPropertyChangeAttrs["previousFormatting"]> => {
  const propertyChanges = expectParagraphAttrs(node)._propertyChanges;
  if (!Array.isArray(propertyChanges)) {
    return paragraphPropertiesSnapshot(node);
  }
  const earliestBatchChange = propertyChanges.find(({ info }) => batchRevisionIds.has(info.id));
  if (!earliestBatchChange) {
    return paragraphPropertiesSnapshot(node);
  }
  return earliestBatchChange.previousFormatting ?? {};
};

type ApplyReplaceBlockStyleIdResult = {
  tr: Transaction;
  revisionId: number | null;
};

type ApplyReplaceBlockStyleIdOptions = {
  item: ResolvedOperation;
  tr: Transaction;
  styleResolver: ReturnType<typeof getDocumentStyleResolver>;
  revisionInfo?: ParagraphPropertyChangeAttrs["info"];
};

const applyReplaceBlockStyleId = ({
  item,
  tr,
  styleResolver,
  revisionInfo,
}: ApplyReplaceBlockStyleIdOptions): ApplyReplaceBlockStyleIdResult => {
  if (item.operation.type !== "replaceBlock" || item.operation.styleId === undefined) {
    return { tr, revisionId: null };
  }

  const blockPosition = tr.mapping.map(item.blockFrom, -1);
  const block = tr.doc.nodeAt(blockPosition);
  if (!block) {
    return { tr, revisionId: null };
  }

  const attrs = expectParagraphAttrs(block);
  const resolvedAlignmentFromStyle =
    styleResolver?.resolveParagraphStyle(item.operation.styleId).paragraphFormatting?.alignment ??
    ((item.operation.styleId ?? undefined) === (attrs.styleId ?? undefined)
      ? attrs.alignmentFromStyle
      : undefined);
  const patch = paragraphPropertiesPatch({
    node: block,
    properties: { styleId: item.operation.styleId },
    resolvedAlignmentFromStyle,
  });
  if (patch === null) {
    return { tr, revisionId: null };
  }

  const existing = attrs._propertyChanges;
  const change: ParagraphPropertyChangeAttrs | null = revisionInfo
    ? {
        type: "paragraphPropertyChange",
        info: revisionInfo,
        previousFormatting: paragraphPropertiesSnapshot(block),
      }
    : null;
  return {
    tr: tr.setNodeMarkup(blockPosition, undefined, {
      ...block.attrs,
      ...patch,
      ...(change
        ? { _propertyChanges: [...(Array.isArray(existing) ? existing : []), change] }
        : {}),
    }),
    revisionId: change?.info.id ?? null,
  };
};

type ApplyInlineFormattingOptions = {
  tr: Transaction;
  schema: Schema;
  from: number;
  to: number;
  formatting: InlineFormattingPatch;
  includedProperties?: readonly string[];
};

const REPLACEMENT_BACKGROUND_CLEAR_FORMATTING = {
  highlight: false,
  runShading: false,
} as const;

type InlineFormattingPatch = FolioAIInlineFormatting &
  Partial<typeof REPLACEMENT_BACKGROUND_CLEAR_FORMATTING>;

const INLINE_FORMATTING_MARK_NAMES = {
  bold: "bold",
  italic: "italic",
  underline: "underline",
  strike: "strike",
  fontFamily: "fontFamily",
  fontSizePt: "fontSize",
  color: "textColor",
  highlight: "highlight",
  runShading: "runShading",
} as const;

const DIRECT_FONT_PROPERTIES = ["fontFamily", "fontSize", "color"] as const;

const formattingMarkName = (property: string): string | null => {
  if (!Object.hasOwn(INLINE_FORMATTING_MARK_NAMES, property)) {
    return null;
  }
  const name: unknown = Reflect.get(INLINE_FORMATTING_MARK_NAMES, property);
  return typeof name === "string" ? name : null;
};

const formattingMarkAttrs = (
  property: string,
  value: boolean | string | number,
): Record<string, unknown> | null => {
  switch (property) {
    case "underline":
      return { style: "single" };
    case "fontFamily":
      return typeof value === "string" ? { ascii: value, hAnsi: value } : null;
    case "fontSizePt":
      return typeof value === "number" ? { size: value * 2 } : null;
    case "color":
      return typeof value === "string" ? { rgb: value.replace(/^#/u, "").toUpperCase() } : null;
    default:
      return {};
  }
};

const applyInlineFormatting = ({
  tr,
  schema,
  from,
  to,
  formatting,
  includedProperties,
}: ApplyInlineFormattingOptions): Transaction => {
  for (const [property, value] of Object.entries(formatting)) {
    if (includedProperties && !includedProperties.includes(property)) {
      continue;
    }
    const markName = formattingMarkName(property);
    const markType = markName ? schema.marks[markName] : undefined;
    if (!markType) {
      continue;
    }
    if (value !== false && value !== null) {
      const attrs = formattingMarkAttrs(property, value);
      if (attrs) {
        tr.addMark(from, to, markType.create(attrs));
      }
      continue;
    }
    tr.removeMark(from, to, markType);
  }
  applyDirectFontProvenance({
    tr,
    schema,
    from,
    to,
    formatting,
    ...(includedProperties ? { includedProperties } : {}),
  });
  return tr;
};

const applyDirectFontProvenance = ({
  tr,
  schema,
  from,
  to,
  formatting,
  includedProperties,
}: ApplyInlineFormattingOptions): void => {
  if ("highlight" in formatting) {
    return;
  }
  const updates = [
    ["fontFamily", "fontFamily", formatting.fontFamily],
    ["fontSize", "fontSizePt", formatting.fontSizePt],
    ["color", "color", formatting.color],
  ] as const;
  const changed = updates.filter(
    ([, inputProperty, value]) =>
      value !== undefined && (!includedProperties || includedProperties.includes(inputProperty)),
  );
  const markType = schema.marks["runFormattingOverride"];
  if (!markType || changed.length === 0) {
    return;
  }

  tr.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) {
      return;
    }
    const segmentFrom = Math.max(from, pos);
    const segmentTo = Math.min(to, pos + node.nodeSize);
    const existing = node.marks.find((mark) => mark.type === markType);
    const attrs = existing ? expectRunFormattingOverrideMarkAttrs(existing) : {};
    const directFontProperties = new Set(attrs.directFontProperties);
    for (const [property, , value] of changed) {
      if (value === null) {
        directFontProperties.delete(property);
      } else {
        directFontProperties.add(property);
      }
    }

    tr.removeMark(segmentFrom, segmentTo, markType);
    const nextAttrs = {
      ...attrs,
      directFontProperties: DIRECT_FONT_PROPERTIES.filter((property) =>
        directFontProperties.has(property),
      ),
    };
    if (
      Object.entries(nextAttrs).some(([key, value]) =>
        key === "directFontProperties"
          ? Array.isArray(value) && value.length > 0
          : value !== null && value !== undefined,
      )
    ) {
      tr.addMark(segmentFrom, segmentTo, markType.create(nextAttrs));
    }
  });
};

const directFontPropertyForFormattingProperty = (
  property: string,
): (typeof DIRECT_FONT_PROPERTIES)[number] | null => {
  switch (property) {
    case "fontFamily":
      return "fontFamily";
    case "fontSizePt":
      return "fontSize";
    case "color":
      return "color";
    default:
      return null;
  }
};

const formattingPropertyWouldChange = (
  marks: readonly Mark[],
  property: string,
  value: boolean | string | number | null,
): boolean => {
  const markName = formattingMarkName(property);
  const mark = marks.find((candidate) => candidate.type.name === markName);
  const directFontProperty = directFontPropertyForFormattingProperty(property);
  if (directFontProperty) {
    const overrideMark = marks.find((candidate) => candidate.type.name === "runFormattingOverride");
    const isDirect = overrideMark
      ? (expectRunFormattingOverrideMarkAttrs(overrideMark).directFontProperties ?? []).includes(
          directFontProperty,
        )
      : false;
    if (value === false || value === null) {
      return isDirect;
    }
    if (!isDirect) {
      return true;
    }
  } else if (value === false || value === null) {
    return mark !== undefined;
  }
  if (property === "underline") {
    return mark?.attrs["style"] !== "single";
  }
  if (property === "fontFamily") {
    return mark?.attrs["ascii"] !== value || mark?.attrs["hAnsi"] !== value;
  }
  if (property === "fontSizePt") {
    return Number(mark?.attrs["size"]) !== Number(value) * 2;
  }
  if (property === "color") {
    const current = mark?.attrs["rgb"];
    const normalizedCurrent =
      typeof current === "string" ? current.replace(/^#/u, "").toUpperCase() : null;
    return normalizedCurrent !== String(value).replace(/^#/u, "").toUpperCase();
  }
  return mark === undefined;
};

const formattingChangesForMarks = (
  marks: readonly Mark[],
  formatting: InlineFormattingPatch,
): string[] => {
  const changes: string[] = [];
  for (const [property, value] of Object.entries(formatting)) {
    if (formattingPropertyWouldChange(marks, property, value)) {
      changes.push(property);
    }
  }
  return changes;
};

type ApplyTrackedInlineFormattingOptions = ApplyInlineFormattingOptions & {
  doc: PMNode;
  revisionId: number;
  author: string;
  date: string;
  /** Optional author initials (w:initials) stamped alongside the author. */
  initials?: string | undefined;
  /** Non-null stamps the produced `runPropertyChange` mark as a suggestion. */
  suggestionId?: string | null;
};

type ClearReplacementBackgroundOptions = {
  tr: Transaction;
  schema: Schema;
  from: number;
  to: number;
  mode: FolioAIEditApplyMode;
  revisionId: number;
  author: string;
  date: string;
  initials?: string | undefined;
  suggestionId?: string | null;
};

const clearReplacementBackground = ({
  tr,
  schema,
  from,
  to,
  mode,
  revisionId,
  author,
  date,
  initials,
  suggestionId = null,
}: ClearReplacementBackgroundOptions): Transaction => {
  const hasBackground = [schema.marks["highlight"], schema.marks["runShading"]].some(
    (markType) => markType !== undefined && tr.doc.rangeHasMark(from, to, markType),
  );
  if (!hasBackground) {
    return tr;
  }
  if (mode === "direct") {
    return applyInlineFormatting({
      tr,
      schema,
      from,
      to,
      formatting: REPLACEMENT_BACKGROUND_CLEAR_FORMATTING,
    });
  }
  return applyTrackedInlineFormatting({
    tr,
    schema,
    doc: tr.doc,
    from,
    to,
    formatting: REPLACEMENT_BACKGROUND_CLEAR_FORMATTING,
    revisionId,
    author,
    date,
    initials,
    suggestionId,
  });
};

const applyTrackedInlineFormatting = ({
  tr,
  schema,
  doc,
  from,
  to,
  formatting,
  revisionId,
  author,
  date,
  initials,
  suggestionId = null,
}: ApplyTrackedInlineFormattingOptions): Transaction => {
  const propertyChangeType = schema.marks["runPropertyChange"];
  if (!propertyChangeType) {
    return tr;
  }

  const segments: {
    from: number;
    to: number;
    includedProperties: string[];
    changes: RunPropertyChange[];
  }[] = [];
  doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText) {
      return;
    }
    const includedProperties = formattingChangesForMarks(node.marks, formatting);
    if (includedProperties.length === 0) {
      return;
    }
    const segmentFrom = Math.max(from, pos);
    const segmentTo = Math.min(to, pos + node.nodeSize);
    const previousFormatting = marksToTextFormatting(node.marks);
    const existingMark = node.marks.find((mark) => mark.type === propertyChangeType);
    const existingChanges = existingMark
      ? expectRunPropertyChangeMarkAttrs(existingMark).changes
      : [];
    const change: RunPropertyChange = {
      type: "runPropertyChange",
      info: { id: revisionId, author, date, ...(initials ? { initials } : {}) },
      ...(Object.keys(previousFormatting).length > 0 ? { previousFormatting } : {}),
    };
    segments.push({
      from: segmentFrom,
      to: segmentTo,
      includedProperties,
      changes: [...existingChanges, change],
    });
  });

  if (segments.length === 0) {
    return tr;
  }
  const suggestionAttrs = suggestionId === null ? {} : { provenance: "suggested", suggestionId };
  for (const segment of segments) {
    applyInlineFormatting({
      tr,
      schema,
      from: segment.from,
      to: segment.to,
      formatting,
      includedProperties: segment.includedProperties,
    });
    tr.addMark(
      segment.from,
      segment.to,
      propertyChangeType.create({ changes: segment.changes, ...suggestionAttrs }),
    );
  }
  return tr;
};

type LiveBlockEntry = { from: number; to: number; node: PMNode };

/**
 * Module-scoped monotonic counter for tracked-change revision ids.
 * Seeded once from `Date.now()` so ids are roughly time-ordered for
 * humans reading raw DOCX, then incremented per allocation. A bare
 * `Date.now()` seed per applyAIEditOperations call would collide
 * across batches that fire within the same millisecond (the panel's
 * Accept-all loop does exactly that — multiple calls in tight
 * succession). Reserving a contiguous range up front guarantees
 * uniqueness across overlapping calls in the same JS realm.
 */
let revisionIdCursor = Date.now() * 1000;
const nextRevisionSeed = (revisionIdCount: number): number => {
  // The caller has already summed a safe per-operation reservation (see
  // `estimateRevisionIdReservation`); reserving exactly that many ids keeps
  // this batch's range from overlapping the next `nextRevisionSeed` call's
  // range. Returning the start of the reserved range as the seed is enough
  // — the caller bumps it.
  const start = revisionIdCursor;
  revisionIdCursor += Math.max(revisionIdCount, 1);
  return start;
};

/**
 * Ids one resolved operation may allocate from the shared revision-id range
 * reserved by `nextRevisionSeed`. Most operation types allocate at most
 * three (delete + insert + background-format clearing); reserving four per
 * operation is a safe cushion above that. A multi-paragraph
 * `insertAfterBlock` / `insertBeforeBlock` (`text` split on line breaks,
 * see `splitInsertParagraphTexts`) allocates one id per paragraph in
 * tracked-changes mode plus one for each paragraph MARK it brings, so it
 * needs more than four once split into more than two paragraphs — reserving
 * less than that would let a later `nextRevisionSeed` call reuse an id this
 * operation already stamped on the document.
 */
const REVISION_IDS_PER_OPERATION = 4;
/** Ids one inserted paragraph allocates: its runs, and its paragraph mark. */
const REVISION_IDS_PER_INSERTED_PARAGRAPH = 2;
const estimateRevisionIdReservation = (item: ResolvedOperation): number => {
  if (item.operation.type === "insertAfterBlock" || item.operation.type === "insertBeforeBlock") {
    return Math.max(
      (item.insertTexts?.length ?? 1) * REVISION_IDS_PER_INSERTED_PARAGRAPH,
      REVISION_IDS_PER_OPERATION,
    );
  }
  return REVISION_IDS_PER_OPERATION;
};

/**
 * Walk the live doc once and bucket every textblock by its
 * normalised text hash. Resolution then maps each snapshot anchor
 * to the live block at the same ordinal among same-hash siblings —
 * unrelated edits that shift absolute positions no longer break
 * the lookup, and a sibling sharing text content with the target
 * doesn't trigger a false "changed" skip either.
 */
const collectLiveBlocksByHash = (doc: PMNode) => {
  const byHash = new Map<string, LiveBlockEntry[]>();
  doc.descendants((node, pos) => {
    // The snapshot skips a hidden row's whole subtree, and resolution pairs a
    // snapshot anchor with the live block at the same ordinal among the blocks
    // sharing its hash. Counting hidden blocks here and not there shifts every
    // later ordinal, which resolves an operation onto the wrong paragraph
    // rather than skipping it. The two walks have to agree.
    if (isHiddenTableRow(node)) {
      return false;
    }
    if (!node.isTextblock) {
      return true;
    }
    // Hash from the post-tracked-changes view so the snapshot
    // (taken with the same view) and live doc bucket the same
    // block under the same key. Otherwise a block mid-edit gets a
    // different hash than the snapshot recorded and the resolver
    // skips it as "changed".
    const cleanText = buildCleanBlockText(node, pos).text;
    const hash = hashFolioAIBlockText(normalizeFolioAIBlockText(cleanText));
    const bucket = byHash.get(hash) ?? [];
    bucket.push({ from: pos, to: pos + node.nodeSize, node });
    byHash.set(hash, bucket);
    return false;
  });
  return byHash;
};

/**
 * Index live textblocks by their `w14:paraId`. Used by the resolver
 * to prefer a paraId-anchored lookup when the snapshot id encodes
 * one. Direct lookup avoids
 * the hash+ordinal failure mode where an earlier-in-document
 * duplicate of the same text gets picked instead of the actual
 * referenced paragraph.
 */
const collectLiveBlocksByParaId = (doc: PMNode) => {
  const byParaId = new Map<string, LiveBlockEntry>();
  doc.descendants((node, pos) => {
    // The same rule as the hash index and the snapshot: a hidden row's whole
    // subtree is not walked. A paraId is only unique because Word keeps it so,
    // and a package that reuses one across a hidden and a visible paragraph
    // would otherwise resolve the visible block onto content no reader can
    // see, and edit it there. All three walks have to agree on what exists.
    if (isHiddenTableRow(node)) {
      return false;
    }
    if (!node.isTextblock) {
      return true;
    }
    const paraId: unknown = node.attrs["paraId"];
    if (typeof paraId !== "string" || paraId.length === 0) {
      return false;
    }
    // First-write-wins so an Enter-split duplicate (briefly co-existing
    // before the allocator re-issues) doesn't override the original.
    if (!byParaId.has(paraId)) {
      byParaId.set(paraId, { from: pos, to: pos + node.nodeSize, node });
    }
    return false;
  });
  return byParaId;
};

type TableCellMerge = {
  tablePosition: number;
  rectangle: TableRectangle;
};

type TableCellSplit = TableCellMerge;

const getTableMutationPlanTarget = (item: ResolvedOperation): TableMutationPlanTarget => {
  if (item.tableCellMerge) {
    return { type: "mergeCells", ...item.tableCellMerge };
  }
  if (item.tableCellSplit) {
    return { type: "splitCell", ...item.tableCellSplit };
  }

  const tablePosition =
    item.tableRowInsertion?.tableStart !== undefined
      ? item.tableRowInsertion.tableStart - 1
      : (item.tableRowDeletion?.tablePosition ??
        item.tableColumnInsertion?.tablePosition ??
        item.tableColumnDeletion?.tablePosition);
  if (tablePosition !== undefined) {
    return { type: "tableStructure", tablePosition };
  }
  return { type: "none" };
};

/**
 * The snapshot recorded an `hashOccurrenceCount` per anchor but
 * not which ordinal within that bucket the block was — recompute
 * on demand from the snapshot's anchor map. Stable iteration
 * (object insertion order) means anchors with the same hash come
 * out in document order, which is what we want.
 */
const ordinalAmongSameHash = (snapshot: FolioAIEditSnapshot, blockId: string): number => {
  const target = snapshot.anchors[blockId];
  if (!target) {
    return -1;
  }
  let ordinal = 0;
  for (const anchor of Object.values(snapshot.anchors)) {
    if (anchor.id === blockId) {
      return ordinal;
    }
    if (anchor.textHash === target.textHash) {
      ordinal += 1;
    }
  }
  return -1;
};

/**
 * Length of the underscore signature rule. Mirrors the constant
 * in `docx-core/legal-source/compile.ts` so the on-screen signature
 * line matches what `create-document` produces.
 */
const SIGNATURE_LINE = "_".repeat(28);

type BuildTableNodeOptions = {
  schema: Schema;
  rows: readonly (readonly string[])[];
  /** Present in tracked mode: every row is stamped as an insertion. */
  revision?: TableStructureRevision;
};

/**
 * A plain table from a grid of cell texts, `null` when the schema has no
 * tables. In tracked mode every row carries `trIns` and every run inside it
 * carries `w:ins`, which is how the format says "this table is new": there is
 * no whole-table insertion element, only rows that were inserted, and a
 * consumer that reads only run-level revisions keeps the text of a rejected
 * insertion when the row marker stands alone.
 */
const buildTableNode = ({ schema, rows, revision }: BuildTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (!paragraphType || !cellType || !rowType || !tableType || rows.length === 0) {
    return null;
  }
  const insertionType = schema.marks["insertion"];
  const marks =
    revision && insertionType
      ? [
          insertionType.create({
            revisionId: revision.revisionId,
            author: revision.author,
            date: revision.date,
            ...(revision.initials != null && { initials: revision.initials }),
            ...(revision.provenance != null && { provenance: revision.provenance }),
            ...(revision.suggestionId != null && { suggestionId: revision.suggestionId }),
          }),
        ]
      : undefined;
  const rowNodes = rows.map((cells) =>
    rowType.create(
      revision ? { trIns: revision } : null,
      cells.map((text) =>
        cellType.create(
          null,
          splitCellParagraphTexts(text).map((line) =>
            paragraphType.create(null, line.length > 0 ? schema.text(line, marks) : null),
          ),
        ),
      ),
    ),
  );
  return tableType.create(null, rowNodes);
};

type BuildSignatureTableNodeOptions = {
  schema: Schema;
  parties: readonly FolioAISignatureParty[];
};

/**
 * Build a borderless PM table mirroring docx-core's
 * `signatureTable` helper: one row, one cell per party, each cell
 * containing party name (bold), two spacer paragraphs, a signature
 * rule, then optional signatory and italic title lines.
 *
 * Returns `null` when the editor schema is missing one of the
 * required node types — callers should surface the op as a skip
 * rather than crashing.
 */
const buildSignatureTableNode = ({
  schema,
  parties,
}: BuildSignatureTableNodeOptions): PMNode | null => {
  const paragraphType = schema.nodes["paragraph"];
  const cellType = schema.nodes["tableCell"];
  const rowType = schema.nodes["tableRow"];
  const tableType = schema.nodes["table"];
  if (!paragraphType || !cellType || !rowType || !tableType) {
    return null;
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];

  const buildParagraph = (
    text: string,
    options: {
      styleId: string;
      bold?: boolean;
      italic?: boolean;
    },
  ): PMNode => {
    const marks: Mark[] = [];
    if (options.bold && boldType) {
      marks.push(boldType.create());
    }
    if (options.italic && italicType) {
      marks.push(italicType.create());
    }
    const content = text.length > 0 ? schema.text(text, marks) : null;
    return paragraphType.create({ styleId: options.styleId }, content);
  };

  const buildCell = (party: FolioAISignatureParty): PMNode => {
    const cellContent: PMNode[] = [
      buildParagraph(party.name, { styleId: "SignatureParty", bold: true }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph("", { styleId: "SignatureSpacer" }),
      buildParagraph(SIGNATURE_LINE, { styleId: "SignatureRule" }),
    ];
    if (party.signatory && party.signatory.length > 0) {
      cellContent.push(buildParagraph(party.signatory, { styleId: "SignatureField" }));
    }
    if (party.title && party.title.length > 0) {
      cellContent.push(
        buildParagraph(party.title, {
          styleId: "SignatureField",
          italic: true,
        }),
      );
    }
    // Cell attrs left to schema defaults — column widths and
    // borders are decided by the table renderer; we just need the
    // structural shape.
    return cellType.create({}, cellContent);
  };

  const cells = parties.map(buildCell);
  const row = rowType.create({}, cells);
  return tableType.create({}, row);
};

const LINE_BREAK_PATTERN = /\r\n|\r|\n/;

/**
 * Split `insertAfterBlock` / `insertBeforeBlock` text on line breaks into one
 * string per paragraph the operation should produce. A model routinely sends
 * a whole clause — heading plus body — as one `text` with embedded newlines;
 * Word has no such thing as a paragraph containing a line break, so each line
 * becomes its own paragraph instead of literal newline characters inside one.
 * Blank (whitespace-only) lines are dropped rather than becoming empty
 * paragraphs, so a trailing/leading newline collapses away. Falls back to a
 * single empty string when every line is blank, so callers never get zero
 * paragraphs out of non-empty input.
 */
const splitInsertParagraphTexts = (text: string): string[] => {
  const nonBlankLines = text.split(LINE_BREAK_PATTERN).filter((line) => line.trim().length > 0);
  return nonBlankLines.length > 0 ? nonBlankLines : [""];
};

/**
 * Build the inline content for an inserted or replaced block, promoting the
 * model's `**bold**` / `***bold italic***` markdown into real Word marks (the
 * edit-tool schema has no inline-format channel, so the model improvises with
 * markdown that would otherwise land as literal asterisks). Falls back to a
 * single verbatim text node when no emphasis is present, so plain prose is
 * never reshaped. `baseMarks` (insertion / comment) ride on every run.
 */
const buildEmphasisInlineContent = (
  schema: Schema,
  text: string,
  baseMarks: readonly Mark[],
): PMNode[] => {
  const runs = parseInlineEmphasisRuns(text);
  if (!runs.some((run) => run.bold || run.italic)) {
    return [schema.text(text, [...baseMarks])];
  }
  const boldType = schema.marks["bold"];
  const italicType = schema.marks["italic"];
  const nodes: PMNode[] = [];
  for (const run of runs) {
    if (run.text.length === 0) {
      continue;
    }
    const marks: Mark[] = [...baseMarks];
    if (run.bold && boldType) {
      marks.push(boldType.create());
    }
    if (run.italic && italicType) {
      marks.push(italicType.create());
    }
    nodes.push(schema.text(run.text, marks));
  }
  return nodes.length > 0 ? nodes : [schema.text(text, [...baseMarks])];
};

type BuildInsertedParagraphsOptions = {
  item: ResolvedOperation;
  schema: Schema;
  alignmentFromStyle: ParagraphAlignment | undefined;
  mode: FolioAIEditApplyMode;
  author: string;
  date: string;
  initials: string | undefined;
  commentMark: Mark | null;
  suggestionId: string | null;
  revisionSeed: number;
  isPairedMove: (moveId: string | undefined) => moveId is string;
};

type BuiltInsertedParagraphs = {
  nodes: PMNode[];
  revisionIds: number[];
  nextRevisionId: number;
};

const isBatchableParagraphInsertion = (
  item: ResolvedOperation,
  mode: FolioAIEditApplyMode,
): boolean => {
  if (item.operation.type !== "insertAfterBlock" && item.operation.type !== "insertBeforeBlock") {
    return false;
  }
  const insertTexts = item.insertTexts ?? [""];
  const isEmptyInsert = insertTexts.length === 1 && insertTexts[0]?.length === 0;
  return !(mode === "tracked-changes" && item.operation.pageBreakBefore === true && isEmptyInsert);
};

type RotateAddedFinalBreaksOptions = {
  tr: Transaction;
  /** Revision ids this batch minted, so marks it did not write are left alone. */
  batchRevisionIds: ReadonlySet<number>;
  revisionSeed: number;
  author: string;
  date: string;
  initials: string | undefined;
};

type RotatedAddedFinalBreaks = {
  transaction: Transaction;
  nextRevisionId: number;
};

/**
 * Move an ADDED paragraph break off every paragraph its container ends with.
 *
 * A container's final mark carries no revision in either direction: nothing
 * follows it, so "join this paragraph with the one after it" and its mirror
 * both state an edit that cannot be carried out, and a reader is left with a
 * revision neither accepting nor rejecting everything can clear.
 *
 * Appending at a container's end therefore rotates exactly as removing from it
 * does. The break that was added sits between the paragraph the run was
 * appended after and the first appended one, so THAT paragraph's mark is the
 * inserted one, each appended paragraph but the last keeps an inserted mark of
 * its own, and the paragraph the container now ends with takes the free mark
 * the run was appended after. Only which paragraph is left markless changes.
 *
 * A paragraph's properties live on its mark, so the paragraph that inherits
 * the free one records the other's as `w:pPrChange` — what a rejection reads
 * to put the container's ending back the way it was.
 *
 * This runs once over the finished document rather than at each insertion:
 * which paragraph ends a container is only settled when the batch is, and an
 * insertion that looked final was undone by the next operation writing a table
 * after it.
 */
const withRotatedAddedFinalBreaks = ({
  tr,
  batchRevisionIds,
  revisionSeed,
  author,
  date,
  initials,
}: RotateAddedFinalBreaksOptions): RotatedAddedFinalBreaks => {
  const paragraphTypeName = tr.doc.type.schema.nodes["paragraph"]?.name ?? "paragraph";
  const rotations = finalParagraphsOf(tr.doc, paragraphTypeName).filter(({ node }) => {
    const mark: unknown = node.attrs["pPrMark"];
    if (typeof mark !== "object" || mark === null || !("kind" in mark)) {
      return false;
    }
    if (mark.kind !== "ins" && mark.kind !== "moveTo") {
      return false;
    }
    const info: unknown = "info" in mark ? mark.info : undefined;
    const revisionId =
      typeof info === "object" && info !== null && "id" in info ? info.id : undefined;
    return typeof revisionId === "number" && batchRevisionIds.has(revisionId);
  });

  let next = tr;
  let nextRevisionId = revisionSeed;
  // Attribute writes do not move anything, so the positions stay valid.
  for (const { position: finalPosition, node: final } of rotations) {
    const carrier = addedBreakCarrierBefore(next.doc.resolve(finalPosition), final.type.name);
    if (!carrier) {
      // Nothing to hand the break to. Writing it would be worse than losing
      // it: the redline would carry a revision no reader can resolve.
      next = next.setNodeAttribute(finalPosition, "pPrMark", null);
      continue;
    }
    next = next.setNodeAttribute(carrier.position, "pPrMark", final.attrs["pPrMark"]);
    const previousFormatting = paragraphPropertiesBeforeBatch(carrier.node, batchRevisionIds);
    // Both snapshots are built by the same fixed walk over the in-scope keys,
    // so comparing them serialized compares them key for key.
    if (JSON.stringify(previousFormatting) === JSON.stringify(paragraphPropertiesSnapshot(final))) {
      next = next.setNodeAttribute(finalPosition, "pPrMark", null);
      continue;
    }
    const existing = expectParagraphAttrs(final)._propertyChanges;
    next = next.setNodeMarkup(finalPosition, undefined, {
      ...final.attrs,
      pPrMark: null,
      _propertyChanges: [
        ...(Array.isArray(existing) ? existing : []),
        {
          type: "paragraphPropertyChange",
          info: { id: nextRevisionId++, author, date, ...(initials ? { initials } : {}) },
          previousFormatting,
        } satisfies ParagraphPropertyChangeAttrs,
      ],
    });
  }
  return { transaction: next, nextRevisionId };
};

const buildInsertedParagraphs = ({
  item,
  schema,
  alignmentFromStyle,
  mode,
  author,
  date,
  initials,
  commentMark,
  suggestionId,
  revisionSeed,
  isPairedMove,
}: BuildInsertedParagraphsOptions): BuiltInsertedParagraphs => {
  const operation = item.operation;
  if (operation.type !== "insertAfterBlock" && operation.type !== "insertBeforeBlock") {
    panic("Only paragraph insertions can build inserted paragraphs", { type: operation.type });
  }
  const insertionType = schema.marks["insertion"];
  const producesTrackedChanges = mode !== "direct";
  const isSuggested = mode === "suggested";
  const trackedRevisionExtras = {
    ...(initials ? { initials } : {}),
    ...(suggestionId !== null ? { provenance: "suggested" as const, suggestionId } : {}),
  };
  // Only the first paragraph split from one operation inherits the anchor's
  // formatting. Later lines are new body paragraphs, not anchor clones.
  const baseAttrs =
    operation.inheritFormatting === false ? {} : stripBlockIdentityAttrs(item.blockNode.attrs);
  const insertTexts = item.insertTexts ?? [""];
  const revisionIds: number[] = [];
  const nodes: PMNode[] = [];
  let nextRevisionId = revisionSeed;

  for (const [paragraphIndex, text] of insertTexts.entries()) {
    const isFirstParagraph = paragraphIndex === 0;
    const marks: Mark[] = [];
    let paragraphRevisionId: number | null = null;
    if (producesTrackedChanges && insertionType) {
      paragraphRevisionId = nextRevisionId++;
      marks.push(
        insertionType.create({
          revisionId: paragraphRevisionId,
          author,
          date,
          ...(isPairedMove(operation.moveId) && { moveKind: "moveTo" }),
          ...trackedRevisionExtras,
        }),
      );
      revisionIds.push(paragraphRevisionId);
    }
    if (commentMark) {
      marks.push(commentMark);
    }
    const content = text.length > 0 ? buildEmphasisInlineContent(schema, text, marks) : null;
    const attrs: Record<string, unknown> = isFirstParagraph ? { ...baseAttrs } : {};
    if (isFirstParagraph && operation.pageBreakBefore === true) {
      attrs["pageBreakBefore"] = true;
    }
    if (isFirstParagraph && operation.listLevel === null) {
      attrs["numPr"] = null;
      Object.assign(attrs, CLEARED_LIST_MARKER_ATTRS);
    } else if (isFirstParagraph && operation.listLevel !== undefined) {
      const anchorNumPr: unknown = baseAttrs["numPr"];
      const numId =
        typeof anchorNumPr === "object" && anchorNumPr !== null && "numId" in anchorNumPr
          ? anchorNumPr.numId
          : undefined;
      attrs["numPr"] = {
        ...(typeof numId === "number" && { numId }),
        ilvl: operation.listLevel,
      };
    }
    if (isFirstParagraph && operation.styleId !== undefined) {
      attrs["styleId"] = operation.styleId;
      if (operation.inheritFormatting !== false && operation.styleId !== null) {
        Object.assign(attrs, CLEARED_LIST_MARKER_ATTRS);
      }
    }
    if (
      isFirstParagraph &&
      (operation.styleId !== undefined || operation.alignment !== undefined)
    ) {
      const inheritedDirectAlignment =
        operation.inheritFormatting === false
          ? undefined
          : directParagraphAlignment(expectParagraphAttrs(item.blockNode));
      const directAlignment =
        operation.alignment === undefined
          ? inheritedDirectAlignment
          : (operation.alignment ?? undefined);
      attrs["alignmentFromStyle"] = alignmentFromStyle;
      attrs["alignment"] = directAlignment ?? alignmentFromStyle ?? null;
      const sourceFormatting = attrs["_originalFormatting"];
      const originalFormatting: ParagraphFormatting =
        typeof sourceFormatting === "object" && sourceFormatting !== null
          ? { ...sourceFormatting }
          : {};
      if (operation.styleId !== undefined) {
        if (operation.styleId === null) {
          Reflect.deleteProperty(originalFormatting, "styleId");
        } else {
          originalFormatting.styleId = operation.styleId;
        }
      }
      if (operation.alignment !== undefined || inheritedDirectAlignment !== undefined) {
        if (directAlignment === undefined) {
          Reflect.deleteProperty(originalFormatting, "alignment");
        } else {
          originalFormatting.alignment = directAlignment;
        }
      }
      attrs["_originalFormatting"] =
        Object.keys(originalFormatting).length > 0 ? originalFormatting : null;
    }
    if (isSuggested && suggestionId !== null && paragraphRevisionId !== null) {
      attrs["_suggestedInsert"] = {
        suggestionId,
        revisionId: paragraphRevisionId,
        author,
        date,
        ...(initials ? { initials } : {}),
      };
    }
    nodes.push(item.blockNode.type.create(attrs, content));
  }

  if (producesTrackedChanges && !isSuggested) {
    // Each inserted paragraph owns its paragraph mark. Reusing the anchor's
    // mark fails when another operation inserts a table between the two.
    for (const [index, node] of nodes.entries()) {
      const revisionId = nextRevisionId++;
      nodes[index] = node.type.create(
        {
          ...node.attrs,
          pPrMark: {
            kind: isPairedMove(operation.moveId) ? "moveTo" : "ins",
            info: { id: revisionId, author, date, ...trackedRevisionExtras },
          },
        },
        node.content,
      );
      revisionIds.push(revisionId);
    }
  }

  return { nodes, revisionIds, nextRevisionId };
};

const applyFolioAIEditOperationsInternal = ({
  view,
  snapshot,
  operations,
  mode = "tracked-changes",
  author = "AI",
  initials,
  createCommentId,
  revisionStamp,
  revisionIdSeed,
  wordDiff,
  tableTemplates,
}: ApplyFolioAIEditOperationsInternalOptions): FolioAIEditApplyOutcome => {
  const applied: FolioAIEditAppliedOperation[] = [];
  const skipped: FolioAIEditSkippedOperation[] = [];
  const normalizations: FolioAIEditNormalization[] = [];
  const resolved: ResolvedOperation[] = [];
  const insertionType = view.state.schema.marks["insertion"];
  const deletionType = view.state.schema.marks["deletion"];
  const commentType = view.state.schema.marks["comment"];
  const styleResolver = getDocumentStyleResolver(view.state);
  const alignmentFromStyleForInsertion = (item: ResolvedOperation) => {
    if (item.operation.type !== "insertAfterBlock" && item.operation.type !== "insertBeforeBlock") {
      return undefined;
    }
    const anchorAttrs = expectParagraphAttrs(item.blockNode);
    const styleId =
      item.operation.styleId !== undefined
        ? item.operation.styleId
        : item.operation.inheritFormatting === false
          ? null
          : anchorAttrs.styleId;
    return (
      styleResolver?.resolveParagraphStyle(styleId).paragraphFormatting?.alignment ??
      ((styleId ?? undefined) === (anchorAttrs.styleId ?? undefined)
        ? anchorAttrs.alignmentFromStyle
        : undefined)
    );
  };
  const claimedTableRows = new Set<string>();
  const claimedTableColumns = new Set<string>();

  // `"suggested"` is `"tracked-changes"` plus a provenance stamp, so every
  // tracked-change code path below keys off this rather than an exact
  // `=== "tracked-changes"` check.
  const producesTrackedChanges = mode !== "direct";
  const isSuggested = mode === "suggested";

  if (producesTrackedChanges && (!insertionType || !deletionType)) {
    return {
      applied,
      skipped: operations.map((operation) => ({
        id: operation.id,
        reason: "unsupportedBlock",
      })),
      nextRevisionId: revisionIdSeed ?? revisionStamp?.idSeed ?? revisionIdCursor,
    };
  }

  // Build the live-block indexes once per batch so individual op
  // resolutions don't each re-walk the doc. ParaId-anchored ids are
  // resolved against `liveBlocksByParaId`; the hash bucket is only
  // the fallback for ordinal-encoded snapshot ids.
  const liveBlocks = collectLiveBlocksByHash(view.state.doc);
  const liveBlocksByParaId = collectLiveBlocksByParaId(view.state.doc);

  for (const [index, operation] of operations.entries()) {
    const commentText = getOperationCommentText(operation);
    if (commentText !== undefined && (!commentType || createCommentId === undefined)) {
      skipped.push({ id: operation.id, reason: "unsupportedBlock" });
      continue;
    }

    const resolution = resolveOperation({
      snapshot,
      operation,
      liveBlocks,
      liveBlocksByParaId,
      doc: view.state.doc,
    });
    if (resolution.type === "skip") {
      skipped.push({ id: operation.id, reason: resolution.reason });
      continue;
    }

    const deletion = resolution.operation.tableRowDeletion;
    if (deletion) {
      const rowKey = `${deletion.tablePosition}:${deletion.rowIndex}`;
      if (claimedTableRows.has(rowKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableRows.add(rowKey);
    }

    const columnDeletion = resolution.operation.tableColumnDeletion;
    if (columnDeletion) {
      const columnKey = getTableColumnCoordinateKey(columnDeletion);
      if (claimedTableColumns.has(columnKey)) {
        skipped.push({ id: operation.id, reason: "noopOperation" });
        continue;
      }
      claimedTableColumns.add(columnKey);
    }

    if (
      (operation.type === "insertAfterBlock" || operation.type === "insertBeforeBlock") &&
      LINE_BREAK_PATTERN.test(operation.text)
    ) {
      normalizations.push({
        id: operation.id,
        code: "splitMultilineText",
        paragraphCount: resolution.operation.insertTexts?.length ?? 1,
      });
    }

    const commentId = commentText !== undefined ? createCommentId?.(commentText) : undefined;
    resolved.push({
      ...resolution.operation,
      originalIndex: index,
      ...(commentId !== undefined && { commentId }),
    });
  }

  const tablePlan = planTableMutations(
    resolved.map((item) => ({
      item,
      operationId: item.operation.id,
      target: getTableMutationPlanTarget(item),
    })),
  );
  skipped.push(...tablePlan.skipped);
  const executableResolved = tablePlan.executable;

  if (executableResolved.length === 0) {
    return {
      applied,
      skipped,
      ...(normalizations.length > 0 && { normalizations }),
      nextRevisionId: revisionIdSeed ?? revisionStamp?.idSeed ?? revisionIdCursor,
    };
  }

  // A `moveId` names one relocation. It is a move only when it reaches both
  // halves: `w:moveTo` without its `w:moveFrom` is a relocation from nowhere,
  // and a reader accepting it would see text appear with no source. An
  // unpaired id degrades to a plain insertion or deletion and is reported.
  const moveSideCounts = new Map<string, { from: number; to: number }>();
  for (const { operation } of executableResolved) {
    const moveId =
      operation.type === "deleteBlock" ||
      operation.type === "insertAfterBlock" ||
      operation.type === "insertBeforeBlock"
        ? operation.moveId
        : undefined;
    if (moveId === undefined) {
      continue;
    }
    const counts = moveSideCounts.get(moveId) ?? { from: 0, to: 0 };
    if (operation.type === "deleteBlock") {
      counts.from += 1;
    } else {
      counts.to += 1;
    }
    moveSideCounts.set(moveId, counts);
  }
  const isPairedMove = (moveId: string | undefined): moveId is string => {
    if (moveId === undefined) {
      return false;
    }
    const counts = moveSideCounts.get(moveId);
    return counts?.from === 1 && counts.to === 1;
  };
  for (const { operation } of executableResolved) {
    const moveId =
      operation.type === "deleteBlock" ||
      operation.type === "insertAfterBlock" ||
      operation.type === "insertBeforeBlock"
        ? operation.moveId
        : undefined;
    if (moveId !== undefined && !isPairedMove(moveId)) {
      normalizations.push({ id: operation.id, code: "unpairedMove", moveId });
    }
  }

  let tr = view.state.tr;
  const revisionIdReservation = executableResolved.reduce(
    (total, item) => total + estimateRevisionIdReservation(item),
    0,
  );
  let revisionSeed =
    revisionIdSeed ?? revisionStamp?.idSeed ?? nextRevisionSeed(revisionIdReservation);
  const date = revisionStamp?.date ?? new Date().toISOString();
  const insertedColumnCounts = new Map<string, number>();

  // Sort right-to-left so each tr.insert / tr.delete leaves earlier
  // positions intact. Column insertions run before deletions at the
  // same snapshot coordinate; the deletion path accounts for those
  // inserted columns so it still removes the original target. Other
  // ties use reverse input order so repeated insertions retain their
  // requested sequence.
  const executionOrder = executableResolved.toSorted((left, right) => {
    const leftCellShape = left.tableCellMerge ?? left.tableCellSplit;
    const rightCellShape = right.tableCellMerge ?? right.tableCellSplit;
    if (!leftCellShape && rightCellShape) {
      return -1;
    }
    if (leftCellShape && !rightCellShape) {
      return 1;
    }
    if (leftCellShape && rightCellShape) {
      if (leftCellShape.tablePosition !== rightCellShape.tablePosition) {
        return rightCellShape.tablePosition - leftCellShape.tablePosition;
      }
      if (leftCellShape.rectangle.bottom !== rightCellShape.rectangle.bottom) {
        return rightCellShape.rectangle.bottom - leftCellShape.rectangle.bottom;
      }
      if (leftCellShape.rectangle.right !== rightCellShape.rectangle.right) {
        return rightCellShape.rectangle.right - leftCellShape.rectangle.right;
      }
      return right.originalIndex - left.originalIndex;
    }
    const leftColumn = left.tableColumnInsertion ?? left.tableColumnDeletion;
    const rightColumn = right.tableColumnInsertion ?? right.tableColumnDeletion;
    if (!leftColumn && rightColumn) {
      return -1;
    }
    if (leftColumn && !rightColumn) {
      return 1;
    }
    if (leftColumn && rightColumn) {
      if (leftColumn.tablePosition !== rightColumn.tablePosition) {
        return rightColumn.tablePosition - leftColumn.tablePosition;
      }
      if (leftColumn.columnIndex !== rightColumn.columnIndex) {
        return rightColumn.columnIndex - leftColumn.columnIndex;
      }
      const leftIsInsertion = left.tableColumnInsertion !== undefined;
      const rightIsInsertion = right.tableColumnInsertion !== undefined;
      if (leftIsInsertion !== rightIsInsertion) {
        return leftIsInsertion ? -1 : 1;
      }
      return right.originalIndex - left.originalIndex;
    }
    if (left.from !== right.from) {
      return right.from - left.from;
    }
    return right.originalIndex - left.originalIndex;
  });
  for (let executionIndex = 0; executionIndex < executionOrder.length; executionIndex++) {
    const item = executionOrder[executionIndex];
    if (!item) {
      panic("The operation execution index exceeded the resolved plan", { executionIndex });
    }
    if (isBatchableParagraphInsertion(item, mode)) {
      const insertionRun: ResolvedOperation[] = [item];
      for (let lookahead = executionIndex + 1; lookahead < executionOrder.length; lookahead++) {
        const candidate = executionOrder[lookahead];
        if (
          !candidate ||
          candidate.from !== item.from ||
          !isBatchableParagraphInsertion(candidate, mode)
        ) {
          break;
        }
        insertionRun.push(candidate);
      }
      if (insertionRun.length > 1) {
        const nodeGroups: PMNode[][] = [];
        const runApplied: FolioAIEditAppliedOperation[] = [];
        for (const insertion of insertionRun) {
          const insertionSuggestionId = isSuggested
            ? (insertion.operation.suggestionId ?? insertion.operation.id)
            : null;
          const insertionCommentMark =
            insertion.commentId !== undefined && commentType
              ? commentType.create({ commentId: insertion.commentId })
              : null;
          const built = buildInsertedParagraphs({
            item: insertion,
            schema: view.state.schema,
            alignmentFromStyle: alignmentFromStyleForInsertion(insertion),
            mode,
            author,
            date,
            initials,
            commentMark: insertionCommentMark,
            suggestionId: insertionSuggestionId,
            revisionSeed,
            isPairedMove,
          });
          revisionSeed = built.nextRevisionId;
          nodeGroups.push(built.nodes);
          runApplied.push({
            id: insertion.operation.id,
            ...(insertion.commentId !== undefined && { commentId: insertion.commentId }),
            ...(built.revisionIds[0] !== undefined && {
              revisionId: built.revisionIds[0],
              revisionIds: built.revisionIds,
            }),
            ...(insertionSuggestionId !== null && { suggestionId: insertionSuggestionId }),
          });
        }
        const nodes: PMNode[] = [];
        for (const group of nodeGroups.toReversed()) {
          for (const node of group) {
            nodes.push(node);
          }
        }
        tr = tr.insert(item.from, nodes);
        applied.push(...runApplied);
        executionIndex += insertionRun.length - 1;
        continue;
      }
    }
    const commentMark =
      item.commentId !== undefined && commentType
        ? commentType.create({ commentId: item.commentId })
        : null;

    // Snapshot the transaction's step count so we can detect when an
    // operation produced zero document changes and report it as a
    // skipped no-op instead of a phantom "applied" entry. This caught
    // the silent accept-failure bug where a replaceInBlock on a
    // block with pending tracked changes computed the wrong source
    // text and the diff produced no marks; the panel said "accepted"
    // but the doc was untouched.
    const stepsBefore = tr.steps.length;
    let appliedRevisionIds: number[] | undefined;

    // Suggested mode covers the inline text/format operations plus the block and
    // table row/column structural operations (see
    // `SUGGESTED_SUPPORTED_OPERATION_TYPES`). Every revision they produce — an
    // inline mark, a whole-node `_suggestedInsert` marker, or a suggested
    // `trIns`/`trDel`/`cellMarker` — is removed by the serialization strip.
    // Operations outside the allowlist (comment ops, cell merge/split) report
    // `unsupportedMode` so no unstrippable suggested state can be produced.
    if (isSuggested && !SUGGESTED_SUPPORTED_OPERATION_TYPES.has(item.operation.type)) {
      skipped.push({ id: item.operation.id, reason: "unsupportedMode" });
      continue;
    }

    // Every mark a suggested operation produces is stamped with this id so the
    // host can accept/reject the whole suggestion at once. Falls back to the
    // operation id when the caller does not supply one.
    const suggestionId: string | null = isSuggested
      ? (item.operation.suggestionId ?? item.operation.id)
      : null;

    // Fields merged into every node-attr revision (trIns/trDel/cellMarker) and
    // whole-node `_suggestedInsert` marker this operation produces.
    const trackedRevisionExtras = {
      ...(initials ? { initials } : {}),
      ...(suggestionId !== null ? { provenance: "suggested" as const, suggestionId } : {}),
    };
    // The structural revision a table op writes. Uses the current `revisionSeed`
    // (each table branch consumes it with `revisionSeed++` after applying); one
    // op runs per iteration, so no prior increment shifts this value.
    const structuralRevision: TableStructureRevision | null = producesTrackedChanges
      ? { revisionId: revisionSeed, author, date, ...trackedRevisionExtras }
      : null;

    switch (item.operation.type) {
      case "replaceInBlock":
      case "replaceRange": {
        const revisionIdDelete = revisionSeed++;
        const revisionIdInsert = revisionSeed++;
        const revisionIdBackground = revisionSeed++;
        const stepsBeforeBackgroundClear = tr.steps.length;
        tr = clearReplacementBackground({
          tr,
          schema: view.state.schema,
          from: item.from,
          to: item.to,
          mode,
          revisionId: revisionIdBackground,
          author,
          date,
          initials,
          suggestionId,
        });
        const clearedBackground = tr.steps.length > stepsBeforeBackgroundClear;
        tr = applyTextReplacement({
          tr,
          item,
          mode,
          author,
          date,
          revisionIdDelete,
          revisionIdInsert,
          commentMark,
          suggestionId,
          initials,
          granularity: wordDiff?.granularity ?? "word",
        });
        if (producesTrackedChanges) {
          appliedRevisionIds = [
            revisionIdDelete,
            revisionIdInsert,
            ...(clearedBackground ? [revisionIdBackground] : []),
          ];
        }
        break;
      }
      case "commentOnRange": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "formatRange": {
        if (producesTrackedChanges) {
          const revisionId = revisionSeed++;
          tr = applyTrackedInlineFormatting({
            tr,
            schema: view.state.schema,
            doc: tr.doc,
            from: item.from,
            to: item.to,
            formatting: item.operation.formatting,
            revisionId,
            author,
            date,
            initials,
            suggestionId,
          });
          appliedRevisionIds = [revisionId];
          break;
        }
        tr = applyInlineFormatting({
          tr,
          schema: view.state.schema,
          from: item.from,
          to: item.to,
          formatting: item.operation.formatting,
        });
        break;
      }
      case "replaceBlock": {
        if (!isResolvedReplaceBlockOperation(item)) {
          panic("Resolved replaceBlock operation lost its impact discriminator");
        }
        const { changesStyle, changesText } = REPLACE_BLOCK_IMPACT[item.replaceBlockImpact];
        const revisionIdDelete = revisionSeed;
        const revisionIdInsert = changesText ? revisionSeed + 1 : revisionSeed;
        const revisionIdBackground = changesText ? revisionSeed + 2 : revisionSeed;
        if (changesText) {
          revisionSeed += 3;
        }
        const revisionIdParagraph = producesTrackedChanges && changesStyle ? revisionSeed++ : null;
        // Default to preserving formatting (existing behaviour);
        // when explicitly disabled and we're in direct mode, swap
        // the whole block node for a fresh paragraph that drops
        // all block-level attrs. tracked-changes mode keeps the
        // attrs because the visible diff is text-only.
        if (item.operation.preserveFormatting === false && mode === "direct") {
          const replacement = item.operation.text;
          const paragraphType = view.state.schema.nodes["paragraph"];
          if (paragraphType) {
            const node = paragraphType.create(
              null,
              replacement.length === 0
                ? null
                : buildEmphasisInlineContent(view.state.schema, replacement, []),
            );
            tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
            tr = applyReplaceBlockStyleId({ item, tr, styleResolver }).tr;
            break;
          }
        }
        // Direct mode, formatting preserved (the default): when the replacement
        // carries inline emphasis, rebuild the block node keeping its own attrs
        // so `**bold**` becomes real marks. The tracked-changes path can't carry
        // inline marks through its word-diff redline, so it still strips them;
        // plain replacements fall through to the text-only swap below unchanged.
        if (mode === "direct" && hasInlineEmphasis(item.operation.text)) {
          const node = item.blockNode.type.create(
            { ...item.blockNode.attrs },
            buildEmphasisInlineContent(
              view.state.schema,
              item.operation.text,
              commentMark ? [commentMark] : [],
            ),
          );
          tr = tr.replaceWith(item.blockFrom, item.blockTo, node);
          tr = applyReplaceBlockStyleId({ item, tr, styleResolver }).tr;
          break;
        }
        let clearedBackground = false;
        if (changesText) {
          const stepsBeforeBackgroundClear = tr.steps.length;
          tr = clearReplacementBackground({
            tr,
            schema: view.state.schema,
            from: item.from,
            to: item.to,
            mode,
            revisionId: revisionIdBackground,
            author,
            date,
            initials,
            suggestionId,
          });
          clearedBackground = tr.steps.length > stepsBeforeBackgroundClear;
          tr = applyTextReplacement({
            tr,
            item,
            mode,
            author,
            date,
            revisionIdDelete,
            revisionIdInsert,
            commentMark,
            suggestionId,
            initials,
            granularity: wordDiff?.granularity ?? "word",
          });
        }
        const styleResult = applyReplaceBlockStyleId({
          item,
          tr,
          styleResolver,
          ...(revisionIdParagraph !== null
            ? {
                revisionInfo: {
                  id: revisionIdParagraph,
                  author,
                  date,
                  ...trackedRevisionExtras,
                },
              }
            : {}),
        });
        tr = styleResult.tr;
        if (producesTrackedChanges) {
          appliedRevisionIds = [
            ...(changesText ? [revisionIdDelete, revisionIdInsert] : []),
            ...(clearedBackground ? [revisionIdBackground] : []),
            ...(styleResult.revisionId === null ? [] : [styleResult.revisionId]),
          ];
        }
        break;
      }
      case "insertAfterBlock":
      case "insertBeforeBlock": {
        const insertTexts = item.insertTexts ?? [""];
        const isEmptyInsert = insertTexts.length === 1 && insertTexts[0]?.length === 0;
        if (
          mode === "tracked-changes" &&
          item.operation.pageBreakBefore === true &&
          isEmptyInsert
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }
        const built = buildInsertedParagraphs({
          item,
          schema: view.state.schema,
          alignmentFromStyle: alignmentFromStyleForInsertion(item),
          mode,
          author,
          date,
          initials,
          commentMark,
          suggestionId,
          revisionSeed,
          isPairedMove,
        });
        revisionSeed = built.nextRevisionId;
        if (built.revisionIds.length > 0) {
          appliedRevisionIds = built.revisionIds;
        }
        tr = tr.insert(item.from, built.nodes);
        break;
      }
      case "insertSignatureTable": {
        // Tracked-changes mode has no whole-table-insert primitive in OOXML, so
        // it stays unsupported. Suggested mode CAN carry it: the whole table is
        // flagged `_suggestedInsert` and stripped until accepted (accepting
        // applies it directly — see acceptSuggestion).
        if (mode === "tracked-changes") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedMode",
          });
          continue;
        }

        const signatureTable = buildSignatureTableNode({
          schema: view.state.schema,
          parties: item.operation.parties,
        });
        if (!signatureTable) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        // Direct mode: tables don't carry tracked-change marks — the table
        // structure itself is the insert, and inline insertion marks on the
        // paragraph runs inside would double-up with the structural addition.
        let node = signatureTable;
        if (isSuggested && suggestionId !== null) {
          const revisionId = revisionSeed++;
          node = signatureTable.type.create(
            {
              ...signatureTable.attrs,
              _suggestedInsert: {
                suggestionId,
                revisionId,
                author,
                date,
                ...(initials ? { initials } : {}),
              },
            },
            signatureTable.content,
            signatureTable.marks,
          );
          appliedRevisionIds = [revisionId];
        }
        tr = tr.insert(item.from, node);
        break;
      }
      case "insertTableRow": {
        const insertion = item.tableRowInsertion;
        if (!insertion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const template = tableTemplates?.get(item.operation.id);
        const result = applyTableRowInsertion({
          tr,
          insertion,
          cellTexts: item.operation.cellTexts,
          revision,
          ...(template !== undefined && { template }),
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "insertTableColumn": {
        const insertion = item.tableColumnInsertion;
        if (!insertion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableColumnInsertion({
          tr,
          insertion,
          cellTexts: item.operation.cellTexts,
          revision,
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        const columnKey = getTableColumnCoordinateKey(insertion);
        insertedColumnCounts.set(columnKey, (insertedColumnCounts.get(columnKey) ?? 0) + 1);
        break;
      }
      case "deleteTableColumn": {
        const deletion = item.tableColumnDeletion;
        if (!deletion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const columnKey = getTableColumnCoordinateKey(deletion);
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableColumnDeletion({
          tr,
          deletion,
          insertedColumnCount: insertedColumnCounts.get(columnKey) ?? 0,
          revision,
        });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "mergeTableCells": {
        const merge = item.tableCellMerge;
        const tablePosition = merge ? tr.mapping.map(merge.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !merge ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? mergeTableRectangle({ tr, tablePosition, table, rectangle: merge.rectangle })
            : mergeTrackedVerticalTableCells({
                tr,
                tablePosition,
                table,
                rectangle: merge.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "splitTableCell": {
        const split = item.tableCellSplit;
        const tablePosition = split ? tr.mapping.map(split.tablePosition, 1) : null;
        const table = tablePosition === null ? null : tr.doc.nodeAt(tablePosition);
        if (
          !split ||
          tablePosition === null ||
          !table ||
          table.type.spec["tableRole"] !== "table"
        ) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revisionId = mode === "tracked-changes" ? revisionSeed : null;
        const nextTr =
          revisionId === null
            ? splitTableRectangle({ tr, tablePosition, table, rectangle: split.rectangle })
            : splitTrackedVerticalTableCell({
                tr,
                tablePosition,
                table,
                rectangle: split.rectangle,
                revisionId,
                author,
                date,
              });
        if (!nextTr) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = nextTr;
        if (revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [revisionId];
        }
        break;
      }
      case "deleteTableRow": {
        const deletion = item.tableRowDeletion;
        if (!deletion) {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        const revision: TableStructureRevision | null = structuralRevision;
        const result = applyTableRowDeletion({ tr, deletion, revision });
        if (result.type === "unsupported") {
          skipped.push({
            id: item.operation.id,
            reason: "unsupportedBlock",
          });
          continue;
        }
        tr = result.transaction;
        if (result.revisionId !== null) {
          revisionSeed++;
          appliedRevisionIds = [result.revisionId];
        }
        break;
      }
      case "deleteBlock": {
        if (mode === "direct") {
          // A container has to END with a paragraph: a body, a cell, a header,
          // a note and a text box each do, and one left ending in a table is a
          // package a consumer refuses. Removing the node removes its mark
          // with it, so the last paragraph of a container that nothing else
          // could terminate keeps its place and loses only its content. Where
          // a paragraph precedes it, that one becomes the terminator and this
          // node goes as any other would — which is what the tracked path
          // resolves to as well, its mark deleted one paragraph earlier.
          const at = tr.doc.resolve(item.blockFrom);
          const endsItsContainer = paragraphEndsItsContainer(at, item.blockNode.type.name);
          const leavesAParagraph =
            !endsItsContainer || at.nodeBefore?.type.name === item.blockNode.type.name;
          tr = leavesAParagraph
            ? tr.delete(item.blockFrom, item.blockTo)
            : tr.delete(item.blockFrom + 1, item.blockTo - 1);
          break;
        }

        if (deletionType) {
          const revisionId = revisionSeed++;
          const deletionMark = deletionType.create({
            revisionId,
            author,
            date,
            ...(isPairedMove(item.operation.moveId) && { moveKind: "moveFrom" }),
            ...trackedRevisionExtras,
          });
          tr = tr.addMark(item.from, item.to, deletionMark);
          // The range above spans the block's clean TEXT, so an inline image or
          // field sits outside it whenever it leads or trails the words.
          // Deleting a block deletes what is in it: left unmarked, accepting
          // the deletion kept a paragraph standing around an orphan image.
          const atomRanges = undeletedContentAtomRanges(tr, item.blockFrom);
          for (const { from, to } of atomRanges) {
            tr = tr.addMark(from, to, deletionMark);
          }
          const inlineRevisionApplied = item.from < item.to || atomRanges.length > 0;
          appliedRevisionIds = inlineRevisionApplied ? [revisionId] : [];
          // Deleting a paragraph's words leaves its paragraph mark behind, and
          // an accepted redline then holds a blank line where the paragraph
          // was: a deleted paragraph carries `w:pPr/w:rPr/w:del` as well as
          // its deleted runs.
          //
          // A mark belongs to the paragraph it ends, so the paragraph's own
          // mark is the one that went — except on the paragraph that ends its
          // container. A deleted mark says "join this paragraph with the one
          // after it", and a container's last paragraph has none: a body, a
          // cell, a header, a note and a text box each end with a paragraph
          // that nothing follows. Such a mark states an edit that cannot be
          // carried out, and a consumer reading it refuses the package. That
          // paragraph therefore loses its words and keeps its mark, and the
          // caller that wants the paragraph gone deletes the mark of the one
          // BEFORE it, which merges forward into this carrier.
          //
          // A paragraph before a TABLE still carries its own mark: the table
          // is a following sibling, so the paragraph does not end anything.
          //
          // Read from the document as it stands, which is exact: operations
          // run right to left, so everything after this paragraph has already
          // landed and the paragraph it will end up next to is the one here.
          //
          // A relocation's source break is `w:moveFrom`, not `w:del`: the two
          // resolve alike, and the kind is what tells a reader this end has a
          // matching one elsewhere rather than being a deletion of its own.
          const markPosition = tr.mapping.map(item.blockFrom);
          const markPlace = tr.doc.resolve(markPosition);
          const endsItsContainer = paragraphEndsItsContainer(markPlace, item.blockNode.type.name);
          if (!endsItsContainer && tr.doc.nodeAt(markPosition)?.attrs["pPrMark"] == null) {
            const markRevisionId = revisionSeed++;
            tr = tr.setNodeAttribute(markPosition, "pPrMark", {
              kind: isPairedMove(item.operation.moveId) ? "moveFrom" : "del",
              info: { id: markRevisionId, author, date, ...trackedRevisionExtras },
            });
            appliedRevisionIds = inlineRevisionApplied
              ? [revisionId, markRevisionId]
              : [markRevisionId];
          }
        }
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      case "insertTable": {
        const revision = { revisionId: revisionSeed, author, date, ...trackedRevisionExtras };
        const template = tableTemplates?.get(item.operation.id);
        const table =
          (template === undefined
            ? null
            : tableFromTemplate({
                schema: view.state.schema,
                template,
                ...(producesTrackedChanges && { revision }),
              })) ??
          buildTableNode({
            schema: view.state.schema,
            rows: item.operation.rows,
            ...(producesTrackedChanges && { revision }),
          });
        if (!table) {
          skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
          continue;
        }
        if (producesTrackedChanges) {
          appliedRevisionIds = [revisionSeed++];
        }
        tr = tr.insert(item.from, table);
        markStructuralChange(tr);
        break;
      }
      case "deleteTable": {
        const deleted = item.deletedTable;
        if (!deleted) {
          skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
          continue;
        }
        if (mode === "direct") {
          const position = tr.mapping.map(deleted.position, 1);
          const live = tr.doc.nodeAt(position);
          if (!live || live.type.spec["tableRole"] !== "table") {
            skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
            continue;
          }
          tr = tr.delete(position, position + live.nodeSize);
          markStructuralChange(tr);
          break;
        }
        // Word deletes a table by marking every row deleted; there is no
        // "this table went away" element. Read the rows from the LIVE
        // document: an earlier operation in this batch may have replaced text
        // inside the table, and a row position from before that is stale.
        const livePosition = tr.mapping.map(deleted.position, 1);
        const liveTable = tr.doc.nodeAt(livePosition);
        if (!liveTable || liveTable.type.spec["tableRole"] !== "table") {
          skipped.push({ id: item.operation.id, reason: "unsupportedBlock" });
          continue;
        }
        const revision = { revisionId: revisionSeed++, author, date, ...trackedRevisionExtras };
        const rowPositions: number[] = [];
        let rowOffset = livePosition + 1;
        liveTable.forEach((row) => {
          if (row.type.spec["tableRole"] === "row" && row.attrs["trDel"] == null) {
            rowPositions.push(rowOffset);
          }
          rowOffset += row.nodeSize;
        });
        if (rowPositions.length === 0) {
          skipped.push({ id: item.operation.id, reason: "noopOperation" });
          continue;
        }
        for (const rowPosition of rowPositions.toReversed()) {
          tr = tr.setNodeAttribute(rowPosition, "trDel", revision);
          // The row marker alone is not the deletion: a consumer reading only
          // run-level revisions keeps the text on accept. Same rule the
          // row-level deletion follows.
          markTableRowContent({ tr, rowPosition, kind: "deletion", revision });
        }
        appliedRevisionIds = [revision.revisionId];
        markStructuralChange(tr);
        break;
      }
      case "setBlockParagraphProperties": {
        // Rewriting a node's attributes replaces ALL of them, so they have to
        // be read from the document as it stands rather than from the
        // resolution-time snapshot. Another operation of the same batch may
        // already have written to this paragraph — a merge puts its deleted
        // paragraph mark in `pPrMark`, and the batch runs right to left, so
        // the mark lands before the properties do. Rebuilding the node from
        // the stale attributes dropped that mark without a trace: the
        // properties applied, the merge silently did not.
        const blockPosition = tr.mapping.map(item.blockFrom);
        const liveBlock = tr.doc.nodeAt(blockPosition) ?? item.blockNode;
        const resolvedAlignmentFromStyle =
          item.operation.properties.styleId === undefined
            ? undefined
            : styleResolver?.resolveParagraphStyle(item.operation.properties.styleId)
                .paragraphFormatting?.alignment;
        const patch = paragraphPropertiesPatch({
          node: liveBlock,
          properties: item.operation.properties,
          resolvedAlignmentFromStyle,
        });
        if (patch === null) {
          skipped.push({ id: item.operation.id, reason: "noopOperation" });
          continue;
        }
        if (mode === "direct") {
          tr = tr.setNodeMarkup(blockPosition, undefined, { ...liveBlock.attrs, ...patch });
          break;
        }
        // Word stores the COMPLETE old pPr inside `w:pPrChange`, so rejecting
        // restores the properties wholesale within that scope. Storing only
        // the keys this operation touched would leave a reject unable to tell
        // "the change did not set this" from "the change cleared it".
        const revisionId = revisionSeed++;
        const change: ParagraphPropertyChangeAttrs = {
          type: "paragraphPropertyChange",
          info: { id: revisionId, author, date, ...trackedRevisionExtras },
          previousFormatting: paragraphPropertiesSnapshot(liveBlock),
        };
        const existing = expectParagraphAttrs(liveBlock)._propertyChanges;
        tr = tr.setNodeMarkup(blockPosition, undefined, {
          ...liveBlock.attrs,
          ...patch,
          _propertyChanges: [...(Array.isArray(existing) ? existing : []), change],
        });
        appliedRevisionIds = [revisionId];
        break;
      }
      case "splitBlock": {
        if (mode === "direct") {
          if (item.to > item.from) {
            tr = tr.delete(item.from, item.to);
          }
          tr = tr.split(item.from);
          break;
        }
        const revisionIdMark = revisionSeed++;
        const info = { id: revisionIdMark, author, date, ...trackedRevisionExtras };
        appliedRevisionIds = [revisionIdMark];
        if (item.to > item.from && deletionType) {
          const revisionIdSeparator = revisionSeed++;
          tr = tr.addMark(
            item.from,
            item.to,
            deletionType.create({
              revisionId: revisionIdSeparator,
              author,
              date,
              ...trackedRevisionExtras,
            }),
          );
          appliedRevisionIds = [revisionIdMark, revisionIdSeparator];
        }
        tr = tr.split(item.from);
        // The mark goes on the FIRST half: the break belongs to the paragraph
        // it now ends, and a reader rejecting it closes that paragraph back
        // over the second half.
        tr = tr.setNodeAttribute(item.blockFrom, "pPrMark", { kind: "ins", info });
        break;
      }
      case "mergeBlockWithNext": {
        const separator = item.operation.separator ?? "";
        const insertAt = item.blockTo - 1;
        if (mode === "direct") {
          if (separator.length > 0) {
            tr = tr.insertText(separator, insertAt);
          }
          tr = tr.join(item.blockTo + separator.length);
          break;
        }
        const revisionIdMark = revisionSeed++;
        appliedRevisionIds = [revisionIdMark];
        if (separator.length > 0 && insertionType) {
          const revisionIdSeparator = revisionSeed++;
          tr = tr.insertText(separator, insertAt);
          tr = tr.addMark(
            insertAt,
            insertAt + separator.length,
            insertionType.create({
              revisionId: revisionIdSeparator,
              author,
              date,
              ...trackedRevisionExtras,
            }),
          );
          appliedRevisionIds = [revisionIdMark, revisionIdSeparator];
        }
        tr = tr.setNodeAttribute(item.blockFrom, "pPrMark", {
          kind: "del",
          info: { id: revisionIdMark, author, date, ...trackedRevisionExtras },
        });
        break;
      }
      case "commentOnBlock": {
        if (commentMark) {
          tr = tr.addMark(item.from, item.to, commentMark);
        }
        break;
      }
      default:
        break;
    }

    // commentOnBlock is intentionally a doc-mutation-free op
    // (adds a mark; that DOES count as a step) but covers the
    // edge case where the comment mark is missing. Treat any op
    // that emitted zero transaction steps as a no-op skip.
    if (tr.steps.length === stepsBefore) {
      skipped.push({ id: item.operation.id, reason: "noopOperation" });
      continue;
    }

    // Surface the primary id (first one) on the legacy `revisionId`
    // field so callers that just need a stable scroll/visual
    // reference keep working. The full set is on `revisionIds` for
    // accept/reject paths that must clear every mark belonging to
    // this op.
    applied.push({
      id: item.operation.id,
      ...(item.commentId !== undefined && { commentId: item.commentId }),
      ...(appliedRevisionIds !== undefined &&
        appliedRevisionIds[0] !== undefined && {
          revisionId: appliedRevisionIds[0],
          revisionIds: appliedRevisionIds,
        }),
      ...(suggestionId !== null && { suggestionId }),
    });
  }

  if (tr.docChanged) {
    const batchRevisionIds = new Set(applied.flatMap(({ revisionIds }) => revisionIds ?? []));
    tr = withoutRevisionsOnZeroWidthAnchors(tr, batchRevisionIds);
    const rotated = withRotatedAddedFinalBreaks({
      tr,
      batchRevisionIds,
      revisionSeed,
      author,
      date,
      initials,
    });
    tr = rotated.transaction;
    revisionSeed = rotated.nextRevisionId;
    if (revisionStamp) {
      // Paragraphs this batch creates get their `w14:paraId` from the
      // allocator plugin, which is random by default. A stamped batch has
      // promised its caller a reproducible package, so the ids must be
      // derived from the stamp too.
      requestDeterministicParaIds(tr, `${revisionStamp.date}:${String(revisionStamp.idSeed)}`);
    }
    view.dispatch(tr);
  }

  return {
    applied,
    skipped,
    ...(normalizations.length > 0 && { normalizations }),
    nextRevisionId: revisionSeed,
  };
};

export const applyFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyOutcome => applyFolioAIEditOperationsInternal(options);

export const previewFolioAIEditOperations = (
  options: ApplyFolioAIEditOperationsOptions,
): FolioAIEditApplyOutcome => {
  const { view, createCommentId, ...applyOptions } = options;
  let previewCommentId = -1;
  const previewView: FolioAIEditView = {
    state: view.state,
    dispatch: (transaction) => {
      previewView.state = previewView.state.apply(transaction);
    },
  };
  const result = applyFolioAIEditOperationsInternal({
    ...applyOptions,
    view: previewView,
    ...(createCommentId !== undefined && {
      createCommentId: () => previewCommentId--,
    }),
    revisionIdSeed: -1_000_000_000,
  });
  return {
    applied: result.applied.map(({ id }) => ({ id })),
    skipped: result.skipped,
    ...(result.normalizations !== undefined && { normalizations: result.normalizations }),
    // A preview allocates from a sentinel range and commits nothing, so the
    // next id is still the one the batch would have started from.
    nextRevisionId: options.revisionStamp?.idSeed ?? revisionIdCursor,
  };
};

type TextReplacementOptions = {
  tr: Transaction;
  item: ResolvedOperation;
  mode: FolioAIEditApplyMode;
  author: string;
  date: string;
  /**
   * Distinct revision id used for the deletion-side marks.
   * fromProseDoc treats a single revisionId carrying BOTH ins and
   * del marks as a Word "moveTo/moveFrom" pair on serialization,
   * which is wrong for an AI replace — so the engine allocates one
   * id for the deletion side and a separate one for the insertion
   * side of the same operation.
   */
  revisionIdDelete: number;
  /** Distinct revision id used for the insertion-side marks. */
  revisionIdInsert: number;
  commentMark: Mark | null;
  /** Non-null stamps every produced insertion/deletion mark as a suggestion. */
  suggestionId?: string | null;
  /** Optional author initials stamped on the produced marks. */
  initials?: string | undefined;
  /** Token size the redline is cut at. */
  granularity: WordDiffGranularity;
};

const applyTextReplacement = ({
  tr,
  item,
  mode,
  author,
  date,
  revisionIdDelete,
  revisionIdInsert,
  commentMark,
  suggestionId = null,
  initials,
  granularity,
}: TextReplacementOptions): Transaction => {
  let nextTr = tr;
  const replacement = stripInlineEmphasisMarkers(
    (() => {
      if (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") {
        return item.operation.replace;
      }
      if (item.operation.type === "replaceBlock") {
        return item.operation.text;
      }
      return "";
    })(),
  );

  if (mode === "direct") {
    // Partial-block replacement carrying inline emphasis: rebuild the matched
    // range as real bold/italic runs instead of stripping the markers to plain
    // text. Full-block replaceBlock in direct mode is intercepted earlier and
    // never reaches here with emphasis; the tracked-changes path below still
    // strips, since its word-diff redline can't carry inline marks.
    if (
      (item.operation.type === "replaceInBlock" || item.operation.type === "replaceRange") &&
      hasInlineEmphasis(item.operation.replace)
    ) {
      const content = buildEmphasisInlineContent(
        nextTr.doc.type.schema,
        item.operation.replace,
        commentMark ? [commentMark] : [],
      );
      return nextTr.replaceWith(item.from, item.to, content);
    }
    nextTr = nextTr.insertText(replacement, item.from, item.to);
    if (commentMark && replacement.length > 0) {
      nextTr = nextTr.addMark(item.from, item.from + replacement.length, commentMark);
    }
    return nextTr;
  }

  const insertionType = nextTr.doc.type.schema.marks["insertion"];
  const deletionType = nextTr.doc.type.schema.marks["deletion"];
  const suggestionAttrs = suggestionId === null ? {} : { provenance: "suggested", suggestionId };
  const initialsAttr = initials ? { initials } : {};
  const delAttrs = {
    revisionId: revisionIdDelete,
    author,
    date,
    ...initialsAttr,
    ...suggestionAttrs,
  };
  const insAttrs = {
    revisionId: revisionIdInsert,
    author,
    date,
    ...initialsAttr,
    ...suggestionAttrs,
  };

  // Word-level diff is only safe when the source range maps to PM
  // positions losslessly. The block must have no atomic inline
  // nodes (hard breaks, inline images) — those break textContent /
  // PM-position alignment in ways the offsets array can't resolve.
  //
  // Existing tracked-change marks ARE handled here: we walk PM
  // positions through `cleanBlock.offsets[]` (built from the
  // post-tracked-changes view), so each clean-text char anchors at
  // the right live position even when the block has pending
  // deletion runs interleaved between surviving chars. Naively
  // accumulating `cursor += seg.text.length` would skip the gap
  // introduced by those deletion runs and write marks onto the
  // wrong live characters — the silent accept-failure bug.
  const blockHasOnlyTextChildren =
    item.blockNode.content.size === item.blockNode.textContent.length;
  const cleanBlock = blockHasOnlyTextChildren
    ? buildCleanBlockText(item.blockNode, item.blockFrom)
    : null;
  let sourceText: string | null = null;
  let sourceCleanStart = 0;
  if (cleanBlock !== null) {
    if (item.operation.type === "replaceInBlock") {
      sourceText = item.operation.find;
      sourceCleanStart = cleanBlock.text.indexOf(item.operation.find);
      if (sourceCleanStart === -1) {
        sourceText = null;
      }
    } else if (item.operation.type === "replaceRange") {
      sourceCleanStart = item.operation.range.startOffset;
      sourceText = cleanBlock.text.slice(sourceCleanStart, item.operation.range.endOffset);
    } else if (item.operation.type === "replaceBlock") {
      sourceText = cleanBlock.text;
      sourceCleanStart = 0;
    }
  }

  if (sourceText !== null && cleanBlock !== null) {
    const segments = diffWordSegments(sourceText, replacement, { granularity });
    const offsets = cleanBlock.offsets;
    const offsetAt = (cleanOffset: number): number | null => offsets[cleanOffset] ?? null;

    type Step =
      | { kind: "del"; from: number; to: number }
      | { kind: "ins"; at: number; text: string };
    const steps: Step[] = [];
    // Cursor walks SOURCE-text offsets (within `sourceText`), then
    // we translate to PM positions through offsets[]. This survives
    // gaps caused by existing deletion-marked runs in the live doc.
    let cursor = 0;
    let allPositionsResolved = true;
    for (const seg of segments) {
      if (seg.type === "equal") {
        cursor += seg.text.length;
        continue;
      }
      if (seg.type === "del") {
        const pmFrom = offsetAt(sourceCleanStart + cursor);
        const pmTo = offsetAt(sourceCleanStart + cursor + seg.text.length);
        if (pmFrom === null || pmTo === null) {
          allPositionsResolved = false;
          break;
        }
        steps.push({ kind: "del", from: pmFrom, to: pmTo });
        cursor += seg.text.length;
        continue;
      }
      const pmAt = offsetAt(sourceCleanStart + cursor);
      if (pmAt === null) {
        allPositionsResolved = false;
        break;
      }
      steps.push({ kind: "ins", at: pmAt, text: seg.text });
    }

    if (allPositionsResolved) {
      // Apply right-to-left so earlier steps' source positions stay
      // valid after later steps mutate the doc.
      for (const step of steps.toReversed()) {
        if (step.kind === "del" && deletionType) {
          nextTr = nextTr.addMark(step.from, step.to, deletionType.create(delAttrs));
          if (commentMark) {
            nextTr = nextTr.addMark(step.from, step.to, commentMark);
          }
          continue;
        }
        if (step.kind === "ins" && insertionType) {
          nextTr = nextTr.insertText(step.text, step.at, step.at);
          nextTr = nextTr.addMark(
            step.at,
            step.at + step.text.length,
            insertionType.create(insAttrs),
          );
          if (commentMark) {
            nextTr = nextTr.addMark(step.at, step.at + step.text.length, commentMark);
          }
        }
      }
      return nextTr;
    }
    // else fall through to the coarse del+ins path below: the
    // offsets array didn't cover one of our boundaries, which only
    // happens for edge cases at the trailing block boundary.
  }

  if (replacement.length > 0 && insertionType) {
    nextTr = nextTr.insertText(replacement, item.to, item.to);
    nextTr = nextTr.addMark(item.to, item.to + replacement.length, insertionType.create(insAttrs));
    if (commentMark) {
      nextTr = nextTr.addMark(item.to, item.to + replacement.length, commentMark);
    }
  }

  if (item.to > item.from && deletionType) {
    nextTr = nextTr.addMark(item.from, item.to, deletionType.create(delAttrs));
    if (commentMark && replacement.length === 0) {
      nextTr = nextTr.addMark(item.from, item.to, commentMark);
    }
  }

  return nextTr;
};

type WithoutResolutionMetadata<T> = T extends ResolvedOperation
  ? Omit<T, "originalIndex" | "commentId">
  : never;
type ResolvedBase = WithoutResolutionMetadata<ResolvedOperation>;

type ResolveOperationArgs = {
  snapshot: FolioAIEditSnapshot;
  operation: FolioAIEditOperation;
  liveBlocks: Map<string, LiveBlockEntry[]>;
  liveBlocksByParaId: Map<string, LiveBlockEntry>;
  doc: PMNode;
};

type ResolveStableBlockArgs = Omit<ResolveOperationArgs, "operation" | "doc"> & {
  blockId: string;
};

type StableBlockResolution = {
  type: "resolved";
  blockNode: PMNode;
  blockFrom: number;
  blockTo: number;
  cleanBlock: ReturnType<typeof buildCleanBlockText>;
  currentText: string;
  currentTextHash: string;
};

const resolveStableBlock = ({
  snapshot,
  blockId,
  liveBlocks,
  liveBlocksByParaId,
}: ResolveStableBlockArgs): StableBlockResolution | OperationResolutionSkip => {
  const anchor = snapshot.anchors[blockId];
  if (!anchor) {
    return { type: "skip", reason: "missingBlock" };
  }

  const encodedParaId = getFolioParaIdFromBlockId(blockId);
  let live: LiveBlockEntry | undefined;
  if (encodedParaId !== null) {
    live = liveBlocksByParaId.get(encodedParaId);
    if (!live) {
      return { type: "skip", reason: "missingBlock" };
    }
  } else {
    const ordinal = ordinalAmongSameHash(snapshot, blockId);
    if (ordinal < 0) {
      return { type: "skip", reason: "missingBlock" };
    }
    live = liveBlocks.get(anchor.textHash)?.[ordinal];
  }
  if (!live || !live.node.isTextblock) {
    return { type: "skip", reason: "changedBlock" };
  }

  const cleanBlock = buildCleanBlockText(live.node, live.from);
  const currentText = cleanBlock.text;
  const currentTextHash = hashFolioAIBlockText(normalizeFolioAIBlockText(currentText));
  if (currentTextHash !== anchor.textHash) {
    return { type: "skip", reason: "changedBlock" };
  }
  return {
    type: "resolved",
    blockNode: live.node,
    blockFrom: live.from,
    blockTo: live.to,
    cleanBlock,
    currentText,
    currentTextHash,
  };
};

const resolveOperation = ({
  snapshot,
  operation,
  liveBlocks,
  liveBlocksByParaId,
  doc,
}: ResolveOperationArgs):
  | { type: "resolved"; operation: ResolvedBase }
  | OperationResolutionSkip => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
      ? operation.range.blockId
      : operation.blockId;
  const primaryBlock = resolveStableBlock({
    snapshot,
    blockId,
    liveBlocks,
    liveBlocksByParaId,
  });
  if (primaryBlock.type === "skip") {
    return primaryBlock;
  }
  const { blockNode, blockFrom, blockTo, cleanBlock, currentText, currentTextHash } = primaryBlock;
  if (
    operation.precondition !== undefined &&
    currentTextHash !== operation.precondition.blockTextHash
  ) {
    return { type: "skip", reason: "preconditionFailed" };
  }

  if (
    operation.type === "replaceRange" ||
    operation.type === "commentOnRange" ||
    operation.type === "formatRange"
  ) {
    const { startOffset, endOffset, selectedTextHash } = operation.range;
    const from = cleanBlock.offsets[startOffset];
    const to = cleanBlock.offsets[endOffset];
    if (from === undefined || to === undefined) {
      return { type: "skip", reason: "staleRange" };
    }
    const selectedText = currentText.slice(startOffset, endOffset);
    if (hashFolioAIBlockText(selectedText) !== selectedTextHash) {
      return { type: "skip", reason: "staleRange" };
    }
    if (operation.type === "replaceRange" && selectedText === operation.replace) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from,
        to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertAfterBlock" || operation.type === "insertBeforeBlock") {
    // An empty `text` inserts a blank paragraph. That is a real edit — a blank
    // line between two clauses is something a document says — and it is the
    // only way to write one, so it is not refused. An insertion always changes
    // the document: there is no empty insert that would leave it as it was.
    //
    // If the anchor lives inside a `tableCell`, the model meant
    // "place the new block adjacent to the table", not "stuff it
    // into the cell". Override the insertion bounds to the table's
    // outer boundary so the synthesized sibling lands as a peer of
    // the table at doc level.
    const tableBoundary = findOutermostTableBoundary(doc, blockFrom);
    const isInsertAfter = operation.type === "insertAfterBlock";
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = isInsertAfter ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = isInsertAfter ? blockTo : blockFrom;
    }
    const insertTexts = LINE_BREAK_PATTERN.test(operation.text)
      ? splitInsertParagraphTexts(operation.text)
      : [operation.text];
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
        insertTexts,
      },
    };
  }

  if (operation.type === "insertSignatureTable") {
    if (operation.parties.length === 0) {
      return { type: "skip", reason: "emptyOperation" };
    }
    const position = operation.position ?? "after";
    const tableBoundary = findOutermostTableBoundary(doc, blockFrom);
    let insertFrom: number;
    if (tableBoundary) {
      insertFrom = position === "after" ? tableBoundary.after : tableBoundary.before;
    } else {
      insertFrom = position === "after" ? blockTo : blockFrom;
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertFrom,
        to: insertFrom,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "insertTableRow") {
    const position = operation.position ?? "after";
    const insertion = findTableRowInsertion({ doc, blockFrom, position });
    // Sized against the table's COLUMN count, which is what a caller reading
    // the table counts. `cells.length` is smaller whenever a cell spans
    // columns or a row above spans down into this one, and refusing on that
    // would refuse a row the table can perfectly well hold.
    if (!insertion || (operation.cellTexts?.length ?? 0) > insertion.columnCount) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: insertion.rowPosition,
        to: insertion.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowInsertion: insertion,
      },
    };
  }

  if (operation.type === "deleteTableRow") {
    const target = findEnclosingTableRow(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.rowPosition,
        to: target.rowPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableRowDeletion: {
          tableStart: target.tableStart,
          tablePosition: target.tablePosition,
          rowIndex: target.rowIndex,
          rowPosition: target.rowPosition,
        },
      },
    };
  }

  if (operation.type === "insertTableColumn") {
    const position = operation.position ?? "after";
    const insertion = findTableColumnInsertion({ doc, blockFrom, position });
    if (!insertion) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const { boundaryPosition, ...tableColumnInsertion } = insertion;
    return {
      type: "resolved",
      operation: {
        operation,
        from: boundaryPosition,
        to: boundaryPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnInsertion,
      },
    };
  }

  if (operation.type === "deleteTableColumn") {
    const target = findEnclosingTableCell(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.cellPosition,
        to: target.cellPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableColumnDeletion: {
          tablePosition: target.tablePosition,
          columnIndex: target.leftColumnIndex,
        },
      },
    };
  }

  if (operation.type === "mergeTableCells") {
    const startCell = findEnclosingTableCell(doc, blockFrom);
    if (!startCell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const table = doc.nodeAt(startCell.tablePosition);
    if (!table || table.type.spec["tableRole"] !== "table") {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const map = TableMap.get(table);
    let rectangle: TableRectangle;
    let endCellPosition: number;
    let endCellEndPosition: number;
    if (operation.rowCount !== undefined) {
      rectangle = {
        left: startCell.leftColumnIndex,
        top: startCell.topRowIndex,
        right: startCell.rightColumnIndex,
        bottom: startCell.topRowIndex + operation.rowCount,
      };
      if (
        rectangle.right - rectangle.left !== 1 ||
        rectangle.bottom > map.height ||
        operation.rowCount < 2
      ) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      const endRelativePosition = map.map[(rectangle.bottom - 1) * map.width + rectangle.left];
      const endCell = endRelativePosition === undefined ? null : table.nodeAt(endRelativePosition);
      if (!endCell || endRelativePosition === undefined) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      endCellPosition = startCell.tablePosition + 1 + endRelativePosition;
      endCellEndPosition = endCellPosition + endCell.nodeSize;
    } else {
      const endTarget = resolveStableBlock({
        snapshot,
        blockId: operation.endBlockId,
        liveBlocks,
        liveBlocksByParaId,
      });
      if (endTarget.type === "skip") {
        return endTarget;
      }
      const endCell = findEnclosingTableCell(doc, endTarget.blockFrom);
      if (!endCell || startCell.tablePosition !== endCell.tablePosition) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      rectangle = {
        left: Math.min(startCell.leftColumnIndex, endCell.leftColumnIndex),
        top: Math.min(startCell.topRowIndex, endCell.topRowIndex),
        right: Math.max(startCell.rightColumnIndex, endCell.rightColumnIndex),
        bottom: Math.max(startCell.bottomRowIndex, endCell.bottomRowIndex),
      };
      endCellPosition = endCell.cellPosition;
      endCellEndPosition = endCell.cellEndPosition;
    }
    if (
      (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) ||
      tableRectangleCutsMergedCell(map, rectangle)
    ) {
      return {
        type: "skip",
        reason:
          rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1
            ? "noopOperation"
            : "unsupportedBlock",
      };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: Math.min(startCell.cellPosition, endCellPosition),
        to: Math.max(startCell.cellEndPosition, endCellEndPosition),
        blockFrom,
        blockTo,
        blockNode,
        tableCellMerge: {
          tablePosition: startCell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "splitTableCell") {
    const cell = findEnclosingTableCell(doc, blockFrom);
    if (!cell) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    const rectangle = {
      left: cell.leftColumnIndex,
      top: cell.topRowIndex,
      right: cell.rightColumnIndex,
      bottom: cell.bottomRowIndex,
    };
    if (rectangle.right - rectangle.left === 1 && rectangle.bottom - rectangle.top === 1) {
      return { type: "skip", reason: "noopOperation" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: cell.cellPosition,
        to: cell.cellEndPosition,
        blockFrom,
        blockTo,
        blockNode,
        tableCellSplit: {
          tablePosition: cell.tablePosition,
          rectangle,
        },
      },
    };
  }

  if (operation.type === "insertTable") {
    // A whole table is a document-level peer, like every other block
    // insertion: `insertTableRow` is the operation for growing one in place.
    const boundary = findOutermostTableBoundary(doc, blockFrom) ?? {
      before: blockFrom,
      after: blockTo,
    };
    const insertFrom =
      (operation.position ?? "after") === "after" ? boundary.after : boundary.before;
    return {
      type: "resolved",
      operation: { operation, from: insertFrom, to: insertFrom, blockFrom, blockTo, blockNode },
    };
  }

  if (operation.type === "deleteTable") {
    const target = findEnclosingTableRow(doc, blockFrom);
    if (!target) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: target.tablePosition,
        to: target.tablePosition + target.table.nodeSize,
        blockFrom,
        blockTo,
        blockNode,
        deletedTable: { position: target.tablePosition, node: target.table },
      },
    };
  }

  if (operation.type === "setBlockParagraphProperties") {
    return {
      type: "resolved",
      operation: { operation, from: blockFrom, to: blockTo, blockFrom, blockTo, blockNode },
    };
  }

  if (operation.type === "splitBlock") {
    // A split at either end of the block moves no words and produces an empty
    // paragraph; the caller meant an insertion.
    const separator = operation.separator ?? "";
    const at = cleanBlock.offsets[operation.offset];
    const after = cleanBlock.offsets[operation.offset + separator.length];
    if (
      at === undefined ||
      after === undefined ||
      operation.offset === 0 ||
      operation.offset + separator.length >= currentText.length
    ) {
      return { type: "skip", reason: "staleRange" };
    }
    if (currentText.slice(operation.offset, operation.offset + separator.length) !== separator) {
      return { type: "skip", reason: "staleRange" };
    }
    if (!canSplit(doc, at)) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: { operation, from: at, to: after, blockFrom, blockTo, blockNode },
    };
  }

  if (operation.type === "mergeBlockWithNext") {
    // `canJoin` is what keeps a deleted paragraph mark off the last paragraph
    // of a table cell, and off the last paragraph of the story: there is no
    // sibling to join with, so accepting the revision could not do what the
    // mark says it does.
    if (!canJoin(doc, blockTo)) {
      return { type: "skip", reason: "unsupportedBlock" };
    }
    return {
      type: "resolved",
      operation: { operation, from: blockTo, to: blockTo, blockFrom, blockTo, blockNode },
    };
  }

  if (operation.type === "deleteBlock" || operation.type === "replaceBlock") {
    const replaceChangesStyle =
      operation.type === "replaceBlock" &&
      operation.styleId !== undefined &&
      operation.styleId !== (expectParagraphAttrs(blockNode).styleId ?? null);
    const range = getTextRangeFromCleanBlock(cleanBlock);
    if (!range) {
      const insertionPoint = cleanBlock.offsets.at(0);
      if (insertionPoint === undefined) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      // A blank paragraph has no run range to mark. Deleting one, or writing
      // into one, is still an edit: what changes is the paragraph itself, and
      // its MARK is where the revision goes. An empty range resolves both —
      // `replaceBlock` inserts at it, `deleteBlock` marks nothing inline and
      // lets the paragraph-mark deletion carry the change.
      const resolvesOnABlank =
        operation.type === "deleteBlock"
          ? true
          : currentText.length === 0 && (operation.text.length > 0 || replaceChangesStyle);
      if (!resolvesOnABlank) {
        return { type: "skip", reason: "unsupportedBlock" };
      }
      if (operation.type === "replaceBlock") {
        return {
          type: "resolved",
          operation: {
            operation,
            from: insertionPoint,
            to: insertionPoint,
            blockFrom,
            blockTo,
            blockNode,
            replaceBlockImpact: resolveReplaceBlockImpact({
              changesText: operation.text !== currentText,
              changesStyle: replaceChangesStyle,
            }),
          },
        };
      }
      return {
        type: "resolved",
        operation: {
          operation,
          from: insertionPoint,
          to: insertionPoint,
          blockFrom,
          blockTo,
          blockNode,
        },
      };
    }
    // The model occasionally emits replaceBlock with text identical
    // to the live block's clean text — usually as a side effect of
    // running through a "review every block" pass. Skip so the
    // panel doesn't show an empty redline (verified in dev-tools
    // trace where one such op surfaced as "Prodávající 3 →
    // Prodávající 3").
    if (operation.type === "replaceBlock" && operation.text === currentText) {
      if (!replaceChangesStyle) {
        return { type: "skip", reason: "noopOperation" };
      }
    }
    if (operation.type === "replaceBlock") {
      return {
        type: "resolved",
        operation: {
          operation,
          from: range.from,
          to: range.to,
          blockFrom,
          blockTo,
          blockNode,
          replaceBlockImpact: resolveReplaceBlockImpact({
            changesText: operation.text !== currentText,
            changesStyle: replaceChangesStyle,
          }),
        },
      };
    }
    return {
      type: "resolved",
      operation: {
        operation,
        from: range.from,
        to: range.to,
        blockFrom,
        blockTo,
        blockNode,
      },
    };
  }

  if (operation.type === "replaceInBlock" && operation.find === operation.replace) {
    return { type: "skip", reason: "noopOperation" };
  }

  const quote = getOperationQuote(operation);
  const range = resolveTextInCleanBlock(cleanBlock, quote || currentText);
  if (range.type !== "resolved") {
    return range;
  }

  return {
    type: "resolved",
    operation: {
      operation,
      from: range.from,
      to: range.to,
      blockFrom,
      blockTo,
      blockNode,
    },
  };
};

type OperationResolutionSkip = { type: "skip"; reason: FolioAIEditSkipReason };

const resolveTextInCleanBlock = (
  cleanBlock: { text: string; offsets: number[] },
  find: string,
):
  | { type: "resolved"; from: number; to: number }
  | { type: "skip"; reason: FolioAIEditSkipReason } => {
  if (find.length === 0) {
    return { type: "skip", reason: "emptyOperation" };
  }

  const { text, offsets } = cleanBlock;
  const firstIndex = text.indexOf(find);
  if (firstIndex === -1) {
    return { type: "skip", reason: "missingFind" };
  }
  if (text.includes(find, firstIndex + 1)) {
    return { type: "skip", reason: "ambiguousFind" };
  }

  const from = offsets[firstIndex];
  const to = offsets[firstIndex + find.length];
  if (from === undefined || to === undefined) {
    return { type: "skip", reason: "unsupportedBlock" };
  }

  return { type: "resolved", from, to };
};

const getTextRangeFromCleanBlock = (cleanBlock: {
  text: string;
  offsets: number[];
}): { from: number; to: number } | null => {
  if (cleanBlock.text.length === 0) {
    return null;
  }
  const from = cleanBlock.offsets[0];
  const to = cleanBlock.offsets[cleanBlock.text.length];
  if (from === undefined || to === undefined) {
    return null;
  }
  return { from, to };
};

const getOperationCommentText = (operation: FolioAIEditOperation): string | undefined =>
  "comment" in operation ? operation.comment?.text : undefined;

const getOperationQuote = (operation: FolioAIEditOperation): string | undefined => {
  if (operation.type === "replaceInBlock") {
    return operation.find;
  }
  if (operation.type === "commentOnBlock") {
    return operation.quote;
  }
  return undefined;
};
