import { panic, Result } from "better-result";

import { INSERTION_SEAM_POLICIES } from "../model/content";
import {
  type AbstractNumbering,
  type BlockContent,
  type BreakContent,
  type Document,
  MAX_REVISION_ID,
  type NumberingInstance,
  type Paragraph,
  type ParagraphContent,
  type TabContent,
  type Table,
  type TableRow,
  type TextFormatting,
} from "../model/document";
import { isNumberingLevel } from "../model/numberingLevel";
import {
  isNumberingReference,
  NO_NUMBERING_NUM_ID,
  paragraphNumberingReference,
} from "../model/paragraphNumbering";
import {
  type AppliedDocumentOp,
  applyDocumentOp,
  applyDocumentOps,
} from "./apply";
import {
  blockListAt,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import { packageIdOccupancy } from "./editorIntentCensus";
import { combineEdits } from "./edits";
import { structurallyEqual } from "./equality";
import { resolveGap } from "./gaps";
/** One editor intent, compiled to direct or tracked document operations. */
import {
  compileHyperlinkIntent,
  type HyperlinkEditorIntent,
} from "./hyperlinkIntent";
import {
  formatParaId,
  freshLeasedIds,
  ID_KIND_LIMITS,
  ID_LEASE_KINDS,
  type IdLease,
  type IdLeaseKind,
} from "./idLease";
import {
  IDENTITY_SPACES,
  idKey,
  packageIdentityKeys,
  packageParagraphIds,
  packageReservedIdentityKeys,
} from "./ids";
import { deleteBetween, gapAfterInserted } from "./inline";
import {
  alikeDepth,
  asParagraphContent,
  childNodes,
  compareGaps,
  defaultInsertionGap,
  type Gap,
  type InlineNode,
  isCommentAnchor,
  isEmptyRecord,
  isParagraphContent,
  leafSpans,
  partitionContent,
  rebuildNode,
  spanningRecords,
  zeroWidthLeavesAt,
} from "./leaves";
import { cloneModel } from "./modelClone";
import {
  isRangeBoundary,
  paragraphLength,
  paragraphLogicalText,
} from "./offsets";
import { findStoryParagraph } from "./paragraphLocator";
import { applyFormattingPatch } from "./patch";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  filterNewIds,
  leavesKeptByDirectDeletion,
  replacementDeletionSegments,
  selectedParagraphRuns,
  trimAppliedNewIds,
} from "./plan";
import {
  cutPieceHandoffs,
  planTrackedReplace,
  rangeEndAfterDeletion,
} from "./rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  isRemovedRevisionNode,
  paragraphPropertiesOf,
  reviewFieldsOf,
} from "./review";
import { groupedDeletionStamp, ownInsertionSlice } from "./revisionGrouping";
import { runsMergeable } from "./runMerge";
import { identitySlots } from "./slots";
import { tableEditParagraphDemand } from "./tableEdits";
import { tableGrid } from "./tableGrid";
import { locateTableRow, tableRowAnchor } from "./tableLocation";
import { rowsWithoutRow } from "./tableRowRemoval";
import { compileGenerateTOCIntent, type GenerateTOCIntent } from "./tocIntent";
import {
  type BlockInsertionPoint,
  type DeleteRangeOp,
  type DeleteRowOp,
  type DeleteTableOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  EMPTY_PROPERTY_SETS,
  type InlineSlice,
  type InsertContentOp,
  type InsertRowOp,
  type JoinInlineOp,
  type NewIds,
  type OpStory,
  type ParagraphPropsPatch,
  PROPERTY_REVIEW_POLICIES,
  type RevisionStamp,
  type RunPropsPatch,
  SECTION_BOUNDARY_POLICIES,
  SPLIT_HALVES,
  type SplitHalf,
  type TableEditOp,
  type TableIntentOperation,
  type TextPosition,
} from "./types";
import { captureDocumentOp } from "./wire";

type SplitParagraphIntent = {
  type: "splitParagraph";
  at: TextPosition;
  to?: TextPosition;
  /**
   * Which half takes the new identity: by default the first, as typing a
   * break moves the text before it to a paragraph of its own; the second
   * keeps the paragraph's identity on the text before the split.
   */
  newHalf?: SplitHalf;
};
export type { TableIntentOperation } from "./types";

/**
 * A row inserted into, or deleted from, the table holding `blockId`. An
 * inserted row's paragraphs take fresh identities from the compiler; the ids
 * it arrives with are not kept.
 */
export type TableRowIntentOperation =
  | Omit<InsertRowOp, "revision" | "newIds">
  | Omit<DeleteRowOp, "revision" | "newIds">;

/**
 * A table inserted after the paragraph `blockId`, with an empty paragraph
 * after it, as the editor inserts one, or between the blocks beside that
 * paragraph alone; or the table holding `blockId`,
 * removed, or in suggesting mode recorded as deleted. An inserted table's
 * paragraphs take fresh identities from the compiler.
 */
export type TableBlockIntentOperation =
  | Omit<DeleteTableOp, "revision" | "newIds" | "expected">
  | {
      type: typeof DOCUMENT_OP_TYPES.INSERT_TABLE;
      story: OpStory;
      blockId: string;
      table: Table;
      /**
       * Where the table goes against the anchor paragraph: after it (the
       * default, as from a caret), or before it, as from a selection of the
       * whole story, which the editor fills at the story's start.
       */
      side?: "after" | "before";
      /**
       * An empty paragraph ahead of the table as well, as the editor writes
       * a table it inserts after another, so the two never touch.
       */
      separated?: boolean;
      /**
       * The empty paragraph after the table: written, as the editor gives
       * the caret a place after a table it inserts (the default), or left
       * out, for a table that goes between blocks already standing on both
       * sides of it.
       */
      paragraphAfter?: "empty" | "none";
    };

/** Positions use canonical physical offsets, including retained deleted content. */
export type EditorIntent =
  | HyperlinkEditorIntent
  | GenerateTOCIntent
  | {
      type: "table";
      operation: TableIntentOperation;
      /**
       * What each cell an inserted column creates holds, in row order: the
       * compiler writes it into the new cells after the column, as text
       * typed there, so a suggestion records it as inserted. A cell past
       * the list, or given no paragraphs, stays empty. Only a column
       * insertion takes it.
       */
      cellContent?: readonly (readonly Paragraph[])[];
    }
  | { type: "tableRow"; operation: TableRowIntentOperation }
  | { type: "tableBlock"; operation: TableBlockIntentOperation }
  | {
      /**
       * Whole paragraphs beside a paragraph, which keeps its identity: the
       * compiler gives each a fresh id, and a suggestion records each as
       * inserted.
       */
      type: "insertParagraphs";
      story: OpStory;
      at: BlockInsertionPoint;
      paragraphs: readonly Paragraph[];
      /** A suggestion records the paragraphs as a move's destination. */
      move?: true;
      /** A list instance the paragraphs' numbering names, created first. */
      newList?: NewListInstance;
    }
  | {
      /** Whole adjacent paragraphs removed with their marks. */
      type: "deleteParagraphs";
      story: OpStory;
      blockIds: readonly string[];
      /** A suggestion records the paragraphs as a move's source. */
      move?: true;
    }
  | {
      type: "replaceFragment";
      from: TextPosition;
      to: TextPosition;
      paragraphs: readonly Paragraph[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
    }
  | {
      type: "moveFragment";
      from: TextPosition;
      to: TextPosition;
      target: TextPosition;
      paragraphs: readonly Paragraph[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
    }
  | {
      type: "replaceText";
      from: TextPosition;
      to: TextPosition;
      text: string;
      runProps?: TextFormatting;
      runPropsPatch?: RunPropsPatch;
      /**
       * Whether the caret at each end stood just inside the edge of an inline
       * container (a content control) that opens at `from` or closes at
       * `to`, where the source gap alone cannot tell inside from outside.
       */
      inside?: { start?: boolean; end?: boolean };
    }
  | {
      type: "formatRun";
      from: TextPosition;
      to: TextPosition;
      patch: RunPropsPatch;
    }
  | {
      type: "formatParagraph";
      at: TextPosition;
      patch: ParagraphPropsPatch;
      /** A list instance the patch's numbering names, created first. */
      newList?: NewListInstance;
    }
  | {
      type: "setList";
      items: readonly { at: TextPosition; ilvl: number }[];
      target:
        | { type: "existing"; numId: number }
        | {
            type: "new";
            num: NumberingInstance;
            abstractNum?: AbstractNumbering;
          };
    }
  | {
      type: "insertAtom";
      from: TextPosition;
      to: TextPosition;
      atom: TabContent | BreakContent;
      runProps?: TextFormatting;
      runPropsPatch?: RunPropsPatch;
    }
  | (SplitParagraphIntent & { newBlockId: string })
  | {
      type: "joinParagraphs";
      story: OpStory;
      blockId: string;
      nextBlockId: string;
      /**
       * Which paragraph's identity a direct join keeps, by default the
       * second's; a tracked join always keeps the second.
       */
      survivor?: SplitHalf;
    };

/** A numbering instance, with its abstract definition, an edit creates. */
export type NewListInstance = {
  num: NumberingInstance;
  abstractNum?: AbstractNumbering;
};

const newListOps = (newList: NewListInstance | undefined): DocumentOp[] =>
  newList === undefined
    ? []
    : [
        {
          type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
          num: newList.num,
          ...(newList.abstractNum === undefined
            ? {}
            : { abstractNum: newList.abstractNum }),
        },
      ];

/** Revision metadata and every fresh identity are supplied before compilation. */
export type EditorIntentMode = {
  /** Identities consumed during this editor lifetime, including undone insertions. */
  reservedBlockIds?: ReadonlySet<string>;
  /** The only ids the compiler may mint, when the author holds a lease. */
  idLease?: IdLease;
} & (
  | { type: "editing"; newIds?: NewIds }
  | { type: "suggesting"; revision: RevisionStamp; newIds: NewIds }
);

type CompileEditorIntentOptions = {
  intent: EditorIntent;
  mode: EditorIntentMode;
  /** Lowest candidate for compiler-owned pasted block IDs, including retired session IDs. */
  firstBlockId?: number;
};
type CompiledEditorIntent = {
  ops: DocumentOp[];
  selection: TextPosition;
  /**
   * What applying `ops` to the compiled document gives, when the compiler
   * already applied exactly them; a caller applies `ops` itself otherwise.
   */
  applied?: AppliedDocumentOp;
};

/** A split's paragraph identity is allocated with its other fresh identities. */
type EditorIntentAllocation = EditorIntent | SplitParagraphIntent;

const intentEndpoints = (intent: EditorIntentAllocation) => {
  switch (intent.type) {
    case "table":
    case "tableRow":
    case "tableBlock":
      return {
        story: intent.operation.story,
        fromId: intent.operation.blockId,
        toId: intent.operation.blockId,
        fromOffset: 0,
        toOffset: 0,
      };
    case "insertParagraphs":
      return {
        story: intent.story,
        fromId: intent.at.blockId,
        toId: intent.at.blockId,
        fromOffset: 0,
        toOffset: 0,
      };
    case "deleteParagraphs":
      return {
        story: intent.story,
        fromId: intent.blockIds[0] ?? "",
        toId: intent.blockIds.at(-1) ?? "",
      };
    case "setHyperlink":
    case "removeHyperlink":
    case "insertHyperlink":
    case "replaceFragment":
    case "moveFragment":
    case "replaceText":
    case "insertAtom":
    case "formatRun":
      return {
        story: intent.from.story,
        fromId: intent.from.blockId,
        toId: intent.to.blockId,
        fromOffset: intent.from.offset,
        toOffset: intent.to.offset,
      };
    case "formatParagraph":
      return {
        story: intent.at.story,
        fromId: intent.at.blockId,
        toId: intent.at.blockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    case "generateTOC":
      return {
        story: intent.at.story,
        fromId: intent.at.blockId,
        toId: intent.at.blockId,
        fromOffset: intent.at.offset,
        toOffset: intent.at.offset,
      };
    case "setList": {
      const first = intent.items.at(0);
      const last = intent.items.at(-1);
      if (first === undefined || last === undefined)
        panic("List allocation requires a paragraph.");
      return {
        story: first.at.story,
        fromId: first.at.blockId,
        toId: last.at.blockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    }
    case "splitParagraph":
      return {
        story: intent.at.story,
        fromId: intent.at.blockId,
        toId: (intent.to ?? intent.at).blockId,
        fromOffset: intent.at.offset,
        toOffset: (intent.to ?? intent.at).offset,
      };
    case "joinParagraphs":
      return {
        story: intent.story,
        fromId: intent.blockId,
        toId: intent.nextBlockId,
        fromOffset: undefined,
        toOffset: undefined,
      };
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
};

const PARAGRAPH_REVISION_DEMAND = 3;
const leafRevisionDemand = (ancestorCount: number) => 4 * (1 + ancestorCount);

/**
 * Cache only immutable document versions; never retain a retired version
 * strongly. A new paragraph identity avoids `reservedBlockIds`, the ones
 * consumed elsewhere, such as an undone insertion's.
 */
export const createEditorIntentIdAllocator = ({
  reservedBlockIds,
}: { reservedBlockIds?: ReadonlySet<string> } = {}) => {
  // A story's paragraphs, by the body object: a body an edit left is read once.
  const storyLocations = new WeakMap<
    ReturnType<typeof storyBody>,
    ReturnType<typeof storyParagraphs>
  >();
  return (
    document: Document,
    intent: EditorIntentAllocation,
    lease?: IdLease,
  ): {
    revisionId: number;
    newBlockId: string;
    newIds: {
      revision: number[];
      control: number[];
    };
  } => {
    const census = packageIdOccupancy(document.package);
    const { story, fromId, toId, fromOffset, toOffset } =
      intentEndpoints(intent);
    const body = storyBody(document, story);
    let locations = storyLocations.get(body);
    if (locations === undefined) {
      locations = storyParagraphs(body);
      storyLocations.set(body, locations);
    }
    const first = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(fromId),
    );
    const last = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(toId),
    );
    // Every selected leaf may start a deletion segment; ancestors may be cut
    // at both endpoints. Each paragraph adds join/mark/property stamps.
    let demand = 1;
    const selected =
      first < 0 || last < first ? [] : locations.slice(first, last + 1);
    for (const { paragraph } of intent.type === "table" ? [] : selected) {
      demand += PARAGRAPH_REVISION_DEMAND;
      const blockId = idKey(paragraph.paraId ?? "");
      for (const span of leafSpans(paragraph.content)) {
        if (
          blockId === idKey(fromId) &&
          fromOffset !== undefined &&
          span.after.offset < fromOffset
        )
          continue;
        if (
          blockId === idKey(toId) &&
          toOffset !== undefined &&
          span.before.offset > toOffset
        )
          continue;
        demand += leafRevisionDemand(span.ancestors.length);
      }
    }
    const incoming =
      intent.type === "replaceFragment" || intent.type === "moveFragment"
        ? intent.paragraphs
        : [];
    for (const paragraph of incoming) {
      demand += 4;
      for (const span of leafSpans(paragraph.content))
        demand += 4 * (1 + span.ancestors.length);
    }
    if (incoming.length > 0) demand += 8;
    // A suggested link change strikes the text it changes and inserts the
    // changed copy after each stretch of it; a suggested contents list
    // records every paragraph it inserts and the split it may make.
    if (
      intent.type === "setHyperlink" ||
      intent.type === "removeHyperlink" ||
      intent.type === "insertHyperlink"
    )
      demand = 2 * demand + 8;
    if (intent.type === "generateTOC")
      demand +=
        8 +
        (intent.headings.length + 1) *
          (PARAGRAPH_REVISION_DEMAND + 4 * leafRevisionDemand(3));
    if (intent.type === "moveFragment") {
      const target = storyParagraphs(
        storyBody(document, intent.target.story),
      ).find(
        ({ paragraph }) =>
          idKey(paragraph.paraId ?? "") === idKey(intent.target.blockId),
      );
      if (target !== undefined) {
        demand += 3;
        for (const span of leafSpans(target.paragraph.content))
          demand += 4 * (1 + span.ancestors.length);
      }
    }
    const fresh = (
      occupied: { has: (id: number) => boolean },
      count: number,
      kind: IdLeaseKind,
    ) =>
      freshLeasedIds({
        lease,
        kind,
        count,
        taken: (id) => occupied.has(id),
      });
    if (intent.type === "insertParagraphs") {
      // Each paragraph's mark and content wrappers, the anchor's mark, and
      // the property change the paragraph that takes over that mark records.
      demand = 3;
      for (const paragraph of intent.paragraphs) {
        demand += PARAGRAPH_REVISION_DEMAND;
        for (const span of leafSpans(paragraph.content))
          demand += leafRevisionDemand(span.ancestors.length);
      }
    }
    if (intent.type === "tableBlock") {
      const op = intent.operation;
      // An inserted table records each row and each of its paragraphs, and
      // the empty paragraph after it; a deleted one, each row and paragraph.
      const table =
        op.type === DOCUMENT_OP_TYPES.INSERT_TABLE
          ? op.table
          : locateTableRow(document, op)
              .map(({ table: found }) => found)
              .unwrapOr(undefined);
      if (table !== undefined) {
        // Each leaf may start its own deletion record where pending
        // revisions already cut the text, as a range deletion counts it.
        demand = 2 + table.rows.length + PARAGRAPH_REVISION_DEMAND;
        for (const { paragraph } of storyParagraphs({
          content: [{ type: "table", rows: table.rows }],
        })) {
          demand += PARAGRAPH_REVISION_DEMAND;
          for (const span of leafSpans(paragraph.content))
            demand += leafRevisionDemand(span.ancestors.length);
        }
      }
    }
    if (intent.type === "table") {
      const op = intent.operation;
      const located = locateTableRow(document, op);
      demand = 1;
      if (located.isOk()) {
        switch (op.type) {
          case DOCUMENT_OP_TYPES.INSERT_COLUMN:
            demand = located.value.table.rows.length + 2;
            // Each new cell's content, as a fragment written into it.
            for (const paragraphs of intent.cellContent ?? []) {
              demand += 8;
              for (const paragraph of paragraphs) {
                demand += 4;
                for (const span of leafSpans(paragraph.content))
                  demand += 4 * (1 + span.ancestors.length);
              }
            }
            break;
          case DOCUMENT_OP_TYPES.DELETE_COLUMN: {
            const table = located.value.table;
            demand = table.rows.length + 2;
            const grid = tableGrid(table, op.type);
            if (op.column !== 0 || grid.isErr() || grid.value.width !== 1)
              break;
            // Removing the final column tracks the whole table, including paragraph
            // marks and every inline wrapper identity, rather than just its rows.
            for (const { paragraph } of storyParagraphs({ content: [table] })) {
              demand += PARAGRAPH_REVISION_DEMAND;
              for (const span of leafSpans(paragraph.content))
                demand += leafRevisionDemand(span.ancestors.length);
            }
            break;
          }
          case DOCUMENT_OP_TYPES.MERGE_CELLS:
            demand = (op.bottom - op.top + 1) * 2;
            break;
          case DOCUMENT_OP_TYPES.SPLIT_CELL:
            // A cell for each column of the span in each row a vertical
            // merge reaches, which is at most every row.
            demand =
              (located.value.table.rows
                .at(located.value.rowIndex)
                ?.cells.at(located.value.cellIndex)?.formatting?.gridSpan ??
                1) * located.value.table.rows.length;
            break;
          case DOCUMENT_OP_TYPES.SET_TABLE_GRID:
            demand = 2;
            break;
          case DOCUMENT_OP_TYPES.SET_CELL_PROPS:
          case DOCUMENT_OP_TYPES.SET_ROW_PROPS:
          case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
            break;
          default: {
            const unreachable: never = op;
            return unreachable;
          }
        }
      }
    }
    const revision = fresh(
      census.revisions,
      demand + 1,
      ID_LEASE_KINDS.REVISION,
    );
    const taken = (id: string) =>
      census.paragraphs.has(id) || (reservedBlockIds?.has(id) ?? false);
    // An exhausted space or lease gives ids no operation accepts, so the edit
    // is refused rather than handed an id that wraps or collides.
    const paragraph = freshLeasedIds({
      lease,
      kind: ID_LEASE_KINDS.PARAGRAPH,
      count: 1,
      taken: (id) => taken(formatParaId(id)),
    }).at(0);
    const newBlockId = formatParaId(paragraph ?? ID_KIND_LIMITS.paragraph);
    return {
      revisionId: revision.at(0) ?? MAX_REVISION_ID + 1,
      newBlockId,
      newIds: {
        revision: revision.slice(1),
        control: fresh(
          census.controls,
          intent.type === "table" ? 0 : demand,
          ID_LEASE_KINDS.CONTROL,
        ),
      },
    };
  };
};

/** Fresh identities bounded by the edit, with one package census per document version. */
export const allocateEditorIntentIds: (
  document: Document,
  intent: EditorIntentAllocation,
  lease?: IdLease,
) => {
  revisionId: number;
  newBlockId: string;
  newIds: {
    revision: number[];
    control: number[];
  };
} = createEditorIntentIdAllocator();

/** Accepted-view text without changing the canonical paragraph or its offset space. */
export const paragraphVisibleText = (paragraph: Paragraph): string => {
  const text = paragraphLogicalText(paragraph);
  return leafSpans(paragraph.content)
    .filter(({ ancestors }) => !ancestors.some(isRemovedRevisionNode))
    .map(({ before, after }) => text.slice(before.offset, after.offset))
    .join("");
};

/**
 * Right-affine translation: a visible gap follows retained deleted content at
 * that gap, so a new insertion never lands inside a deletion wrapper.
 * Invalid offsets remain invalid for the operation boundary to diagnose.
 */
export const physicalOffsetAtVisibleOffset = (
  paragraph: Paragraph,
  offset: number,
): number => {
  if (!Number.isInteger(offset) || offset < 0) return -1;
  let visible = 0;
  for (const span of leafSpans(paragraph.content)) {
    if (span.ancestors.some(isRemovedRevisionNode)) continue;
    const width = span.after.offset - span.before.offset;
    if (offset < visible + width) return span.before.offset + offset - visible;
    visible += width;
  }
  return offset === visible
    ? paragraphLength(paragraph)
    : paragraphLength(paragraph) + 1;
};

/** The visible offset corresponding to a physical source position. */
const visibleOffsetAtPhysicalOffset = (
  paragraph: Paragraph,
  offset: number,
): number => {
  if (
    !Number.isInteger(offset) ||
    offset < 0 ||
    offset > paragraphLength(paragraph)
  )
    return -1;
  let visible = 0;
  for (const span of leafSpans(paragraph.content)) {
    if (span.before.offset >= offset) break;
    if (span.ancestors.some(isRemovedRevisionNode)) continue;
    visible += Math.min(offset, span.after.offset) - span.before.offset;
  }
  return visible;
};

/** Deleted paragraph marks concatenate physical paragraphs in the editing view. */
export const editorParagraphGroups = (
  document: Document,
  story: OpStory,
): {
  blockId: string;
  paragraphs: Paragraph[];
  text: string;
}[] => {
  const groups: { blockId: string; paragraphs: Paragraph[]; text: string }[] =
    [];
  let pending: Paragraph[] = [];
  let previous: ReturnType<typeof storyParagraphs>[number] | undefined;
  for (const location of storyParagraphs(storyBody(document, story))) {
    const { paragraph } = location;
    if (
      pending.length > 0 &&
      previous &&
      (!sameBlockList(previous.list, location.list) ||
        location.index !== previous.index + 1)
    ) {
      groups.push({
        blockId: previous.paragraph.paraId ?? "",
        paragraphs: pending,
        text: pending.map(paragraphVisibleText).join(""),
      });
      pending = [];
    }
    previous = location;
    pending.push(paragraph);
    if (
      paragraph.pPrMark?.kind === "del" ||
      paragraph.pPrMark?.kind === "moveFrom"
    )
      continue;
    groups.push({
      blockId: paragraph.paraId ?? "",
      paragraphs: pending,
      text: pending.map(paragraphVisibleText).join(""),
    });
    pending = [];
  }
  if (pending.length > 0 && previous)
    groups.push({
      blockId: previous.paragraph.paraId ?? "",
      paragraphs: pending,
      text: pending.map(paragraphVisibleText).join(""),
    });
  return groups;
};

/** One logical paragraph and visible offset for a physical canonical address. */
export const editorPositionAtPhysicalPosition = (
  document: Document,
  at: TextPosition,
): TextPosition => {
  const group = editorParagraphGroups(document, at.story).find(
    ({ paragraphs }) =>
      paragraphs.some(
        ({ paraId }) => idKey(paraId ?? "") === idKey(at.blockId),
      ),
  );
  if (group === undefined) return at;
  let offset = 0;
  for (const paragraph of group.paragraphs) {
    if (idKey(paragraph.paraId ?? "") === idKey(at.blockId)) {
      const local = visibleOffsetAtPhysicalOffset(paragraph, at.offset);
      return {
        story: at.story,
        blockId: group.blockId,
        offset: local < 0 ? -1 : offset + local,
      };
    }
    offset += paragraphVisibleText(paragraph).length;
  }
  return panic("A logical paragraph lost its source position");
};

/** Translate an editing-view paragraph gap to its physical canonical address. */
export const physicalPositionAtEditorOffset = (
  document: Document,
  at: TextPosition,
): TextPosition => {
  const group = editorParagraphGroups(document, at.story).find(
    ({ blockId }) => idKey(blockId) === idKey(at.blockId),
  );
  if (group === undefined) return at;
  let remaining = at.offset;
  for (const [index, paragraph] of group.paragraphs.entries()) {
    const width = paragraphVisibleText(paragraph).length;
    if (remaining < width || index === group.paragraphs.length - 1) {
      return {
        story: at.story,
        blockId: paragraph.paraId ?? "",
        offset: physicalOffsetAtVisibleOffset(paragraph, remaining),
      };
    }
    remaining -= width;
  }
  return at;
};

// A zero-width ordinal belongs to one offset. At a new text offset, the
// projection resolves the available gap instead of reusing the insertion gap.
/**
 * The caret after typing `text` at `at`: directly after the typed text, before
 * any zero-width marker that follows it, so the next keystroke continues the
 * same run instead of landing on the far side of a comment, bookmark or
 * permission boundary.
 */
const replacementCaret = (at: TextPosition, text: string): TextPosition =>
  text.length > 0
    ? {
        story: at.story,
        blockId: at.blockId,
        offset: at.offset + text.length,
        zeroWidthBefore: 0,
      }
    : {
        story: at.story,
        blockId: at.blockId,
        offset: at.offset,
        ...(at.zeroWidthBefore === undefined
          ? {}
          : { zeroWidthBefore: at.zeroWidthBefore }),
      };

const paragraphAt = (document: Document, at: TextPosition) =>
  findStoryParagraph(storyBody(document, at.story), at.blockId)?.paragraph;

/** Zero-width markers and empty leaves never supply the formatting of typed content. */
const visibleFormattingSpans = (paragraph: Paragraph) =>
  leafSpans(paragraph.content).filter(
    ({ ancestors, before, after }) =>
      after.offset > before.offset && !ancestors.some(isRemovedRevisionNode),
  );

/** Capture authored formatting before deletion, using only visible runs. */
const formattingAt = (
  paragraph: Paragraph,
  offset: number,
): TextFormatting | undefined => {
  const spans = visibleFormattingSpans(paragraph);
  const span =
    spans.find(
      ({ before, after }) => before.offset <= offset && offset < after.offset,
    ) ??
    spans.findLast(({ after }) => after.offset <= offset) ??
    spans.find(({ before }) => before.offset > offset);
  return span?.ancestors.findLast((ancestor) => ancestor.type === "run")
    ?.formatting;
};

const authoredFormatting = (
  document: Document,
  from: TextPosition,
  to: TextPosition,
): TextFormatting | undefined => {
  const paragraph = paragraphAt(document, from);
  if (paragraph === undefined)
    panic(
      "A validated input paragraph must exist when reading authored formatting.",
    );
  const collapsed = from.blockId === to.blockId && from.offset === to.offset;
  if (!collapsed) return formattingAt(paragraph, from.offset);
  const left = visibleFormattingSpans(paragraph).some(
    ({ before }) => before.offset < from.offset,
  );
  if (left) return formattingAt(paragraph, from.offset - 1);
  const group = editorParagraphGroups(document, from.story).find(
    ({ paragraphs }) =>
      paragraphs.some(
        ({ paraId }) => idKey(paraId ?? "") === idKey(from.blockId),
      ),
  );
  const index =
    group?.paragraphs.findIndex(
      ({ paraId }) => idKey(paraId ?? "") === idKey(from.blockId),
    ) ?? 0;
  const previous = group?.paragraphs
    .slice(0, index)
    .findLast((item) => paragraphVisibleText(item) !== "");
  return previous === undefined
    ? formattingAt(paragraph, from.offset)
    : formattingAt(previous, paragraphLength(previous));
};

type IntentRunFormattingOptions = {
  from: TextPosition;
  to: TextPosition;
  runProps?: TextFormatting;
  runPropsPatch?: RunPropsPatch;
};

const intentRunFormatting = (
  document: Document,
  { from, to, runProps, runPropsPatch }: IntentRunFormattingOptions,
) => {
  const authored = runProps ?? authoredFormatting(document, from, to);
  return runPropsPatch === undefined
    ? authored
    : (applyFormattingPatch(authored, runPropsPatch) ?? {});
};

/** The ending physical mark owns one logical paragraph's properties. */
const createParagraphMarkOwnerLookup = (document: Document) => {
  const stories = new Map<string, Map<string, Paragraph>>();
  return (at: TextPosition) => {
    const key = JSON.stringify(at.story);
    let owners = stories.get(key);
    if (owners === undefined) {
      owners = new Map();
      for (const { paragraphs } of editorParagraphGroups(document, at.story)) {
        const owner = paragraphs.at(-1);
        if (owner === undefined)
          panic("A logical paragraph has no property owner");
        for (const paragraph of paragraphs)
          owners.set(idKey(paragraph.paraId ?? ""), owner);
      }
      stories.set(key, owners);
    }
    return owners.get(idKey(at.blockId));
  };
};

/** A new logical paragraph inherits the surviving paragraph mark's formatting. */
const splitParagraphFields = (document: Document, at: TextPosition) => {
  const survivor = createParagraphMarkOwnerLookup(document)(at);
  return survivor?.formatting === undefined
    ? {}
    : { formatting: survivor.formatting };
};

/** A compound intent consumes one identity pool across its sequential plans. */
const modeAfterOps = (
  document: Document,
  mode: EditorIntentMode,
): EditorIntentMode => {
  if (mode.newIds === undefined) return mode;
  const occupied = new Set(
    packageIdentityKeys(document.package).concat(
      packageReservedIdentityKeys(document.package),
    ),
  );
  const revisionIds = mode.newIds.revision?.filter(
    (id) => !occupied.has(`${IDENTITY_SPACES.REVISION}:${id}`),
  );
  const controlIds = mode.newIds.control?.filter(
    (id) => !occupied.has(`${IDENTITY_SPACES.CONTROL}:${id}`),
  );
  const newIds = {
    ...(revisionIds === undefined ? {} : { revision: revisionIds }),
    ...(controlIds === undefined ? {} : { control: controlIds }),
  };
  const leased = mode.idLease === undefined ? {} : { idLease: mode.idLease };
  if (mode.type === "editing") return { type: "editing", newIds, ...leased };
  if (!occupied.has(`${IDENTITY_SPACES.REVISION}:${mode.revision.id}`))
    return { type: "suggesting", revision: mode.revision, newIds, ...leased };
  const id = newIds.revision?.at(0) ?? MAX_REVISION_ID + 1;
  return {
    ...leased,
    type: "suggesting",
    revision: { ...mode.revision, id },
    newIds: {
      ...newIds,
      ...(revisionIds === undefined ? {} : { revision: revisionIds.slice(1) }),
    },
  };
};

/** Imported identities are package-local facts, never destination identities. */
const isCopiedRangeStart = (value: object): boolean =>
  "type" in value &&
  (value.type === "bookmarkStart" ||
    value.type === "moveFromRangeStart" ||
    value.type === "moveToRangeStart");

const isCopiedRangeMarker = (value: object): boolean =>
  isCopiedRangeStart(value) ||
  ("type" in value &&
    (value.type === "bookmarkEnd" ||
      value.type === "moveFromRangeEnd" ||
      value.type === "moveToRangeEnd"));

const isMoveAnchor = (node: InlineNode): boolean =>
  isCommentAnchor(node) || (isParagraphContent(node) && isRangeBoundary(node));

/** Suggested moves retain source anchors; their inserted copy carries only tracked content. */
const withoutMoveAnchors = (nodes: readonly InlineNode[]): InlineNode[] => {
  const out: InlineNode[] = [];
  for (const node of nodes) {
    if (isMoveAnchor(node)) continue;
    const children = childNodes(node);
    const retained =
      children === undefined
        ? node
        : rebuildNode(node, withoutMoveAnchors(children));
    if (!isEmptyRecord(retained)) out.push(retained);
  }
  return out;
};

type IdentifyClipboardParagraphsOptions = {
  document: Document;
  paragraphs: readonly Paragraph[];
  mode: EditorIntentMode;
  firstBlockId: number | undefined;
};

const identifyClipboardParagraphs = ({
  document,
  paragraphs,
  mode,
  firstBlockId,
}: IdentifyClipboardParagraphsOptions): Result<
  Paragraph[],
  DocumentOpRefusal
> => {
  const copy = cloneModel([...paragraphs]);
  const occupied = new Set(
    packageIdentityKeys(document.package).concat(
      packageReservedIdentityKeys(document.package),
    ),
  );
  for (const id of mode.newIds?.revision ?? [])
    occupied.add(`${IDENTITY_SPACES.REVISION}:${id}`);
  for (const id of mode.newIds?.control ?? [])
    occupied.add(`${IDENTITY_SPACES.CONTROL}:${id}`);
  if (mode.type === "suggesting")
    occupied.add(`${IDENTITY_SPACES.REVISION}:${mode.revision.id}`);
  const blockIds = new Set(packageParagraphIds(document.package).map(idKey));
  const markers = new Set<number>();
  const names = new Set<string>();
  const renamedBookmarks = new Map<string, string>();
  const visit = (value: unknown, callback: (value: object) => void): void => {
    if (typeof value !== "object" || value === null) return;
    if (Array.isArray(value)) {
      for (const item of value) visit(item, callback);
      return;
    }
    callback(value);
    for (const field of Object.values(value)) visit(field, callback);
  };
  visit(document.package, (value) => {
    if (
      "type" in value &&
      "id" in value &&
      typeof value.id === "number" &&
      isCopiedRangeMarker(value)
    )
      markers.add(value.id);
    if (
      isCopiedRangeStart(value) &&
      "name" in value &&
      typeof value.name === "string"
    )
      names.add(value.name);
  });
  visit(copy, (value) => {
    if (
      !isCopiedRangeStart(value) ||
      !("name" in value) ||
      typeof value.name !== "string"
    )
      return;
    const original = value.name;
    const renamed = renamedBookmarks.get(original);
    if (renamed !== undefined) {
      value.name = renamed;
      return;
    }
    let name = original;
    let suffix = 1;
    while (names.has(name)) {
      name = `${original}_paste${suffix}`;
      suffix += 1;
    }
    names.add(name);
    renamedBookmarks.set(original, name);
    value.name = name;
  });
  const remapped = new Map<string, number>();
  const nextInSpace = new Map<string, number>();
  let nextBlock = firstBlockId ?? 1;
  const freshBlock = () => {
    const fresh = freshLeasedIds({
      lease: mode.idLease,
      kind: ID_LEASE_KINDS.PARAGRAPH,
      count: 1,
      from: nextBlock,
      taken: (candidate) => blockIds.has(formatParaId(candidate)),
    }).at(0);
    if (fresh === undefined) return undefined;
    const id = formatParaId(fresh);
    blockIds.add(id);
    nextBlock = fresh + 1;
    return id;
  };
  // A leased author mints revision and content-control ids from its lease.
  let leaseExhausted = false;
  const freshId = (space: string, oldId: number) => {
    const key = `${space}:${oldId}`;
    const existing = remapped.get(key);
    if (existing !== undefined) return existing;
    let id = nextInSpace.get(space) ?? 1;
    if (
      mode.idLease !== undefined &&
      (space === IDENTITY_SPACES.REVISION || space === IDENTITY_SPACES.CONTROL)
    ) {
      const leased = freshLeasedIds({
        lease: mode.idLease,
        kind:
          space === IDENTITY_SPACES.REVISION
            ? ID_LEASE_KINDS.REVISION
            : ID_LEASE_KINDS.CONTROL,
        count: 1,
        from: id,
        taken: (candidate) => occupied.has(`${space}:${candidate}`),
      }).at(0);
      if (leased === undefined) leaseExhausted = true;
      id = leased ?? MAX_REVISION_ID + 1;
    }
    while (
      space === "bookmark" ? markers.has(id) : occupied.has(`${space}:${id}`)
    )
      id += 1;
    if (space === "bookmark") markers.add(id);
    else occupied.add(`${space}:${id}`);
    remapped.set(key, id);
    nextInSpace.set(space, id + 1);
    return id;
  };
  const remintedControls = new WeakSet<object>();
  visit(copy, (value) => {
    if (
      "type" in value &&
      value.type === "inlineSdt" &&
      "properties" in value &&
      typeof value.properties === "object" &&
      value.properties !== null &&
      "id" in value.properties &&
      typeof value.properties.id === "number"
    ) {
      value.properties.id = freshId(
        IDENTITY_SPACES.CONTROL,
        value.properties.id,
      );
      remintedControls.add(value.properties);
    }
    if (
      "type" in value &&
      value.type === "hyperlink" &&
      "anchor" in value &&
      typeof value.anchor === "string"
    )
      value.anchor = renamedBookmarks.get(value.anchor) ?? value.anchor;
    if (
      "id" in value &&
      typeof value.id === "number" &&
      !remintedControls.has(value)
    ) {
      if (
        "space" in value &&
        (value.space === IDENTITY_SPACES.REVISION ||
          value.space === IDENTITY_SPACES.CONTROL)
      ) {
        // Cut provenance names the same copied records as info/properties IDs.
        value.id = freshId(value.space, value.id);
      } else if ("sdtType" in value)
        value.id = freshId(IDENTITY_SPACES.CONTROL, value.id);
      else if ("author" in value && !("type" in value))
        value.id = freshId(IDENTITY_SPACES.REVISION, value.id);
      else if (isCopiedRangeMarker(value))
        value.id = freshId("bookmark", value.id);
    }
  });
  if (leaseExhausted)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message: "The pasted content needs more ids than the lease holds.",
        opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      }),
    );
  // Only split-off paragraphs need new IDs; the inline tail keeps its
  // destination identity, including when the session has exhausted block IDs.
  for (const [index, paragraph] of copy.entries()) {
    delete paragraph.paraId;
    if (index === copy.length - 1) continue;
    const id = freshBlock();
    if (id === undefined) {
      return Result.err(
        new DocumentOpRefusal({
          reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
          message: "The pasted block identity allocator is exhausted.",
          opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        }),
      );
    }
    paragraph.paraId = id;
  }
  return Result.ok(copy);
};

type JoinParagraphPropertyAlignmentOptions = {
  document: Document;
  current: Document;
  intent: Extract<EditorIntent, { type: "joinParagraphs" }>;
  revision?: RevisionStamp;
};

/**
 * Direct and tracked joins choose paragraph properties from the same
 * editing-view group: the joined group shows the first group's properties,
 * or the following group's when the first shows no text. A paragraph whose
 * mark is already pending deletion is one group with the paragraph after it,
 * so joining the two physically leaves that group's properties as they were.
 * The properties go on the paragraph that ends the joined group once the join
 * is applied, whichever paragraph the join keeps, since that paragraph's mark
 * gives the group its properties.
 */
const joinParagraphPropertyAlignment = ({
  document,
  current,
  intent,
  revision,
}: JoinParagraphPropertyAlignmentOptions) => {
  const holding = (
    groups: ReturnType<typeof editorParagraphGroups>,
    blockId: string,
  ) =>
    groups.find(({ paragraphs }) =>
      paragraphs.some(({ paraId }) => idKey(paraId ?? "") === idKey(blockId)),
    );
  const groups = editorParagraphGroups(document, intent.story);
  const firstGroup = holding(groups, intent.blockId);
  const followingGroup = holding(groups, intent.nextBlockId);
  if (!firstGroup || !followingGroup)
    panic("A valid canonical join must belong to paragraph groups.");
  const source =
    firstGroup.text.length === 0
      ? followingGroup.paragraphs.at(-1)
      : firstGroup.paragraphs.at(-1);
  // A direct join may keep the first paragraph; a tracked one keeps the second.
  const keptId =
    revision === undefined && intent.survivor === "first"
      ? intent.blockId
      : intent.nextBlockId;
  const joinedGroup = holding(
    editorParagraphGroups(current, intent.story),
    keptId,
  );
  const survivor = joinedGroup?.paragraphs.at(-1);
  const survivorId = survivor?.paraId;
  if (!survivor || survivorId === undefined)
    panic("An applied join must keep the paragraph it names as kept.");
  const desired = paragraphPropertiesOf(source?.formatting);
  const existing = paragraphPropertiesOf(survivor.formatting);
  if (structurallyEqual(existing, desired)) return undefined;
  return {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: intent.story,
    blockId: survivorId,
    patch: Object.fromEntries([
      ...Object.keys(existing ?? {}).map((key) => [key, null]),
      ...Object.entries(desired ?? {}),
    ]),
    whenEmpty:
      source?.formatting === undefined
        ? EMPTY_PROPERTY_SETS.OMIT
        : EMPTY_PROPERTY_SETS.KEEP,
    ...(revision === undefined ? {} : { revision }),
  } as const;
};

const RUN_SEAM_POLICIES = {
  plain: "plain",
  sourceCut: "sourceCut",
} as const;

type RunSeamDepthOptions = {
  paragraph: Paragraph;
  at: Gap;
  policy: (typeof RUN_SEAM_POLICIES)[keyof typeof RUN_SEAM_POLICIES];
};

/** Join runs meeting at an edited gap under its consolidation policy. */
const runSeamDepth = ({
  paragraph,
  at,
  policy,
}: RunSeamDepthOptions): number => {
  const spans = leafSpans(paragraph.content);
  for (const [index, right] of spans.entries()) {
    const left = spans.at(index - 1);
    if (
      index === 0 ||
      left === undefined ||
      compareGaps(left.after, at) !== 0 ||
      compareGaps(right.before, at) !== 0
    )
      continue;
    const leftRun = left.ancestors.at(-1);
    const rightRun = right.ancestors.at(-1);
    if (
      leftRun?.type !== "run" ||
      rightRun?.type !== "run" ||
      leftRun === rightRun ||
      leftRun.content.at(-1) !== left.node ||
      rightRun.content.at(0) !== right.node ||
      left.ancestors.length !== right.ancestors.length ||
      !left.ancestors
        .slice(0, -1)
        .every(
          (parent, parentIndex) => parent === right.ancestors[parentIndex],
        ) ||
      (policy === RUN_SEAM_POLICIES.plain && !runsMergeable(leftRun, rightRun))
    )
      continue;
    return alikeDepth(leftRun, rightRun);
  }
  return 0;
};

/**
 * Where a direct replacement inside a source run rejoins the run it cut: at
 * the replacement's start, where accepting the tracked replacement (struck
 * content first, then the new) removes the deletion and joins its sides.
 */
const replacementSourceRunStart = (
  document: Document,
  op: InsertContentOp,
): TextPosition | undefined => {
  const paragraph = paragraphAt(document, op.at);
  if (paragraph === undefined)
    panic("An accepted replacement insertion has its source paragraph.");
  const at = {
    offset: op.at.offset,
    zeroWidthBefore:
      op.at.zeroWidthBefore ??
      defaultInsertionGap(paragraph.content, op.at.offset).zeroWidthBefore,
  };
  if (
    !spanningRecords(paragraph.content, [at], 0, 1).some(
      (node) => node.type === "run",
    )
  )
    return undefined;
  return { ...op.at, ...at };
};

type ReplacementPreservesOutsideContentOptions = {
  paragraph: Paragraph;
  from: TextPosition;
  to: TextPosition;
};

const replacementPreservesOutsideContent = ({
  paragraph,
  from,
  to,
}: ReplacementPreservesOutsideContentOptions): boolean => {
  const start = resolveGap({ paragraph, position: from, fallback: "afterAll" });
  const end = resolveGap({ paragraph, position: to, fallback: "beforeAll" });
  if (typeof start === "string" || typeof end === "string") return false;
  const [before = [], , after = []] = partitionContent(paragraph.content, [
    start,
    end,
  ]);
  const preserved = (node: InlineNode) =>
    isParagraphContent(node) &&
    (isRangeBoundary(node) || node.type === "commentReference");
  return before.every(preserved) && after.every(preserved);
};

/**
 * Where the caret lands once whole paragraphs go directly: at the start of the
 * paragraph that follows them, or, when they end their container, in the
 * final paragraph that stays, where its predecessor's content ends.
 */
const directDeletionCaret = (
  document: Document,
  { story, blockIds }: { story: OpStory; blockIds: readonly string[] },
): TextPosition => {
  const located = storyParagraphs(storyBody(document, story));
  const indexOf = (id: string | undefined) =>
    located.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(id ?? ""),
    );
  const first = located[indexOf(blockIds[0])];
  const last = located[indexOf(blockIds.at(-1))];
  // A deletion naming no paragraph is refused when it applies.
  if (first === undefined || last === undefined)
    return { story, blockId: blockIds[0] ?? "", offset: 0 };
  const list = blockListAt(storyBody(document, story).content, last.list);
  const following = list[last.index + 1];
  if (following?.type === "paragraph" && following.paraId !== undefined)
    return { story, blockId: following.paraId, offset: 0 };
  const preceding = list[first.index - 1];
  return {
    story,
    blockId: last.paragraph.paraId ?? "",
    offset:
      last.index === list.length - 1 && preceding?.type === "paragraph"
        ? paragraphLength(preceding)
        : 0,
  };
};

/** Compile one batch; the caller applies it atomically and journals its exact inverse. */
export const compileEditorIntent = (
  document: Document,
  { intent, mode, firstBlockId }: CompileEditorIntentOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  if (
    firstBlockId !== undefined &&
    (!Number.isInteger(firstBlockId) ||
      firstBlockId < 1 ||
      firstBlockId > 0x80000000)
  ) {
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
        message: "The pasted block identity allocator is exhausted or invalid.",
        opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      }),
    );
  }
  const allocationFields =
    mode.newIds === undefined ? {} : { newIds: mode.newIds };
  const tracked =
    mode.type === "suggesting"
      ? { revision: mode.revision, newIds: mode.newIds }
      : allocationFields;
  let ops: DocumentOp[];
  let selection: TextPosition;
  // The application of the ops a branch compiled, kept only while `ops`
  // still holds exactly the ops it applied.
  let handedBack:
    { ops: readonly DocumentOp[]; result: AppliedDocumentOp } | undefined;
  const editedSeams: TextPosition[] = [];
  switch (intent.type) {
    case "generateTOC": {
      const compiled = compileGenerateTOCIntent(
        document,
        intent,
        mode,
        (source, part, partMode) =>
          compileEditorIntent(source, {
            intent: part,
            mode: partMode,
            ...(firstBlockId === undefined ? {} : { firstBlockId }),
          }),
      );
      if (compiled.isErr()) return compiled;
      ops = compiled.value.ops;
      selection = compiled.value.selection;
      break;
    }
    case "setHyperlink":
    case "removeHyperlink":
    case "insertHyperlink":
      return compileHyperlinkIntent(document, {
        intent,
        mode,
        compileIntent: (source, part, partMode) =>
          compileEditorIntent(source, {
            intent: part,
            mode: partMode,
            ...(firstBlockId === undefined ? {} : { firstBlockId }),
          }),
      });
    case "replaceFragment": {
      if (intent.paragraphs.length === 0)
        return compileEditorIntent(document, {
          intent: {
            type: "replaceText",
            from: intent.from,
            to: intent.to,
            text: "",
          },
          mode,
        });
      const destination = paragraphAt(document, intent.from);
      if (destination === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The clipboard destination paragraph does not exist.",
            opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          }),
        );
      const endpoint = paragraphAt(document, intent.to);
      const destinationFormatting = splitParagraphFields(
        document,
        intent.from,
      ).formatting;
      const endpointFormatting = splitParagraphFields(
        document,
        intent.to,
      ).formatting;
      const prefixBoundary = intent.openStart === 0 && intent.from.offset > 0;
      const suffixBoundary =
        intent.openEnd === 0 &&
        endpoint !== undefined &&
        intent.to.offset < paragraphLength(endpoint);
      const normalized = [
        ...(prefixBoundary
          ? [
              {
                type: "paragraph",
                content: [],
                ...(destinationFormatting === undefined
                  ? {}
                  : { formatting: destinationFormatting }),
              } satisfies Paragraph,
            ]
          : []),
        ...intent.paragraphs,
        ...(suffixBoundary
          ? [
              {
                type: "paragraph",
                content: [],
                ...(endpointFormatting === undefined
                  ? {}
                  : { formatting: endpointFormatting }),
              } satisfies Paragraph,
            ]
          : []),
      ];
      const identified = identifyClipboardParagraphs({
        document,
        paragraphs: normalized,
        mode,
        firstBlockId,
      });
      if (identified.isErr()) return identified;
      const paragraphs = identified.value;
      for (const [index, paragraph] of paragraphs.entries()) {
        const openStart = index === 0 && intent.openStart === 1;
        const openEnd = index === normalized.length - 1 && intent.openEnd === 1;
        if (!openStart && !openEnd) continue;
        // The compiler owns these cloned paragraphs. Open edges join the
        // destination mark, keeping their authored inline content.
        delete paragraph.propertyChanges;
        delete paragraph.pPrMark;
        delete paragraph.reviewCarrier;
        delete paragraph.sectionProperties;
        delete paragraph.formatting;
        const inherited = openStart
          ? destinationFormatting
          : endpointFormatting;
        if (inherited !== undefined) paragraph.formatting = inherited;
      }
      if (
        paragraphs.some((paragraph) => paragraph.reviewCarrier !== undefined)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A clipboard paragraph cannot import a private review-resolution carrier.",
            opType: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          }),
        );
      }
      if (
        paragraphs.some(
          (paragraph) =>
            (paragraph.sectionProperties?.headerReferences?.length ?? 0) > 0 ||
            (paragraph.sectionProperties?.footerReferences?.length ?? 0) > 0,
        )
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY,
            message:
              "Clipboard section headers and footers require importing their source package parts.",
            opType: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
          }),
        );
      }
      if (
        mode.type === "suggesting" &&
        paragraphs.some((paragraph) => paragraph.pPrMark !== undefined)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A copied paragraph mark cannot coexist with the tracked paste mark in the paragraph's single mark slot.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          }),
        );
      }
      const tail = paragraphs.at(-1);
      if (
        mode.type === "suggesting" &&
        tail !== undefined &&
        (tail.sectionProperties !== undefined ||
          tail.pPrMark !== undefined ||
          (tail.propertyChanges?.length ?? 0) > 0)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A tracked paste cannot attach copied paragraph review or section metadata to a retained destination mark without a separately rejectable mark change.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          }),
        );
      }
      const leading = paragraphs.slice(0, -1);
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          sourceContainerPolicy: "separate",
          from: intent.from,
          to: intent.to,
          revision: mode.revision,
          newIds: mode.newIds,
          seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
          replacement: {
            paragraphs: leading.map(
              ({ sectionProperties: _section, ...paragraph }) => paragraph,
            ),
            tail: { openStart: 0, openEnd: 0, content: tail?.content ?? [] },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        ops = planned.value;
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        // The pasted content follows the struck selection: the caret ends
        // where the pasted tail meets the content after the selection.
        selection = rangeEndAfterDeletion({
          before: document,
          after: applied.value.document,
          from: intent.from,
          to: intent.to,
        });
      } else {
        const deleted = compileEditorIntent(document, {
          intent: {
            type: "replaceText",
            from: intent.from,
            to: intent.to,
            text: "",
          },
          mode,
        });
        if (deleted.isErr()) return deleted;
        ops = [...deleted.value.ops];
        let at = {
          ...deleted.value.selection,
          zeroWidthBefore:
            deleted.value.selection.zeroWidthBefore ??
            defaultInsertionGap(destination.content, intent.from.offset)
              .zeroWidthBefore,
        };
        for (const paragraph of leading) {
          if (paragraph.content.length > 0)
            ops.push({
              type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
              at,
              slice: { openStart: 0, openEnd: 0, content: paragraph.content },
              seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
              ...allocationFields,
            });
          const {
            type: _type,
            paraId,
            content: _content,
            sectionProperties: _section,
            pPrMark: _mark,
            ...newParagraph
          } = paragraph;
          ops.push({
            type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
            at: {
              ...at,
              ...gapAfterInserted(
                { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore ?? 0 },
                paragraph.content,
              ),
            },
            newBlockId: paraId ?? "",
            newHalf: SPLIT_HALVES.FIRST,
            newParagraph,
            ...allocationFields,
          });
          at = {
            story: at.story,
            blockId: at.blockId,
            offset: 0,
            zeroWidthBefore: 0,
          };
        }
        if (tail !== undefined && tail.content.length > 0)
          ops.push({
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice: { openStart: 0, openEnd: 0, content: tail.content },
            ...allocationFields,
          });
        selection = {
          ...at,
          ...gapAfterInserted(
            { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore ?? 0 },
            tail?.content ?? [],
          ),
        };
      }
      if (
        mode.type === "editing" &&
        intent.openStart === 1 &&
        leading.length > 0
      ) {
        const first = leading.at(0);
        if (first?.paraId === undefined)
          panic("An identified leading clipboard paragraph must have an id.");
        editedSeams.push({
          story: intent.from.story,
          blockId: first.paraId,
          offset: intent.from.offset,
        });
      }
      // Open edges and retained suffixes explicitly carry destination
      // properties; closed pasted marks carry their source properties. A
      // range deletion may have joined paragraphs with different formatting.
      if (tail !== undefined) {
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const trailing = editorParagraphGroups(
          applied.value.document,
          selection.story,
        )
          .find(({ paragraphs: group }) =>
            group.some(
              ({ paraId }) => idKey(paraId ?? "") === idKey(selection.blockId),
            ),
          )
          ?.paragraphs.at(-1);
        const trailingMarkAt =
          trailing?.paraId === undefined
            ? selection
            : { ...selection, blockId: trailing.paraId, offset: 0 };
        const current = paragraphAt(applied.value.document, trailingMarkAt);
        const patch = Object.fromEntries([
          ...Object.keys(current?.formatting ?? {}).map((key) => [key, null]),
          ...Object.entries(tail.formatting ?? {}),
        ]);
        const formatting = compileEditorIntent(applied.value.document, {
          intent: { type: "formatParagraph", at: trailingMarkAt, patch },
          mode: modeAfterOps(applied.value.document, mode),
        });
        if (formatting.isErr()) return formatting;
        ops.push(...formatting.value.ops);
      }
      // Closed pasted marks are applied after structural allocation, when
      // their exact destination paragraph and existing review are known.
      for (const [index, paragraph] of paragraphs.entries()) {
        if (
          paragraph.pPrMark === undefined &&
          (paragraph.propertyChanges?.length ?? 0) === 0 &&
          paragraph.sectionProperties === undefined
        )
          continue;
        const blockId =
          index === paragraphs.length - 1
            ? selection.blockId
            : (paragraph.paraId ?? "");
        const applied = applyDocumentOps(document, ops);
        if (applied.isErr()) return Result.err(applied.error);
        const current = paragraphAt(applied.value.document, {
          story: selection.story,
          blockId,
          offset: 0,
        });
        if (current === undefined)
          panic("An allocated clipboard paragraph must remain present.");
        if (paragraph.pPrMark !== undefined && current.pPrMark !== undefined) {
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
              message:
                "A copied paragraph mark cannot replace the destination's pending paragraph mark.",
              opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            }),
          );
        }
        const expected = reviewFieldsOf(current);
        const copiedChange = paragraph.propertyChanges?.at(0);
        const currentChange = current.propertyChanges?.at(0);
        if (
          copiedChange !== undefined &&
          currentChange !== undefined &&
          packageIdentityKeys(document.package).includes(
            `${IDENTITY_SPACES.REVISION}:${currentChange.info.id}`,
          )
        )
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.REVISION_CONFLICT,
              message:
                "Copied paragraph review cannot replace a pre-existing destination review.",
              opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            }),
          );
        const changes = [
          ...(paragraph.propertyChanges ?? current.propertyChanges ?? []),
        ];
        if (paragraph.pPrMark !== undefined || copiedChange !== undefined) {
          const review = reviewFieldsOf(current);
          if (changes.length > 0) review.propertyChanges = changes;
          if (paragraph.pPrMark !== undefined)
            review.pPrMark = paragraph.pPrMark;
          ops.push({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
            story: selection.story,
            blockId,
            expected,
            review,
          });
        }
        if (paragraph.sectionProperties !== undefined) {
          ops.push({
            type: DOCUMENT_OP_TYPES.SET_SECTION_ENDPOINT,
            endpoint: { type: "paragraph", blockId },
            expected:
              current.sectionProperties === undefined
                ? {
                    type: Object.hasOwn(current, "sectionProperties")
                      ? "undefined"
                      : "omitted",
                  }
                : { type: "present", value: current.sectionProperties },
            properties: { type: "present", value: paragraph.sectionProperties },
          });
        }
      }
      if (mode.type === "editing" && tail !== undefined) {
        const width = paragraphLength(tail);
        if (width > 0)
          for (const offset of new Set([
            selection.offset,
            selection.offset - width,
          ]))
            if (offset >= 0) editedSeams.push({ ...selection, offset });
      }
      // A closed paragraph payload ends outside that paragraph. Match the
      // forward text-selection affinity at its closing boundary when a next
      // untouched paragraph provides a text position; at document end the
      // caret remains at the payload's own last text position.
      if (
        intent.openEnd === 0 &&
        endpoint !== undefined &&
        intent.to.offset === paragraphLength(endpoint)
      ) {
        const bodyParagraphs = storyParagraphs(
          storyBody(document, intent.to.story),
        );
        const endLocation = bodyParagraphs.find(
          ({ paragraph }) =>
            idKey(paragraph.paraId ?? "") === idKey(intent.to.blockId),
        );
        const following = bodyParagraphs.find(
          (location) =>
            endLocation !== undefined &&
            sameBlockList(location.list, endLocation.list) &&
            location.index === endLocation.index + 1,
        );
        if (following?.paragraph.paraId !== undefined) {
          selection = {
            story: intent.to.story,
            blockId: following.paragraph.paraId,
            ...defaultInsertionGap(following.paragraph.content, 0),
          };
        }
      }
      break;
    }
    case "moveFragment": {
      const range = selectedParagraphRuns(document, intent.from, intent.to);
      if (range.isErr()) return Result.err(range.error);
      const sourceParagraph = paragraphAt(document, intent.from);
      const fromGap =
        intent.from.zeroWidthBefore ??
        (sourceParagraph === undefined
          ? 0
          : zeroWidthLeavesAt(sourceParagraph.content, intent.from.offset)
              .length);
      const toGap = intent.to.zeroWidthBefore ?? 0;
      if (
        idKey(intent.from.blockId) === idKey(intent.to.blockId) &&
        intent.from.offset === intent.to.offset &&
        fromGap >= toGap
      ) {
        if (
          fromGap > toGap &&
          (intent.from.zeroWidthBefore !== undefined ||
            intent.to.zeroWidthBefore !== undefined)
        ) {
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
              message: "The move source marker gaps are reversed.",
              opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
            }),
          );
        }
        return Result.ok({ ops: [], selection: intent.target });
      }
      const locations = storyParagraphs(storyBody(document, intent.from.story));
      const indexOf = (at: TextPosition) =>
        locations.findIndex(
          ({ paragraph }) =>
            idKey(paragraph.paraId ?? "") === idKey(at.blockId),
        );
      const first = indexOf(intent.from);
      const last = indexOf(intent.to);
      const targetIndex = indexOf(intent.target);
      const targetGap =
        intent.target.zeroWidthBefore ??
        defaultInsertionGap(
          paragraphAt(document, intent.target)?.content ?? [],
          intent.target.offset,
        ).zeroWidthBefore;
      const targetPosition = {
        offset: intent.target.offset,
        zeroWidthBefore: targetGap,
      };
      if (
        intent.target.story === intent.from.story &&
        targetIndex >= first &&
        targetIndex <= last &&
        (targetIndex !== first ||
          compareGaps(targetPosition, {
            offset: intent.from.offset,
            zeroWidthBefore: fromGap,
          }) >= 0) &&
        (targetIndex !== last ||
          compareGaps(targetPosition, {
            offset: intent.to.offset,
            zeroWidthBefore: toGap,
          }) <= 0)
      ) {
        // Moving into the selected range is an explicit no-op, including its edges.
        return Result.ok({ ops: [], selection: intent.target });
      }
      const anchorDeletions: DeleteRangeOp[] = [];
      const sourceAnchors = mode.type === "editing" ? range.value.flat() : [];
      for (const { paragraph } of sourceAnchors) {
        const blockId = paragraph.paraId ?? "";
        const from =
          idKey(blockId) === idKey(intent.from.blockId)
            ? { offset: intent.from.offset, zeroWidthBefore: fromGap }
            : { offset: 0, zeroWidthBefore: 0 };
        const to =
          idKey(blockId) === idKey(intent.to.blockId)
            ? { offset: intent.to.offset, zeroWidthBefore: toGap }
            : {
                offset: paragraphLength(paragraph),
                zeroWidthBefore: zeroWidthLeavesAt(
                  paragraph.content,
                  paragraphLength(paragraph),
                ).length,
              };
        for (const span of leafSpans(paragraph.content)) {
          if (
            !isMoveAnchor(span.node) ||
            span.ancestors.some(isRemovedRevisionNode) ||
            compareGaps(span.before, from) < 0 ||
            compareGaps(span.after, to) > 0
          )
            continue;
          anchorDeletions.push({
            type: DOCUMENT_OP_TYPES.DELETE_RANGE,
            from: { story: intent.from.story, blockId, ...span.before },
            to: { story: intent.from.story, blockId, ...span.after },
          });
        }
      }
      const deleted = compileEditorIntent(document, {
        intent: {
          type: "replaceText",
          from: intent.from,
          to: intent.to,
          text: "",
        },
        mode,
      });
      if (deleted.isErr()) return deleted;
      const logicalDeletions =
        anchorDeletions.length === 0
          ? deleted.value.ops
          : [
              ...deleted.value.ops
                .filter((op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE)
                .concat(anchorDeletions)
                .sort(
                  (left, right) =>
                    indexOf(right.from) - indexOf(left.from) ||
                    compareGaps(
                      {
                        offset: right.from.offset,
                        zeroWidthBefore: right.from.zeroWidthBefore ?? 0,
                      },
                      {
                        offset: left.from.offset,
                        zeroWidthBefore: left.from.zeroWidthBefore ?? 0,
                      },
                    ),
                ),
              ...deleted.value.ops.filter(
                (op) => op.type !== DOCUMENT_OP_TYPES.DELETE_RANGE,
              ),
            ];
      const anchors = new Set(anchorDeletions);
      const deletionOps: DocumentOp[] = [];
      const movedNoteBodies: DocumentOp[] = [];
      let afterDeletion = document;
      for (const planned of logicalDeletions) {
        let op = planned;
        if (
          planned.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
          anchors.has(planned)
        ) {
          const paragraph = paragraphAt(afterDeletion, planned.from);
          if (paragraph === undefined)
            panic("A move anchor lost its selected paragraph.");
          op = {
            type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
            story: planned.from.story,
            blockId: planned.from.blockId,
            expected: paragraph.content,
            content: deleteBetween(
              paragraph.content,
              {
                offset: planned.from.offset,
                zeroWidthBefore: planned.from.zeroWidthBefore ?? 0,
              },
              {
                offset: planned.to.offset,
                zeroWidthBefore: planned.to.zeroWidthBefore ?? 0,
              },
            ).content,
          };
        }
        const applied = applyDocumentOps(afterDeletion, [op]);
        if (applied.isErr()) return Result.err(applied.error);
        afterDeletion = applied.value.document;
        deletionOps.push(op);
        for (const inverse of applied.value.inverse) {
          if (
            inverse.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS &&
            (inverse.parts.footnotes !== undefined ||
              inverse.parts.endnotes !== undefined)
          )
            movedNoteBodies.push({
              type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
              expected: {
                ...(inverse.expected.footnotes === undefined
                  ? {}
                  : { footnotes: inverse.expected.footnotes }),
                ...(inverse.expected.endnotes === undefined
                  ? {}
                  : { endnotes: inverse.expected.endnotes }),
              },
              parts: {
                ...(inverse.parts.footnotes === undefined
                  ? {}
                  : { footnotes: inverse.parts.footnotes }),
                ...(inverse.parts.endnotes === undefined
                  ? {}
                  : { endnotes: inverse.parts.endnotes }),
              },
            });
        }
      }
      if (movedNoteBodies.length > 0) {
        // Each restoration expects the note parts its own deletion left, so
        // they undo the deletions newest first.
        const restorations = movedNoteBodies.toReversed();
        const restored = applyDocumentOps(afterDeletion, restorations);
        if (restored.isErr()) return Result.err(restored.error);
        afterDeletion = restored.value.document;
        deletionOps.push(...restorations);
      }
      let target = intent.target;
      for (const op of logicalDeletions) {
        if (
          op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.from.blockId) &&
          target.story === op.from.story &&
          target.offset >= op.to.offset
        ) {
          const source = paragraphAt(document, op.from);
          const deletionFromGap =
            op.from.zeroWidthBefore ??
            zeroWidthLeavesAt(source?.content ?? [], op.from.offset).length;
          const deletionToGap = op.to.zeroWidthBefore ?? 0;
          const deletionTargetGap =
            target.zeroWidthBefore ??
            defaultInsertionGap(source?.content ?? [], target.offset)
              .zeroWidthBefore;
          const rebased = Object.assign({}, target, {
            offset: target.offset - (op.to.offset - op.from.offset),
          });
          if (target.offset === op.to.offset)
            rebased.zeroWidthBefore =
              deletionFromGap + deletionTargetGap - deletionToGap;
          target = rebased;
        }
        if (
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.blockId) &&
          target.story === op.story
        )
          target = Object.assign({}, target, { blockId: op.nextBlockId });
        else if (
          op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
          op.revision === undefined &&
          idKey(target.blockId) === idKey(op.nextBlockId) &&
          target.story === op.story
        ) {
          const prefix = paragraphAt(document, {
            story: op.story,
            blockId: op.blockId,
            offset: 0,
          });
          const remaining =
            idKey(op.blockId) === idKey(intent.from.blockId)
              ? intent.from.offset
              : 0;
          if (prefix !== undefined) {
            const rebased = Object.assign({}, target, {
              offset: target.offset + remaining,
            });
            if (
              target.offset === 0 &&
              idKey(op.blockId) === idKey(intent.from.blockId)
            )
              rebased.zeroWidthBefore = fromGap + (target.zeroWidthBefore ?? 0);
            target = rebased;
          }
        }
      }
      const paragraphs =
        mode.type === "suggesting"
          ? intent.paragraphs.map((paragraph) =>
              Object.assign({}, paragraph, {
                content: asParagraphContent(
                  withoutMoveAnchors(paragraph.content),
                ),
              }),
            )
          : intent.paragraphs;
      const inserted = compileEditorIntent(afterDeletion, {
        intent: {
          type: "replaceFragment",
          from: target,
          to: target,
          paragraphs,
          openStart: intent.openStart,
          openEnd: intent.openEnd,
        },
        mode: modeAfterOps(afterDeletion, mode),
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
      });
      if (inserted.isErr()) return inserted;
      ops = [...deletionOps, ...inserted.value.ops];
      selection = inserted.value.selection;
      break;
    }
    case "table": {
      const operation = intent.operation;
      const located = locateTableRow(document, operation);
      if (located.isErr()) return Result.err(located.error);
      const demand = tableEditParagraphDemand(located.value, operation);
      if (demand.isErr()) return Result.err(demand.error);
      const fresh = freshParagraphIds({
        document,
        reserved: mode.reservedBlockIds,
        lease: mode.idLease,
        count: demand.value,
        opType: operation.type,
      });
      if (fresh.isErr()) return Result.err(fresh.error);
      const newBlockIds = fresh.value;
      let op: TableEditOp;
      switch (operation.type) {
        case DOCUMENT_OP_TYPES.INSERT_COLUMN:
        case DOCUMENT_OP_TYPES.MERGE_CELLS:
        case DOCUMENT_OP_TYPES.SPLIT_CELL:
          op = { ...operation, newBlockIds, ...tracked };
          break;
        default:
          op = { ...operation, ...tracked };
      }
      let blockId = op.blockId;
      if (op.type === DOCUMENT_OP_TYPES.DELETE_COLUMN) {
        const grid = tableGrid(located.value.table, op.type);
        if (grid.isErr()) return Result.err(grid.error);
        const rows = located.value.table.rows.map((row, index) => ({
          ...row,
          cells: row.cells.filter((_cell, cellIndex) => {
            const entry = grid.value.rows[index]?.find(
              (cell) => cell.index === cellIndex,
            );
            return (
              entry !== undefined &&
              (entry.start > op.column ||
                entry.end <= op.column ||
                entry.end - entry.start > 1)
            );
          }),
        }));
        const survivor =
          tableRowAnchor(rows) ??
          storyParagraphs(storyBody(document, op.story)).find((paragraph) =>
            sameBlockList(paragraph.list, located.value.list),
          )?.paragraph.paraId;
        if (survivor === undefined)
          return Result.err(
            new DocumentOpRefusal({
              opType: op.type,
              reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
              message: "Deleting the final column requires deleting the table.",
            }),
          );
        blockId = survivor;
      }
      const applied = applyDocumentOp(document, op);
      if (applied.isErr()) return Result.err(applied.error);
      ops = [
        trimAppliedNewIds({
          before: document,
          op,
          applied: applied.value,
          story: op.story,
        }),
      ];
      selection = { story: op.story, blockId, offset: 0 };
      if (intent.cellContent === undefined) break;
      if (
        op.type !== DOCUMENT_OP_TYPES.INSERT_COLUMN ||
        intent.cellContent.length > newBlockIds.length
      )
        return Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            message: "Cell content needs a cell the column insertion creates.",
          }),
        );
      // Each new cell holds one empty paragraph, its fresh id in row order;
      // the content goes into it as a fragment typed there.
      let filled = applied.value.document;
      for (const [index, paragraphs] of intent.cellContent.entries()) {
        const cellBlockId = newBlockIds[index];
        if (
          cellBlockId === undefined ||
          paragraphs.length === 0 ||
          (paragraphs.length === 1 && paragraphs[0]?.content.length === 0)
        )
          continue;
        const at = { story: op.story, blockId: cellBlockId, offset: 0 };
        const fill = compileEditorIntent(filled, {
          intent: {
            type: "replaceFragment",
            from: at,
            to: at,
            paragraphs,
            openStart: 1,
            openEnd: 1,
          },
          mode: {
            ...modeAfterOps(filled, mode),
            ...(mode.reservedBlockIds === undefined
              ? {}
              : { reservedBlockIds: mode.reservedBlockIds }),
          },
        });
        if (fill.isErr()) return Result.err(fill.error);
        const written = applyDocumentOps(filled, fill.value.ops);
        if (written.isErr()) return Result.err(written.error);
        filled = written.value.document;
        ops.push(...fill.value.ops);
      }
      break;
    }
    case "tableRow": {
      const operation = intent.operation;
      const located = locateTableRow(document, operation);
      if (located.isErr()) return Result.err(located.error);
      const { table, rowIndex, list } = located.value;
      let op: InsertRowOp | DeleteRowOp;
      let blockId: string | undefined;
      if (operation.type === DOCUMENT_OP_TYPES.INSERT_ROW) {
        const fresh = freshParagraphIds({
          document,
          reserved: mode.reservedBlockIds,
          lease: mode.idLease,
          count: rowParagraphCount(operation.row),
          opType: operation.type,
        });
        if (fresh.isErr()) return Result.err(fresh.error);
        const row = withFreshParagraphIds(operation.row, fresh.value);
        op = { ...operation, row, ...tracked };
        blockId = fresh.value[0];
      } else {
        op = { ...operation, ...tracked };
        // The caret lands in a row that stays, or beside the table when its
        // last row goes with it. A direct deletion's rows are those the
        // removal leaves, a merge it starts handed on to the next row.
        blockId =
          tableRowAnchor(
            op.revision === undefined
              ? rowsWithoutRow(table, rowIndex)
              : table.rows.filter((_row, index) => index !== rowIndex),
          ) ??
          storyParagraphs(storyBody(document, op.story)).find(
            (paragraph) =>
              sameBlockList(paragraph.list, list) &&
              paragraph.paragraph.paraId !== undefined,
          )?.paragraph.paraId;
      }
      if (blockId === undefined)
        return Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            message: "The row edit leaves no paragraph for the caret.",
          }),
        );
      const applied = applyDocumentOp(document, op);
      if (applied.isErr()) return Result.err(applied.error);
      ops = [
        trimAppliedNewIds({
          before: document,
          op,
          applied: applied.value,
          story: op.story,
        }),
      ];
      selection = { story: op.story, blockId, offset: 0 };
      break;
    }
    case "deleteParagraphs": {
      const first = intent.blockIds[0];
      if (first === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
            message: "No paragraphs to delete.",
            opType: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
          }),
        );
      return Result.ok({
        ops: [
          {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: intent.story,
            blockIds: intent.blockIds,
            ...tracked,
            ...(mode.type === "suggesting" && intent.move === true
              ? { move: true as const }
              : {}),
          },
        ],
        selection:
          mode.type === "editing"
            ? directDeletionCaret(document, intent)
            : { story: intent.story, blockId: first, offset: 0 },
      });
    }
    case "insertParagraphs": {
      const fresh = freshParagraphIds({
        document,
        reserved: mode.reservedBlockIds,
        lease: mode.idLease,
        count: intent.paragraphs.length,
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      });
      if (fresh.isErr()) return Result.err(fresh.error);
      const blocks = intent.paragraphs.map((paragraph, index) => ({
        ...paragraph,
        paraId: fresh.value[index] ?? panic("A fresh paragraph id is missing."),
      }));
      const first = blocks[0];
      if (first === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
            message: "No paragraphs to insert.",
            opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          }),
        );
      return Result.ok({
        ops: [
          ...newListOps(intent.newList),
          {
            type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
            story: intent.story,
            at: intent.at,
            blocks,
            ...tracked,
            ...(mode.type === "suggesting" && intent.move === true
              ? { move: true as const }
              : {}),
          },
        ],
        selection: { story: intent.story, blockId: first.paraId, offset: 0 },
      });
    }
    case "tableBlock": {
      if (intent.operation.type === DOCUMENT_OP_TYPES.INSERT_TABLE) {
        const { story, blockId, table } = intent.operation;
        const separated = intent.operation.separated === true;
        const trailing = intent.operation.paragraphAfter !== "none";
        const fresh = freshParagraphIds({
          document,
          reserved: mode.reservedBlockIds,
          lease: mode.idLease,
          count:
            storyParagraphs({ content: [{ type: "table", rows: table.rows }] })
              .length +
            (separated ? 1 : 0) +
            (trailing ? 1 : 0),
          opType: DOCUMENT_OP_TYPES.INSERT_TABLE,
        });
        if (fresh.isErr()) return Result.err(fresh.error);
        const leadingId = separated ? fresh.value[0] : undefined;
        const afterLeading = separated ? fresh.value.slice(1) : fresh.value;
        const paragraphId = trailing ? afterLeading.at(0) : undefined;
        const tableIds = trailing ? afterLeading.slice(1) : afterLeading;
        if (
          (trailing && paragraphId === undefined) ||
          (separated && leadingId === undefined)
        )
          return panic("A fresh paragraph id is missing.");
        let next = 0;
        const rows = table.rows.map((row: TableRow) => {
          const count = storyParagraphs({
            content: [{ type: "table", rows: [row] }],
          }).length;
          const ids = tableIds.slice(next, next + count);
          next += count;
          return withFreshParagraphIds(row, ids);
        });
        const before = intent.operation.side === "before";
        const paragraphs = [leadingId, paragraphId].flatMap((paraId) =>
          paraId === undefined
            ? []
            : [{ type: "paragraph" as const, paraId, content: [] }],
        );
        // The empty paragraphs first, then the table ahead of the one after
        // it: after the anchor, or before it at the story's start.
        const sequence: DocumentOp[] =
          paragraphs.length === 0
            ? []
            : [
                {
                  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
                  story,
                  at: { type: before ? "before" : "after", blockId },
                  blocks: paragraphs,
                  ...tracked,
                },
              ];
        sequence.push({
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story,
          at: before
            ? { type: "before", blockId: paragraphId ?? blockId }
            : { type: "after", blockId },
          table: { ...table, rows },
          ...tracked,
        });
        let current = document;
        ops = [];
        // The two operations share one pool of revision ids: each takes the
        // next ids the one before it left unused.
        let pool =
          mode.type === "suggesting"
            ? [mode.revision.id, ...(mode.newIds?.revision ?? [])]
            : [];
        for (const planned of sequence) {
          const [id, ...rest] = pool;
          const op =
            mode.type === "suggesting" && id !== undefined
              ? {
                  ...planned,
                  revision: { ...mode.revision, id },
                  newIds: {
                    revision: rest,
                    control: mode.newIds?.control ?? [],
                  },
                }
              : planned;
          const applied = applyDocumentOp(current, op);
          if (applied.isErr()) return Result.err(applied.error);
          const trimmed = trimAppliedNewIds({
            before: current,
            op,
            applied: applied.value,
            story,
          });
          ops.push(trimmed);
          const used = new Set(applied.value.revisions ?? []);
          pool = pool.filter((candidate) => !used.has(candidate));
          current = applied.value.document;
        }
        const first = tableIds.at(0);
        if (first === undefined)
          return panic("The inserted table has no paragraph.");
        selection = { story, blockId: first, offset: 0 };
        break;
      }
      const op = { ...intent.operation, ...tracked };
      const located = locateTableRow(document, op);
      if (located.isErr()) return Result.err(located.error);
      // The caret lands beside the table: after it, or before it at the end
      // of its block list.
      const { list } = located.value;
      const beside = storyParagraphs(storyBody(document, op.story)).filter(
        ({ list: other, paragraph }) =>
          sameBlockList(other, list) && paragraph.paraId !== undefined,
      );
      const position = beside.findIndex(
        ({ index }) => index > located.value.index,
      );
      const anchor = (beside[position] ?? beside.at(-1))?.paragraph.paraId;
      if (anchor === undefined)
        return Result.err(
          new DocumentOpRefusal({
            opType: op.type,
            reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
            message: "Removing the table leaves no paragraph for the caret.",
          }),
        );
      const applied = applyDocumentOp(document, op);
      if (applied.isErr()) return Result.err(applied.error);
      ops = [
        trimAppliedNewIds({
          before: document,
          op,
          applied: applied.value,
          story: op.story,
        }),
      ];
      selection = { story: op.story, blockId: anchor, offset: 0 };
      break;
    }
    case "setList": {
      const first = intent.items.at(0);
      if (first === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "A list intent must address a paragraph.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          }),
        );
      const numId =
        intent.target.type === "existing"
          ? intent.target.numId
          : intent.target.num.numId;
      if (
        !Number.isInteger(numId) ||
        !isNumberingReference(numId) ||
        numId < NO_NUMBERING_NUM_ID ||
        numId > MAX_REVISION_ID ||
        intent.items.some(({ ilvl }) => !isNumberingLevel(ilvl))
      )
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_NEW_ID,
            message: "List references require a valid numbering id and level.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          }),
        );
      const creation =
        intent.target.type === "new"
          ? [
              {
                type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
                num: intent.target.num,
                ...(intent.target.abstractNum === undefined
                  ? {}
                  : { abstractNum: intent.target.abstractNum }),
              } as const,
            ]
          : [];
      const itemsByMark = new Map<string, (typeof intent.items)[number]>();
      const paragraphMarkOwner = createParagraphMarkOwnerLookup(document);
      for (const { at, ilvl } of intent.items) {
        const blockId = paragraphMarkOwner(at)?.paraId ?? at.blockId;
        itemsByMark.set(`${JSON.stringify(at.story)}:${idKey(blockId)}`, {
          at: Object.assign({}, at, { blockId }),
          ilvl,
        });
      }
      const paragraphOps = Array.from(
        itemsByMark.values(),
        ({ at, ilvl }) =>
          ({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: at.story,
            blockId: at.blockId,
            patch: { numPr: paragraphNumberingReference({ numId, ilvl }) },
            ...(mode.type === "suggesting"
              ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
              : {}),
            ...(mode.type === "suggesting" ? { revision: mode.revision } : {}),
          }) as const,
      );
      if (mode.type === "editing") ops = [...creation, ...paragraphOps];
      else {
        // Numbering definitions are package resources, not OOXML revision records.
        // Both modes allocate the same resource; the paragraph property is tracked.
        const created = applyDocumentOps(document, creation);
        if (created.isErr()) return Result.err(created.error);
        const plan = createTrackedPlan({
          document: created.value.document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        for (const op of paragraphOps) {
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = [...creation, ...plan.ops];
      }
      selection = first.at;
      break;
    }
    case "formatParagraph": {
      ops = [
        {
          type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
          story: intent.at.story,
          blockId:
            createParagraphMarkOwnerLookup(document)(intent.at)?.paraId ??
            intent.at.blockId,
          patch: intent.patch,
          ...(mode.type === "suggesting"
            ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
            : {}),
          ...(mode.type === "suggesting" ? { revision: mode.revision } : {}),
        },
      ];
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({
          document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        for (const op of ops) {
          if (op.type !== DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS)
            panic(
              "A paragraph-format intent compiled to a different operation.",
            );
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = plan.ops;
      }
      ops = [...newListOps(intent.newList), ...ops];
      selection = intent.at;
      break;
    }
    case "formatRun": {
      const partitioned = selectedParagraphRuns(
        document,
        intent.from,
        intent.to,
      );
      if (partitioned.isErr()) return Result.err(partitioned.error);
      ops = partitioned.value.flatMap((run) =>
        run.map(({ paragraph }) => {
          const blockId = paragraph.paraId ?? "";
          return {
            type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
            from:
              idKey(blockId) === idKey(intent.from.blockId)
                ? intent.from
                : { story: intent.from.story, blockId, offset: 0 },
            to:
              idKey(blockId) === idKey(intent.to.blockId)
                ? intent.to
                : {
                    story: intent.to.story,
                    blockId,
                    offset: paragraphLength(paragraph),
                  },
            patch: intent.patch,
            ...(mode.type === "suggesting"
              ? { propertyReview: PROPERTY_REVIEW_POLICIES.APPEND }
              : {}),

            ...tracked,
          } as const;
        }),
      );
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({
          document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        for (const op of ops) {
          if (op.type !== DOCUMENT_OP_TYPES.SET_RUN_PROPS)
            panic("A run-format intent compiled to a different operation.");
          const appended = plan.append(op);
          if (appended.isErr()) return Result.err(appended.error);
        }
        ops = plan.ops;
      }
      selection = intent.to;
      break;
    }
    case "insertAtom": {
      const formatting = intentRunFormatting(document, intent);
      if (mode.type === "suggesting") {
        const planned = planTrackedReplace(document, {
          sourceContainerPolicy: "join",
          from: intent.from,
          to: intent.to,
          revision: mode.revision,
          newIds: mode.newIds,
          replacement: {
            paragraphs: [],
            tail: {
              openStart: 0,
              openEnd: 0,
              content: [
                {
                  type: "run",
                  ...(formatting === undefined ? {} : { formatting }),
                  content: [intent.atom],
                },
              ],
            },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        ops = planned.value;
        const insertion = ops.findLast(
          (op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT,
        );
        selection =
          insertion?.type === DOCUMENT_OP_TYPES.INSERT_CONTENT
            ? { ...insertion.at, offset: insertion.at.offset + 1 }
            : intent.from;
        break;
      }
      const deletion = compileEditorIntent(document, {
        intent: {
          type: "replaceText",
          from: intent.from,
          to: intent.to,
          text: "",
        },
        mode,
      });
      if (deletion.isErr()) return deletion;
      const at = deletion.value.selection;
      const content = [
        {
          type: "run",
          ...(formatting === undefined ? {} : { formatting }),
          content: [intent.atom],
        },
      ] satisfies ParagraphContent[];
      const insertion = {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at,
        slice: {
          openStart: 0,
          openEnd: 0,
          content,
        },
        ...tracked,
      } as const satisfies DocumentOp;
      ops = [...deletion.value.ops, insertion];
      selection = { ...at, offset: at.offset + 1 };
      break;
    }
    case "replaceText": {
      const { text } = intent;
      let { from, to } = intent;
      const paragraph = paragraphAt(document, from);
      if (paragraph === undefined) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
            message: "The input paragraph does not exist.",
            opType: DOCUMENT_OP_TYPES.INSERT_TEXT,
          }),
        );
      }
      if (idKey(from.blockId) === idKey(to.blockId)) {
        const fromGap = {
          offset: from.offset,
          zeroWidthBefore:
            from.zeroWidthBefore ??
            zeroWidthLeavesAt(paragraph.content, from.offset).length,
        };
        const toGap = {
          offset: to.offset,
          zeroWidthBefore: to.zeroWidthBefore ?? 0,
        };
        // Every leaf the selection touches, including a partly selected text leaf.
        const selected = leafSpans(paragraph.content).filter(
          ({ before, after }) =>
            compareGaps(before, toGap) < 0 && compareGaps(fromGap, after) < 0,
        );
        const marker = selected.find(({ node }) => node.type === "noteMarker");
        if (
          marker !== undefined &&
          selected.every(
            ({ node, ancestors }) =>
              node.type === "noteMarker" ||
              ancestors.some(isRemovedRevisionNode),
          )
        ) {
          // Only an empty replacement is a no-op: typed text over the protected marker
          // is an insertion right after it, in both modes.
          if (text.length === 0)
            return Result.err(
              new DocumentOpRefusal({
                reason: DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
                message:
                  "The range contains only the note story's automatic marker.",
                opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
              }),
            );
          from = { ...from, ...marker.after };
          to = from;
        }
      }
      const runProps = intentRunFormatting(document, intent);
      // Both modes insert the same authored run; source XML attributes belong
      // to the existing run rather than to newly typed text.
      const content = (
        text === ""
          ? []
          : [
              {
                type: "run",
                ...(runProps === undefined ? {} : { formatting: runProps }),
                content: [{ type: "text", text }],
              },
            ]
      ) satisfies ParagraphContent[];
      const insertionOnly =
        idKey(from.blockId) === idKey(to.blockId) &&
        from.offset === to.offset &&
        (from.zeroWidthBefore ?? 0) === (to.zeroWidthBefore ?? 0);
      const continued =
        insertionOnly && text !== ""
          ? continuation(paragraph.content, from, to, {
              ...(intent.inside === undefined ? {} : { inside: intent.inside }),
            })
          : undefined;
      const continuing =
        continued === undefined
          ? undefined
          : continuedSlice(continued, content);
      if (mode.type === "suggesting" && continued !== undefined) {
        // Typed beside the author's own pending insertion: the text goes
        // into it, so consecutive typing reads as one change.
        const at = {
          ...from,
          zeroWidthBefore:
            from.zeroWidthBefore ??
            defaultInsertionGap(paragraph.content, from.offset).zeroWidthBefore,
        };
        const slice = ownInsertionSlice({
          items: paragraph.content,
          gap: { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore },
          continued: continued.containers,
          author: mode.revision.author,
          content,
        });
        if (slice !== undefined) {
          const insert: InsertContentOp = {
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice,
            ...(mode.newIds === undefined ? {} : { newIds: mode.newIds }),
          };
          const inserted = applyDocumentOp(document, insert);
          if (inserted.isErr()) return Result.err(inserted.error);
          // Only the ids a cut record took stay with the insertion; the
          // insertion applies the same with only them.
          ops = [
            trimAppliedNewIds({
              before: document,
              op: insert,
              applied: inserted.value,
              story: at.story,
            }),
          ];
          const steps: AppliedDocumentOp[] = [inserted.value];
          // The typed run joins the alike runs of the insertion it meets.
          let current = inserted.value.document;
          for (const seam of [
            { offset: at.offset + text.length, zeroWidthBefore: 0 },
            { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore },
          ]) {
            const typedInto = paragraphAt(current, at);
            if (typedInto === undefined)
              panic("A typed paragraph must exist at its seam.");
            const depth = runSeamDepth({
              paragraph: typedInto,
              at: seam,
              policy: RUN_SEAM_POLICIES.plain,
            });
            if (depth === 0) continue;
            // A seam join, so a concurrent edit beside the typed run leaves
            // it apart rather than refusing it.
            const join = {
              type: DOCUMENT_OP_TYPES.JOIN_INLINE,
              at: { ...at, ...seam },
              depth,
              ifAlike: true,
            } as const;
            const joined = applyDocumentOp(current, join);
            if (joined.isErr()) return Result.err(joined.error);
            ops.push(join);
            steps.push(joined.value);
            current = joined.value.document;
          }
          handedBack = {
            ops: [...ops],
            result: {
              ...combineEdits(document, steps),
              revisions: steps.flatMap(({ revisions }) => revisions),
            },
          };
          selection = replacementCaret(at, text);
          break;
        }
      }
      if (mode.type === "suggesting" && continuing !== undefined) {
        // Typed inside a hyperlink, content control or wrapper: the tracked
        // insertion goes inside it too, rather than cutting it in two.
        const at = {
          ...from,
          zeroWidthBefore:
            from.zeroWidthBefore ??
            defaultInsertionGap(paragraph.content, from.offset).zeroWidthBefore,
        };
        ops = [
          {
            type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
            at,
            slice: continuing,
            seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
            revision: mode.revision,
            ...(mode.newIds === undefined ? {} : { newIds: mode.newIds }),
          },
        ];
        selection = replacementCaret(at, text);
        break;
      }
      if (mode.type === "suggesting") {
        // A deletion beside the author's own pending one continues it, and
        // so does the struck part of a replacement; the typed text keeps its
        // own stamp.
        const struck = groupedDeletionStamp({
          start: {
            items: paragraph.content,
            gap: {
              offset: from.offset,
              zeroWidthBefore:
                from.zeroWidthBefore ??
                zeroWidthLeavesAt(paragraph.content, from.offset).length,
            },
          },
          end: {
            items: (paragraphAt(document, to) ?? paragraph).content,
            gap: {
              offset: to.offset,
              zeroWidthBefore: to.zeroWidthBefore ?? 0,
            },
          },
          stamp: mode.revision,
        });
        // The grouped stamp keeps this edit's id, so only its record's
        // metadata changes; the plan applies the deletion with it, so the
        // typed text takes a fresh id whenever the deletion makes a record.
        const planned = planTrackedReplace(document, {
          sourceContainerPolicy: "join",
          from,
          to,
          revision: text === "" ? struck : mode.revision,
          ...(text !== "" && struck !== mode.revision
            ? { deletionRevision: struck }
            : {}),
          newIds: mode.newIds,
          seamPolicy: INSERTION_SEAM_POLICIES.MERGE_PLAIN_RUNS,
          replacement: {
            paragraphs: [],
            tail: {
              openStart: 0,
              openEnd: 0,
              content,
            },
          },
        });
        if (planned.isErr()) return Result.err(planned.error);
        // The typed text follows the tracked deletion, inside the content
        // controls and wrappers the direct replacement continues.
        const resumed =
          idKey(from.blockId) === idKey(to.blockId) && text !== ""
            ? continuation(paragraph.content, from, to, {
                ...(intent.inside === undefined
                  ? {}
                  : { inside: intent.inside }),
                struck: true,
              })
            : undefined;
        ops = [];
        for (const op of planned.value) {
          if (
            resumed === undefined ||
            op.type !== DOCUMENT_OP_TYPES.INSERT_CONTENT ||
            idKey(op.at.blockId) !== idKey(from.blockId) ||
            op.slice.openStart !== 0 ||
            op.slice.openEnd !== 0
          ) {
            ops.push(op);
            continue;
          }
          const slice = continuedSlice(resumed, op.slice.content);
          ops.push(slice === undefined ? op : { ...op, slice });
        }
        selection = { ...from, offset: from.offset + text.length };
        for (const op of ops) {
          if (
            op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS &&
            op.revision === undefined &&
            idKey(op.blockId) === idKey(selection.blockId)
          )
            selection = { ...selection, blockId: op.nextBlockId };
          if (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT)
            selection = replacementCaret(op.at, text);
        }
        break;
      }
      const replacesWholeParagraph =
        idKey(from.blockId) === idKey(to.blockId) &&
        from.offset === 0 &&
        to.offset === paragraphLength(paragraph) &&
        paragraphLength(paragraph) > 0 &&
        text.length > 0 &&
        (from.zeroWidthBefore ?? 0) > 0 &&
        replacementPreservesOutsideContent({ paragraph, from, to }) &&
        !(document.package.document.comments ?? []).some(
          ({ parentId, anchorKind }) =>
            parentId === undefined && anchorKind !== "range",
        );
      if (replacesWholeParagraph) {
        ops = [
          {
            type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
            story: from.story,
            blockId: from.blockId,
            expected: paragraph.content,
            content: [
              ...paragraph.content.filter(
                (node) => isParagraphContent(node) && isRangeBoundary(node),
              ),
              ...content,
              ...paragraph.content.filter(
                (node) =>
                  isParagraphContent(node) && node.type === "commentReference",
              ),
            ],
          },
        ];
        selection = { ...from, offset: text.length, zeroWidthBefore: 0 };
        break;
      }
      ops = [];
      let survivorId = from.blockId;
      const deletes =
        from.blockId !== to.blockId ||
        from.offset !== to.offset ||
        (from.zeroWidthBefore ?? 0) !== (to.zeroWidthBefore ?? 0);
      if (deletes) {
        const partitioned = selectedParagraphRuns(document, from, to);
        if (partitioned.isErr()) return Result.err(partitioned.error);
        survivorId =
          partitioned.value.at(0)?.at(-1)?.paragraph.paraId ?? from.blockId;
        for (const run of partitioned.value.toReversed()) {
          const last = run.at(-1)?.paragraph;
          if (!last)
            panic("A selected paragraph run must contain its survivor.");
          for (const { paragraph: item } of run.toReversed()) {
            const blockId = item.paraId ?? "";
            const start =
              idKey(blockId) === idKey(from.blockId)
                ? from
                : { story: from.story, blockId, offset: 0, zeroWidthBefore: 0 };
            const end =
              idKey(blockId) === idKey(to.blockId)
                ? to
                : {
                    story: from.story,
                    blockId,
                    offset: paragraphLength(item),
                    zeroWidthBefore: zeroWidthLeavesAt(
                      item.content,
                      paragraphLength(item),
                    ).length,
                  };
            const segments = replacementDeletionSegments({
              spans: leafSpans(item.content),
              from: {
                offset: start.offset,
                zeroWidthBefore:
                  start.zeroWidthBefore ??
                  zeroWidthLeavesAt(item.content, start.offset).length,
              },
              to: {
                offset: end.offset,
                zeroWidthBefore: end.zeroWidthBefore ?? 0,
              },
              mode: { type: "editing" },
            });
            for (const segment of segments.toReversed()) {
              ops.push({
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                ...allocationFields,
                from: { story: from.story, blockId, ...segment.from },
                to: { story: from.story, blockId, ...segment.to },
              });
            }
          }
          for (const { paragraph: item } of run.slice(0, -1).toReversed()) {
            ops.push({
              type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
              story: from.story,
              blockId: item.paraId ?? "",
              nextBlockId: last.paraId ?? "",
              sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
            });
          }
        }
      }
      // A replacement goes after what the deletion keeps in the range, where
      // accepting the tracked replacement puts it. A caret insertion resolves
      // its affinity in the input.
      let at: TextPosition = {
        ...from,
        blockId: survivorId,
        zeroWidthBefore:
          from.zeroWidthBefore ??
          defaultInsertionGap(paragraph.content, from.offset).zeroWidthBefore,
      };
      if (deletes) {
        const deleted = applyDocumentOps(document, ops);
        if (deleted.isErr()) return Result.err(deleted.error);
        at = rangeEndAfterDeletion({
          before: document,
          after: deleted.value.document,
          from,
          to,
        });
      }
      if (text !== "")
        ops.push({
          type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          at,
          slice: (idKey(from.blockId) === idKey(to.blockId)
            ? continuedSlice(
                continuation(paragraph.content, from, to, {
                  ...(intent.inside === undefined
                    ? {}
                    : { inside: intent.inside }),
                }),
                content,
              )
            : undefined) ?? { openStart: 0, openEnd: 0, content },
          ...allocationFields,
        });
      selection = replacementCaret(at, text);
      break;
    }
    case "splitParagraph": {
      if (intent.to === undefined) {
        const split = {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: intent.at,
          newBlockId: intent.newBlockId,
          newHalf: intent.newHalf ?? SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(document, intent.at),
          ...tracked,
        } as const;
        if (mode.type === "suggesting") {
          const plan = createTrackedPlan({
            document,
            revision: mode.revision,
            newIds: mode.newIds,
          });
          const appended = plan.append(split);
          if (appended.isErr()) return Result.err(appended.error);
          ops = plan.ops;
        } else ops = [split];
        selection = { ...intent.at, offset: 0, zeroWidthBefore: 0 };
        break;
      }
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({
          document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        const deletion = appendTrackedDeletion({
          document,
          options: {
            from: intent.at,
            to: intent.to,
            revision: mode.revision,
            newIds: mode.newIds,
          },
          plan,
        });
        if (deletion.isErr()) return Result.err(deletion.error);
        const at = rangeEndAfterDeletion({
          before: document,
          after: plan.document(),
          from: intent.at,
          to: intent.to,
        });
        const split = plan.append({
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at,
          newBlockId: intent.newBlockId,
          newHalf: intent.newHalf ?? SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(plan.document(), at),
          newIds: mode.newIds,
          revision: mode.revision,
        });
        if (split.isErr()) return Result.err(split.error);
        ops = plan.ops;
        selection = { ...at, offset: 0, zeroWidthBefore: 0 };
        break;
      }
      const deletion = compileEditorIntent(document, {
        intent: {
          type: "replaceText",
          from: intent.at,
          to: intent.to,
          text: "",
        },
        mode,
      });
      if (deletion.isErr()) return deletion;
      const deleted = applyDocumentOps(document, deletion.value.ops);
      if (deleted.isErr()) return Result.err(deleted.error);
      ops = [
        ...deletion.value.ops,
        {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: deletion.value.selection,
          newBlockId: intent.newBlockId,
          newHalf: intent.newHalf ?? SPLIT_HALVES.FIRST,
          newParagraph: splitParagraphFields(
            deleted.value.document,
            deletion.value.selection,
          ),
          ...allocationFields,
        },
      ];
      selection = {
        ...deletion.value.selection,
        offset: 0,
        zeroWidthBefore: 0,
      };
      break;
    }
    case "joinParagraphs": {
      const at = { story: intent.story, blockId: intent.blockId, offset: 0 };
      const paragraph = paragraphAt(document, at);
      const ownInsertedMark =
        mode.type === "suggesting" &&
        (paragraph?.pPrMark?.kind === "ins" ||
          paragraph?.pPrMark?.kind === "moveTo") &&
        paragraph.pPrMark.info.author === mode.revision.author;
      const join = {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: intent.story,
        blockId: intent.blockId,
        nextBlockId: intent.nextBlockId,
        sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
        ...(mode.type === "editing" && intent.survivor !== undefined
          ? { survivor: intent.survivor }
          : {}),
        ...(ownInsertedMark ? {} : tracked),
      } as const;
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({
          document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        const joined = plan.append(join);
        if (joined.isErr()) return Result.err(joined.error);
        if (ownInsertedMark) {
          // Cancelling our inserted boundary is physical, but its formatting
          // effect remains a suggestion on the original trailing paragraph.
          const before = paragraphAt(document, {
            story: intent.story,
            blockId: intent.nextBlockId,
            offset: 0,
          });
          const after = paragraphAt(plan.document(), {
            story: intent.story,
            blockId: intent.nextBlockId,
            offset: 0,
          });
          if (!before || !after)
            panic("An own-mark join lost its trailing paragraph.");
          const restored = plan.append({
            type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
            story: intent.story,
            blockId: intent.nextBlockId,
            patch: Object.fromEntries([
              ...Object.keys(after.formatting ?? {}).map((key) => [key, null]),
              ...Object.entries(before.formatting ?? {}),
            ]),
            whenEmpty:
              before.formatting === undefined
                ? EMPTY_PROPERTY_SETS.OMIT
                : EMPTY_PROPERTY_SETS.KEEP,
          });
          if (restored.isErr()) return Result.err(restored.error);
        }
        const alignment = joinParagraphPropertyAlignment({
          document,
          current: plan.document(),
          intent,
          revision: mode.revision,
        });
        if (alignment !== undefined) {
          const patched = plan.append(alignment);
          if (patched.isErr()) return Result.err(patched.error);
        }
        ops = plan.ops;
      } else {
        ops = [join];
      }
      let offset = 0;
      if (
        (mode.type === "editing" || ownInsertedMark) &&
        paragraph !== undefined
      ) {
        offset = paragraphLength(paragraph);
      }
      // A direct join may keep the first paragraph; a tracked one keeps the second.
      const survivorId =
        mode.type === "editing" && intent.survivor === "first"
          ? intent.blockId
          : intent.nextBlockId;
      selection = { story: intent.story, blockId: survivorId, offset };
      break;
    }
    default: {
      const unreachable: never = intent;
      return unreachable;
    }
  }
  // Simulate direct pieces in order: each consumes its own pool, and later
  // pieces must not reuse identities created by an earlier piece.
  if (mode.type === "editing") {
    let current = document;
    const compact: DocumentOp[] = [];
    const steps: AppliedDocumentOp[] = [];
    const takenRevision = new Set<number>();
    const takenControl = new Set<number>();
    const sourceRunSeams: TextPosition[] = [];
    const replacesSourceContent =
      intent.type === "replaceText" &&
      ops.some((op) => op.type === DOCUMENT_OP_TYPES.DELETE_RANGE);
    const { story } = intentEndpoints(intent);
    for (const input of ops) {
      const op = filterNewIds(
        input,
        (space, id) =>
          !(
            space === IDENTITY_SPACES.REVISION ? takenRevision : takenControl
          ).has(id),
      );
      const applied = applyDocumentOp(current, op);
      if (applied.isErr()) return Result.err(applied.error);
      const sourceRunStart =
        replacesSourceContent && op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT
          ? replacementSourceRunStart(current, op)
          : undefined;
      const trimmed = trimAppliedNewIds({
        before: current,
        op,
        applied: applied.value,
        story,
      });
      if ("newIds" in trimmed) {
        for (const id of trimmed.newIds?.revision ?? []) takenRevision.add(id);
        for (const id of trimmed.newIds?.control ?? []) takenControl.add(id);
      }
      compact.push(trimmed);
      steps.push(applied.value);
      current = applied.value.document;
      if (sourceRunStart !== undefined) sourceRunSeams.push(sourceRunStart);
    }
    if (intent.type === "joinParagraphs") {
      const alignment = joinParagraphPropertyAlignment({
        document,
        current,
        intent,
      });
      if (alignment !== undefined) {
        const aligned = applyDocumentOp(current, alignment);
        if (aligned.isErr()) return Result.err(aligned.error);
        compact.push(alignment);
        steps.push(aligned.value);
        current = aligned.value.document;
      }
    }
    const insertedWidth =
      intent.type === "replaceText" ? intent.text.length : 0;
    if (insertedWidth > 0) {
      for (const offset of new Set([
        selection.offset,
        selection.offset - insertedWidth,
      ]))
        if (offset >= 0) editedSeams.push({ ...selection, offset });
    }
    // Merge only the edited seams, keeping every authored interior run boundary.
    const plannedSeams = new Map<
      string,
      { at: TextPosition; depth: number; seam: boolean }
    >();
    const seams = [
      ...editedSeams.map((at) => ({ at, policy: RUN_SEAM_POLICIES.plain })),
      ...sourceRunSeams.map((at) => ({
        at,
        policy: RUN_SEAM_POLICIES.sourceCut,
      })),
    ];
    for (const { at, policy } of seams) {
      const paragraph = paragraphAt(current, at);
      if (paragraph === undefined)
        panic("An edited paragraph must exist at its seam.");
      const depth = runSeamDepth({
        paragraph,
        at: { offset: at.offset, zeroWidthBefore: at.zeroWidthBefore ?? 0 },
        policy,
      });
      if (depth === 0) continue;
      const key = JSON.stringify([
        at.story,
        idKey(at.blockId),
        at.offset,
        at.zeroWidthBefore ?? 0,
      ]);
      // An edited seam, and a source cut that rejoins runs, join alike runs
      // where they still meet; a deeper source cut puts a record back.
      const seam =
        (policy === RUN_SEAM_POLICIES.plain || depth === 1) &&
        (plannedSeams.get(key)?.seam ?? true);
      plannedSeams.set(key, { at, depth, seam });
    }
    for (const { at, depth, seam } of [...plannedSeams.values()].toSorted(
      (left, right) =>
        right.at.offset - left.at.offset ||
        (right.at.zeroWidthBefore ?? 0) - (left.at.zeroWidthBefore ?? 0),
    )) {
      const join: JoinInlineOp = {
        type: DOCUMENT_OP_TYPES.JOIN_INLINE,
        at,
        depth,
        ...(seam ? { ifAlike: true } : {}),
      };
      const joined = applyDocumentOp(current, join);
      if (joined.isErr()) return Result.err(joined.error);
      compact.push(join);
      steps.push(joined.value);
      current = joined.value.document;
    }
    ops = compact;
    // Typed text hands its application back; other intents apply again.
    if (intent.type === "replaceText")
      handedBack = {
        ops: [...compact],
        result: {
          ...combineEdits(document, steps),
          revisions: steps.flatMap(({ revisions }) => revisions),
        },
      };
    if (paragraphAt(current, selection) === undefined)
      panic("A direct edit's caret must name a paragraph the edit keeps.");
  }
  const appliedAsCompiled =
    handedBack !== undefined &&
    handedBack.ops.length === ops.length &&
    handedBack.ops.every((op, index) => op === ops[index])
      ? handedBack.result
      : undefined;
  if (mode.type === "suggesting" && cutsBesideRemoval(document, ops)) {
    // A tracked edit that cuts a record whose first piece only holds content
    // on its way out records who takes the record's identities back.
    const cut =
      appliedAsCompiled === undefined
        ? applyDocumentOps(document, ops)
        : Result.ok(appliedAsCompiled);
    if (cut.isErr()) return Result.err(cut.error);
    ops = [
      ...ops,
      ...cutPieceHandoffs({
        before: document,
        after: cut.value.document,
        story: intentEndpoints(intent).story,
      }),
    ];
    return Result.ok({ ops: ops.map(captureDocumentOp), selection });
  }
  return Result.ok({
    ops: ops.map(captureDocumentOp),
    selection,
    ...(appliedAsCompiled === undefined ? {} : { applied: appliedAsCompiled }),
  });
};

/** A record holding removed content inside its own identity, in a paragraph. */
const holdsNestedRemoval = (nodes: readonly InlineNode[]): boolean =>
  nodes.some(
    (node) =>
      (identitySlots(node).length > 0 &&
        (childNodes(node) ?? []).some(isRemovedRevisionNode)) ||
      holdsNestedRemoval(childNodes(node) ?? []),
  );

/**
 * Whether a tracked edit may cut a record whose first piece then holds only
 * removed content: it removes a range, or inserts where such content sits.
 */
const cutsBesideRemoval = (
  document: Document,
  ops: readonly DocumentOp[],
): boolean =>
  ops.some(
    (op) =>
      (op.type === DOCUMENT_OP_TYPES.DELETE_RANGE &&
        op.revision !== undefined) ||
      ((op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT ||
        op.type === DOCUMENT_OP_TYPES.SPLIT_BLOCK) &&
        holdsNestedRemoval(paragraphAt(document, op.at)?.content ?? [])),
  );

type FreshParagraphIdsOptions = {
  document: Document;
  reserved: ReadonlySet<string> | undefined;
  lease: IdLease | undefined;
  count: number;
  opType: DocumentOp["type"];
};

/** The lowest paragraph identities neither the package nor the batch holds. */
const freshParagraphIds = ({
  document,
  reserved,
  lease,
  count,
  opType,
}: FreshParagraphIdsOptions): Result<string[], DocumentOpRefusal> => {
  const occupied = new Set(
    [...packageParagraphIds(document.package), ...(reserved ?? [])].map(idKey),
  );
  const ids = freshLeasedIds({
    lease,
    kind: ID_LEASE_KINDS.PARAGRAPH,
    count,
    taken: (id) => occupied.has(formatParaId(id)),
  });
  if (ids.length < count)
    return Result.err(
      new DocumentOpRefusal({
        opType,
        reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
        message:
          lease === undefined
            ? "The paragraph identity space is exhausted."
            : "The paragraph identity lease is exhausted.",
      }),
    );
  return Result.ok(ids.map(formatParaId));
};

const rowParagraphCount = (row: TableRow): number =>
  storyParagraphs({ content: [{ type: "table", rows: [row] }] }).length;

/** The row with each paragraph, in document order, given the next fresh id. */
const withFreshParagraphIds = (
  row: TableRow,
  ids: readonly string[],
): TableRow => {
  let next = 0;
  const blocks = (content: readonly BlockContent[]): BlockContent[] =>
    content.map((block) => {
      switch (block.type) {
        case "paragraph": {
          const paraId = ids[next] ?? panic("A fresh paragraph id is missing.");
          next += 1;
          const { textId: _textId, ...rest } = block;
          return { ...rest, paraId };
        }
        case "table":
          return {
            ...block,
            rows: block.rows.map((each) => ({
              ...each,
              cells: each.cells.map((cell) => ({
                ...cell,
                content: blocks(cell.content),
              })),
            })),
          };
        case "blockSdt":
        case "blockCustomXml":
          return { ...block, content: blocks(block.content) };
        default:
          return block;
      }
    });
  return {
    ...row,
    cells: row.cells.map((cell) => ({
      ...cell,
      content: blocks(cell.content),
    })),
  };
};

/** Containers typed text continues when it lands between their own leaves. */
const CONTINUED_CONTAINERS: ReadonlySet<string> = new Set([
  "inlineSdt",
  "inlineWrapper",
]);

type Continuation = {
  containers: readonly InlineNode[];
  openStart: number;
  openEnd: number;
};

/**
 * What text typed over the range from `from` to `to` continues: every
 * content control and wrapper that holds the leaves on both sides,
 * so typing within one extends it rather than cutting it in two. At an edge
 * only one side is inside, and the text stays outside, as a mark that does not
 * extend past its end, unless the caret stood just inside that edge.
 *
 * Text replacing a range takes the place after the range's last leaf, where
 * a tracked replacement follows the struck content: it continues what holds
 * that leaf and the leaf after the range. A tracked replacement meets the
 * struck leaf on its start side; a direct one meets the leaf before the range.
 */
const continuation = (
  items: readonly ParagraphContent[],
  from: Pick<TextPosition, "offset" | "zeroWidthBefore">,
  to: Pick<TextPosition, "offset" | "zeroWidthBefore">,
  {
    inside = {},
    struck = false,
  }: { inside?: { start?: boolean; end?: boolean }; struck?: boolean } = {},
): Continuation => {
  const spans = leafSpans(items);
  // A range starts where its deletion does: after the markers at its start
  // unless it names its gap. A caret names one gap for both of its ends.
  const caret =
    from.offset === to.offset && from.zeroWidthBefore === to.zeroWidthBefore;
  const start = {
    offset: from.offset,
    zeroWidthBefore:
      from.zeroWidthBefore ??
      (caret ? 0 : zeroWidthLeavesAt(items, from.offset).length),
  };
  const end = { offset: to.offset, zeroWidthBefore: to.zeroWidthBefore ?? 0 };
  // A leaf the range falls inside stands on both of its sides. A caret just
  // inside a container's edge stands with the leaf inside it.
  const outerLeft = spans.findLast(
    ({ before }) => compareGaps(before, start) < 0,
  );
  const outerRight = spans.find(({ after }) => compareGaps(after, end) > 0);
  const last = spans.findLast(
    ({ before, after }) =>
      compareGaps(before, end) < 0 && compareGaps(after, start) > 0,
  );
  const left =
    last ??
    (inside.start
      ? spans.find(({ after }) => compareGaps(after, start) > 0)
      : outerLeft);
  const right = inside.end
    ? spans.findLast(({ before }) => compareGaps(before, end) < 0)
    : outerRight;
  const containers: InlineNode[] = [];
  for (const [depth, ancestor] of (left?.ancestors ?? []).entries()) {
    if (
      right?.ancestors[depth] !== ancestor ||
      !CONTINUED_CONTAINERS.has(ancestor.type)
    )
      break;
    containers.push(ancestor);
  }
  // Each open end merges with the containers the source holds on that side;
  // at a container's own edge that side has none to meet.
  const continued = (side: (typeof spans)[number] | undefined) => {
    let depth = 0;
    while (
      depth < containers.length &&
      side?.ancestors[depth] === containers[depth]
    )
      depth += 1;
    return depth;
  };
  // A direct replacement goes in after the leaves its deletion keeps, so its
  // start meets the last of them, or the leaf before the range where none
  // stood in it.
  const kept =
    struck || caret
      ? undefined
      : leavesKeptByDirectDeletion({ spans, from: start, to: end }).at(-1);
  return {
    containers,
    openStart: continued(
      struck && last !== undefined ? last : (kept ?? outerLeft),
    ),
    openEnd: continued(outerRight),
  };
};

/** `content` inside copies of the continued containers; `undefined` when there are none. */
const continuedSlice = (
  { containers, openStart, openEnd }: Continuation,
  content: readonly ParagraphContent[],
): InlineSlice | undefined => {
  if (containers.length === 0) return undefined;
  let wrapped: ParagraphContent[] = content.slice();
  for (const container of containers.toReversed()) {
    switch (container.type) {
      case "inlineSdt":
      case "inlineWrapper":
        wrapped = [{ ...container, content: wrapped }];
        break;
      default:
        return undefined;
    }
  }
  return { openStart, openEnd, content: wrapped };
};
