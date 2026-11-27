import { panic, Result } from "better-result";

import { INSERTION_SEAM_POLICIES } from "../model/content";
import {
  type AbstractNumbering,
  type BlockContent,
  type BreakContent,
  type Comment,
  type Document,
  type Endnote,
  type Footnote,
  type InlineSdt,
  MAX_REVISION_ID,
  type NumberingInstance,
  type Paragraph,
  type ParagraphContent,
  type ParagraphFormatting,
  type ParagraphMarkChange,
  type RunContent,
  type SdtProperties,
  type TabContent,
  type Table,
  type TableRow,
  type TextFormatting,
  type TrackedRunContent,
} from "../model/document";
import { maxOf } from "../model/extremes";
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
  combineRecordedRevisions,
} from "./apply";
import {
  blockListAt,
  type ParagraphLocation,
  sameBlockList,
  storyBody,
  storyParagraphs,
} from "./blocks";
import {
  freshCommentId,
  isCommentExhausted,
  linksEmptiedByRetirement,
} from "./comments";
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
  takeParagraphIds,
} from "./idLease";
import {
  IDENTITY_SPACES,
  type IdentitySpace,
  idKey,
  packageIdentityKeys,
  packageReservedIdentityKeys,
  slotKey,
} from "./ids";
import { deleteBetween, gapAfterInserted } from "./inline";
import { isKeptCut, remapKeptCutHolders } from "./keptCut";
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
import { itemBefore } from "./listItems";
import { cloneModel } from "./modelClone";
import {
  noteReferenceKey,
  noteReferenceKind,
  noteReferenceSpans,
} from "./noteMarks";
import {
  childrenOf,
  contentWidth,
  isInlineContainer,
  isRangeBoundary,
  paragraphHoldsContentOnceAccepted,
  paragraphLength,
  withChildren,
} from "./offsets";
import { findStoryParagraph } from "./paragraphLocator";
import { authoredParagraphFormatting } from "./paragraphPropertySnapshot";
import {
  bookmarksIn,
  holdsOnlyPasteAnchors,
  type PasteAnnotations,
  type PlacedAnchor,
  placedPasteAnchors,
  withPasteAnchors,
} from "./pasteAnchors";
import {
  blockParagraphsOf,
  isPastedBlock,
  mapPastedBlockParagraphs,
  type PastedBlock,
} from "./pastedBlocks";
import {
  holdsMoveRange,
  withoutRepeatedMoveRangesIn,
} from "./pastedMoveRanges";
import { applyFormattingPatch } from "./patch";
import {
  joinedPropertiesSource,
  joinsFollowingParagraph,
  paragraphPropertiesSetting,
  paragraphVisibleText,
} from "./pendingJoins";
import {
  appendTrackedDeletion,
  createTrackedPlan,
  filterNewIds,
  leavesKeptByDirectDeletion,
  replacementDeletionSegments,
  selectedParagraphRuns,
  trimAppliedNewIds,
} from "./plan";
import { blockFollower } from "./rangeBoundaries";
import {
  cutPieceHandoffs,
  pastEndedComments,
  planTrackedReplace,
  planTrackedReplacement,
  rangeEndAfterDeletion,
} from "./rangeReplacement";
import { DOCUMENT_OP_REFUSAL_REASONS, DocumentOpRefusal } from "./refusal";
import {
  isRemovedRevisionNode,
  isTrackedWrapper,
  paragraphPropertiesOf,
  reviewFieldsOf,
} from "./review";
import { groupedDeletionStamp, ownInsertionSlice } from "./revisionGrouping";
import { formattingEquals, runsMergeable } from "./runMerge";
import { identitySlots } from "./slots";
import { findStoryBody, type StoryAddress, textBoxStoriesIn } from "./stories";
import {
  columnDeletionTakesTable,
  tableEditControlDemand,
  tableEditParagraphDemand,
} from "./tableEdits";
import { tableGrid } from "./tableGrid";
import {
  locateTableRow,
  tableRemovalAddress,
  tableRowAnchor,
  type TableRowLocation,
} from "./tableLocation";
import { rowsWithoutRow } from "./tableRowRemoval";
import {
  addressedTo,
  namingPackageStories,
  textBoxOfIntent,
} from "./textBoxRouting";
import { compileGenerateTOCIntent, type GenerateTOCIntent } from "./tocIntent";
import {
  type DeleteRangeOp,
  type DeleteRowOp,
  type DeleteTableOp,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  type DocumentOpType,
  EMPTY_PROPERTY_SETS,
  type InlineSlice,
  type InsertBlocksOp,
  type InsertContentOp,
  type InsertRowOp,
  type JoinInlineOp,
  type NewIds,
  OP_STORIES,
  type OpStory,
  type ParagraphPropsPatch,
  PROPERTY_REVIEW_POLICIES,
  type RevisionStamp,
  type RunPropsPatch,
  SECTION_BOUNDARY_POLICIES,
  type SetParagraphPropsOp,
  type SetParagraphReviewOp,
  SPLIT_HALVES,
  type SplitBlockOp,
  type SplitHalf,
  type TableEditOp,
  type TableInsertionPoint,
  type TableIntentOperation,
  type TextPosition,
} from "./types";
import { planWholeParagraphReplacement } from "./wholeParagraphReplacement";
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
export { paragraphVisibleText } from "./pendingJoins";
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
  | ({
      type: typeof DOCUMENT_OP_TYPES.INSERT_TABLE;
      story: OpStory;
      blockId: string;
      table: Table;
      /**
       * Where the table goes against the anchor: after it (the
       * default, as from a caret), or before it, as from a selection of the
       * whole story, which the editor fills at the story's start.
       */
      side?: "after" | "before";
    } & (
      | {
          /** Existing callers default to a paragraph anchor. */
          anchor?: "paragraph";
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
        }
      | {
          /** Beside the innermost table holding blockId, without added paragraphs. */
          anchor: "table";
          separated?: never;
          paragraphAfter?: never;
        }
      | {
          /**
           * Before the first block of the block list holding blockId, without
           * added paragraphs: where no paragraph or table names the place.
           */
          anchor: "containerStart";
          side?: never;
          separated?: never;
          paragraphAfter?: never;
        }
    ));

/**
 * Typed text as run content: a tab is a tab element and a line feed a line
 * break, each one offset wide as the character it stands for.
 */
const typedRunContent = (text: string): RunContent[] =>
  text.split(/(\t|\n)/u).flatMap((part): RunContent[] => {
    if (part === "\t") return [{ type: "tab" }];
    if (part === "\n") return [{ type: "break" }];
    return part === "" ? [] : [{ type: "text", text: part }];
  });

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
       * Whole paragraphs beside a paragraph or table: the compiler gives
       * each a fresh id, and a suggestion records each as inserted. Table-side
       * suggestions refuse boundaries without a reversible review plan.
       */
      type: "insertParagraphs";
      story: OpStory;
      at: TableInsertionPoint;
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
      /**
       * Tracked, what the deletion does with the paragraphs a pending join
       * runs into the first one (`paragraphsJoinedInto`), read as the
       * document stands when it compiles: `deleted`, they go with it, as the
       * joined paragraph they make carries its id.
       */
      joinedWords?: "deleted";
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
      /**
       * Clipboard blocks, tables among them, over a range: the range is
       * cleared, an open first or last paragraph joins the text at its edge,
       * and the blocks between go in before the caret's paragraph, split at
       * the caret when text stands before it. Only when every block is a
       * paragraph is it `replaceFragment`; a table, a container or kept
       * block markup among them is placed as a block.
       */
      type: "pasteBlocks";
      from: TextPosition;
      to: TextPosition;
      /**
       * A block content control or custom XML element goes in whole, its own
       * paragraphs identified as the others are, and so does block markup
       * kept as it was read; neither has a tracked record, so suggesting
       * refuses them.
       */
      blocks: readonly PastedBlock[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
      /**
       * The notes and comments the blocks' references and markers name, by
       * those ids: each is copied under a fresh id where its reference lands.
       */
      annotations?: PasteAnnotations;
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
      /**
       * Blocks, tables and containers among them, dragged from a range to a
       * target. The tables and containers standing wholly between `from`
       * and `to` are removed, a table in suggesting mode recorded as
       * deleted, the text of the range is deleted as `moveFragment` deletes
       * it, and the blocks go in at the target as `pasteBlocks` places them,
       * a suggestion recording them as inserted. A container has no tracked
       * record, so a suggestion moving one refuses. Only when every block
       * is a paragraph is it `moveFragment`. A direct move keeps the notes
       * and comments the blocks' references and markers name, under their
       * own ids.
       */
      type: "moveBlocks";
      from: TextPosition;
      to: TextPosition;
      target: TextPosition;
      blocks: readonly PastedBlock[];
      openStart: 0 | 1;
      openEnd: 0 | 1;
    }
  | {
      type: "replaceText";
      from: TextPosition;
      to: TextPosition;
      text: string;
      /** Where each piece of the inserted text takes its authored formatting. */
      formattingSources?: readonly {
        /** Exclusive UTF-16 offset in `text`, with contiguous pieces. */
        end: number;
        at: TextPosition;
        type: "character" | "caret";
      }[];
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

/**
 * Where text an intent adds takes its formatting from: the caret's, as typing
 * continues the run it stands in, or the body text's, as an operation's text
 * is body text and never continues a note reference.
 */
export const TYPED_TEXT_FORMATTING: Readonly<{
  readonly CARET: "caret";
  readonly BODY_TEXT: "bodyText";
}> = Object.freeze({ CARET: "caret", BODY_TEXT: "bodyText" } as const);

type TypedTextFormatting =
  (typeof TYPED_TEXT_FORMATTING)[keyof typeof TYPED_TEXT_FORMATTING];

type CompileEditorIntentOptions = {
  intent: EditorIntent;
  mode: EditorIntentMode;
  /** By default the caret's. */
  typedFormatting?: TypedTextFormatting;
  /** Source before this operation's other replacement pieces are compiled. */
  replacementFormattingDocument?: Document;
  /** Lowest candidate for compiler-owned pasted block IDs, including retired session IDs. */
  firstBlockId?: number;
  /**
   * Pasted blocks a direct move took from this story: their note references
   * and comment markers keep naming the notes and comments they named, which
   * the move leaves in place, rather than naming copies.
   */
  moved?: true;
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
    case "pasteBlocks":
    case "moveFragment":
    case "moveBlocks":
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
 * The ids deleting the notes `paragraphs` refer to takes, where a tracked
 * deletion takes their last live reference: each note paragraph's mark
 * and its leaves.
 */
const noteDeletionDemand = (
  document: Document,
  paragraphs: readonly Paragraph[],
): number => {
  const seen = new Set<string>();
  let demand = 0;
  for (const paragraph of paragraphs)
    for (const { node } of noteReferenceSpans(paragraph.content)) {
      const kind = noteReferenceKind(node);
      const key = noteReferenceKey(kind, node.id);
      if (seen.has(key)) continue;
      seen.add(key);
      const body = findStoryBody(document, { kind, id: node.id });
      if (body === undefined) continue;
      for (const { paragraph: noteParagraph } of storyParagraphs(body)) {
        demand += PARAGRAPH_REVISION_DEMAND;
        for (const span of leafSpans(noteParagraph.content))
          demand += leafRevisionDemand(span.ancestors.length);
      }
    }
  return demand;
};

type AllocationRange = {
  body: ReturnType<typeof storyBody>;
  fromKey: string;
  toKey: string;
};

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
  const paragraphsBetween = ({ body, fromKey, toKey }: AllocationRange) => {
    // An empty endpoint also matches a paragraph with no id in the story walk.
    if (fromKey !== "" && fromKey === toKey) {
      const location = findStoryParagraph(body, fromKey);
      return location === undefined ? [] : [location];
    }
    let locations = storyLocations.get(body);
    if (locations === undefined) {
      locations = storyParagraphs(body);
      storyLocations.set(body, locations);
    }
    const first = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === fromKey,
    );
    const last = locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === toKey,
    );
    return first < 0 || last < first ? [] : locations.slice(first, last + 1);
  };
  return (
    document: Document,
    requested: EditorIntentAllocation,
    lease?: IdLease,
  ): {
    revisionId: number;
    newBlockId: string;
    newIds: {
      revision: number[];
      control: number[];
    };
  } => {
    // Demand is counted where the intent is compiled: an intent naming a text
    // box's paragraphs through the story drawing it, in the box's own story.
    const routed = textBoxOfIntent(document, requested);
    const intent =
      routed === undefined
        ? requested
        : addressedTo(requested, routed.story, routed.box);
    const census = packageIdOccupancy(document.package);
    const { story, fromId, toId, fromOffset, toOffset } =
      intentEndpoints(intent);
    const body = storyBody(document, story);
    const fromKey = idKey(fromId);
    const toKey = idKey(toId);
    // Every selected leaf may start a deletion segment; ancestors may be cut
    // at both endpoints. Table and whole-paragraph insertion demand is
    // computed below from the table or inserted paragraphs.
    let demand = 1;
    let tableControlDemand = 0;
    const selected =
      intent.type === "table" || intent.type === "insertParagraphs"
        ? []
        : paragraphsBetween({ body, fromKey, toKey });
    for (const { paragraph } of selected) {
      demand += PARAGRAPH_REVISION_DEMAND;
      const blockId = idKey(paragraph.paraId ?? "");
      for (const span of leafSpans(paragraph.content)) {
        if (
          blockId === fromKey &&
          fromOffset !== undefined &&
          span.after.offset < fromOffset
        )
          continue;
        if (
          blockId === toKey &&
          toOffset !== undefined &&
          span.before.offset > toOffset
        )
          continue;
        demand += leafRevisionDemand(span.ancestors.length);
      }
    }
    demand += noteDeletionDemand(
      document,
      selected.map(({ paragraph }) => paragraph),
    );
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
    if (intent.type === "pasteBlocks" || intent.type === "moveBlocks") {
      // The split at the caret, each pasted paragraph's mark and content, and
      // each pasted table's rows; a moved table's rows are recorded twice,
      // deleted at the source and inserted at the target.
      demand += 8 + PARAGRAPH_REVISION_DEMAND;
      for (const block of intent.blocks) {
        if (intent.type === "moveBlocks" && block.type === "table")
          demand += 2 + block.rows.length;
        const paragraphs = blockParagraphs(block);
        if (block.type === "table") demand += 2 + block.rows.length;
        for (const paragraph of paragraphs) {
          demand += PARAGRAPH_REVISION_DEMAND;
          for (const span of leafSpans(paragraph.content))
            demand += leafRevisionDemand(span.ancestors.length);
        }
      }
      // Each copied note's reference and paragraphs; each copied comment's
      // markers may cut records in two.
      for (const anchor of withPasteAnchors(intent.blocks).anchors) {
        if (anchor.kind !== "note") {
          demand += 4;
          continue;
        }
        const note = (
          intent.type === "pasteBlocks" ? intent.annotations : undefined
        )?.notes.find(
          ({ type, id }) => type === anchor.note && id === anchor.id,
        );
        demand += 4;
        for (const block of note?.content ?? []) {
          if (block.type !== "paragraph") continue;
          demand += PARAGRAPH_REVISION_DEMAND;
          for (const span of leafSpans(block.content))
            demand += leafRevisionDemand(span.ancestors.length);
        }
      }
    }
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
    if (intent.type === "moveFragment" || intent.type === "moveBlocks") {
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
    const commentIds = new Set(
      (document.package.document.comments ?? []).map(({ id }) => id),
    );
    const fresh = (
      occupied: typeof census.revisions,
      count: number,
      kind: IdLeaseKind,
    ) =>
      freshLeasedIds({
        lease,
        kind,
        count,
        from:
          lease === undefined && count > 0
            ? occupied.occupiedPrefixEnd(ID_KIND_LIMITS[kind])
            : 1,
        // Drawn from a lease, a revision id also passes over the document's
        // comment ids: a comment the same batch made may hold a leased one.
        taken: (id) =>
          occupied.has(id) ||
          (lease !== undefined &&
            kind === ID_LEASE_KINDS.REVISION &&
            commentIds.has(id)),
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
        const paragraphs = storyParagraphs({
          content: [{ type: "table", rows: table.rows }],
        });
        for (const { paragraph } of paragraphs) {
          demand += PARAGRAPH_REVISION_DEMAND;
          for (const span of leafSpans(paragraph.content))
            demand += leafRevisionDemand(span.ancestors.length);
        }
        demand += noteDeletionDemand(
          document,
          paragraphs.map(({ paragraph }) => paragraph),
        );
      }
    }
    if (intent.type === "tableRow") {
      // A tracked row records itself, and each of its paragraphs its mark
      // and content: the new row's, or the deleted one's.
      const op = intent.operation;
      const row =
        op.type === DOCUMENT_OP_TYPES.INSERT_ROW
          ? op.row
          : locateTableRow(document, op)
              .map(({ table, rowIndex }) => table.rows.at(rowIndex))
              .unwrapOr(undefined);
      if (row !== undefined) {
        demand = 2 + PARAGRAPH_REVISION_DEMAND;
        const paragraphs = storyParagraphs({
          content: [{ type: "table", rows: [row] }],
        });
        for (const { paragraph } of paragraphs) {
          demand += PARAGRAPH_REVISION_DEMAND;
          for (const span of leafSpans(paragraph.content))
            demand += leafRevisionDemand(span.ancestors.length);
        }
        demand += noteDeletionDemand(
          document,
          paragraphs.map(({ paragraph }) => paragraph),
        );
      }
    }
    if (intent.type === "table") {
      const op = intent.operation;
      const located = locateTableRow(document, op);
      demand = 1;
      if (located.isOk()) {
        tableControlDemand = tableEditControlDemand(located.value, op);
        switch (op.type) {
          case DOCUMENT_OP_TYPES.INSERT_COLUMN:
            demand = located.value.table.rows.length * (op.count ?? 1) + 2;
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
            const grid = tableGrid(table, op.type);
            if (grid.isErr()) break;
            const count = op.count ?? 1;
            if (
              !Number.isSafeInteger(count) ||
              count < 1 ||
              count > grid.value.width
            )
              break;
            // Each cell or omission record owns at least one selected slot.
            demand = table.rows.reduce(
              (total, row) => total + Math.min(count, row.cells.length + 1),
              2,
            );
            if (!columnDeletionTakesTable(grid.value, op)) break;
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
            // The merged cell's last paragraph and the one before it, where
            // a decision removes the words it ends with.
            demand = (op.bottom - op.top + 1) * 2 + 2;
            if (op.bottom - op.top === 1)
              demand = Math.max(demand, op.right - op.left + 2);
            // A row the merge leaves only continuations records its
            // deletion: the row and each of its paragraphs. A direct merge
            // across rows pending insertion or deletion records it on the
            // words each such row gives the merged cell, its first one's too.
            for (const row of located.value.table.rows.slice(
              op.top,
              op.bottom,
            )) {
              demand += 1;
              for (const { paragraph } of storyParagraphs({
                content: [{ type: "table", rows: [row] }],
              })) {
                demand += PARAGRAPH_REVISION_DEMAND;
                for (const span of leafSpans(paragraph.content))
                  demand += leafRevisionDemand(span.ancestors.length);
              }
            }
            break;
          case DOCUMENT_OP_TYPES.SPLIT_CELL:
            // A cell for each column of the span, and a merge change, in
            // each row a vertical merge reaches, which is at most every row.
            demand =
              ((located.value.table.rows
                .at(located.value.rowIndex)
                ?.cells.at(located.value.cellIndex)?.formatting?.gridSpan ??
                1) +
                1) *
              located.value.table.rows.length;
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
      from:
        lease === undefined
          ? census.paragraphs.occupiedPrefixEnd(ID_KIND_LIMITS.paragraph)
          : 1,
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
          intent.type === "table" ? tableControlDemand : demand,
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

type EditorParagraphGroup = {
  blockId: string;
  paragraphs: Paragraph[];
  text: string;
};

const paragraphGroupsIn = (
  document: Document,
  story: StoryAddress,
): EditorParagraphGroup[] => {
  const groups: EditorParagraphGroup[] = [];
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
    if (joinsFollowingParagraph(paragraph)) continue;
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

/** Deleted paragraph marks concatenate physical paragraphs in the editing view. */
export const editorParagraphGroups = (
  document: Document,
  story: OpStory,
): {
  blockId: string;
  paragraphs: Paragraph[];
  text: string;
}[] => paragraphGroupsIn(document, story);

/**
 * {@link editorParagraphGroups} of `story` and of every text box drawn in it,
 * each box grouped within its own story: a box's paragraphs are named
 * through the story drawing it, and a mark pending deletion in a box joins
 * the box's next paragraph.
 */
export const editorParagraphGroupsThroughTextBoxes = (
  document: Document,
  story: OpStory,
): ReturnType<typeof editorParagraphGroups> => [
  ...paragraphGroupsIn(document, story),
  ...textBoxStoriesIn(document, story).flatMap((box) =>
    paragraphGroupsIn(document, box),
  ),
];

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

/**
 * Why a join's paragraphs cannot be joined, as the join itself refuses it:
 * one is missing, or the second does not directly follow the first.
 */
const adjacentJoinRefusal = (
  document: Document,
  join: { story: OpStory; blockId: string; nextBlockId: string },
): DocumentOpRefusal | undefined => {
  const paragraphs = storyParagraphs(storyBody(document, join.story));
  const find = (blockId: string) =>
    paragraphs.find(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(blockId),
    );
  const first = find(join.blockId);
  const next = find(join.nextBlockId);
  if (first === undefined || next === undefined)
    return new DocumentOpRefusal({
      reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      message: `${first === undefined ? join.blockId : join.nextBlockId} does not exist.`,
      opType: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    });
  if (!sameBlockList(first.list, next.list) || next.index !== first.index + 1)
    return new DocumentOpRefusal({
      reason: DOCUMENT_OP_REFUSAL_REASONS.NOT_ADJACENT,
      message: `${join.nextBlockId} does not directly follow ${join.blockId}.`,
      opType: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    });
  return undefined;
};

const paragraphAt = (document: Document, at: TextPosition) =>
  findStoryParagraph(storyBody(document, at.story), at.blockId)?.paragraph;

type ContentWithinRange = { paragraph: Paragraph; from: number; to: number };
/**
 * Whether `paragraph` holds content in `[from, to)` once its changes are
 * accepted, by the one paragraph-content rule: range markers, comment
 * references and what accepting takes away hold none. A paste keeps a
 * paragraph for the text beside it only when this holds, in either mode, so
 * accepting a suggested paste leaves the paragraphs the direct paste does.
 */
const holdsContentWithin = ({
  paragraph,
  from,
  to,
}: ContentWithinRange): boolean => {
  // The items in [from, to), a container clipped to its children there, so
  // content of a control outside the interval does not count.
  const clip = (
    items: readonly ParagraphContent[],
    start: number,
  ): ParagraphContent[] => {
    const content: ParagraphContent[] = [];
    let offset = start;
    for (const item of items) {
      const end = offset + contentWidth([item]);
      if (end > from && offset < to)
        content.push(
          isInlineContainer(item)
            ? withChildren(item, clip(childrenOf(item), offset))
            : item,
        );
      offset = end;
    }
    return content;
  };
  return paragraphHoldsContentOnceAccepted({
    ...paragraph,
    content: clip(paragraph.content, 0),
  });
};

/** Bookmark ids the content's starts and ends carry, containers included. */
const bookmarkIdsIn = (items: readonly ParagraphContent[]): number[] =>
  items.flatMap((item) => {
    if (item.type === "bookmarkStart" || item.type === "bookmarkEnd")
      return [item.id];
    return isInlineContainer(item) ? bookmarkIdsIn(childrenOf(item)) : [];
  });

/** Whether a pasted table, or one nested in it, holds bookmarks or records of its own. */
const holdsTableLevelIdentities = (block: BlockContent): boolean => {
  if (block.type === "blockSdt" || block.type === "blockCustomXml")
    return block.content.some(holdsTableLevelIdentities);
  if (block.type !== "table") return false;
  if (
    (block.bookmarks?.length ?? 0) > 0 ||
    (block.propertyChanges?.length ?? 0) > 0 ||
    (block.carrierStack?.length ?? 0) > 0
  )
    return true;
  return block.rows.some(
    (row) =>
      (row.bookmarks?.length ?? 0) > 0 ||
      (row.propertyChanges?.length ?? 0) > 0 ||
      (row.carrierStack?.length ?? 0) > 0 ||
      row.cells.some(
        (cell) =>
          (cell.propertyChanges?.length ?? 0) > 0 ||
          (cell.carrierStack?.length ?? 0) > 0 ||
          cell.content.some(holdsTableLevelIdentities),
      ),
  );
};

/**
 * A pasted block's paragraphs in document order, a table's cells' and a
 * container's among them.
 */
const blockParagraphs = (block: PastedBlock): Paragraph[] =>
  blockParagraphsOf(block);

/** The properties of every block content control `blocks` hold, at any depth. */
const blockControlsOf = (blocks: readonly BlockContent[]): SdtProperties[] =>
  blocks.flatMap((block) => {
    switch (block.type) {
      case "blockSdt":
        return [block.properties, ...blockControlsOf(block.content)];
      case "blockCustomXml":
        return blockControlsOf(block.content);
      case "table":
        return block.rows.flatMap((row) =>
          row.cells.flatMap((cell) => blockControlsOf(cell.content)),
        );
      default:
        return [];
    }
  });

/** `blocks` with their block content controls' properties, in order, from `controls`. */
const withBlockControls = (
  blocks: readonly BlockContent[],
  controls: SdtProperties[],
): BlockContent[] =>
  blocks.map((block) => {
    switch (block.type) {
      case "blockSdt": {
        const properties =
          controls.shift() ?? panic("An identified control is missing.");
        return {
          ...block,
          properties,
          content: withBlockControls(block.content, controls),
        };
      }
      case "blockCustomXml":
        return {
          ...block,
          content: withBlockControls(block.content, controls),
        };
      case "table":
        return {
          ...block,
          rows: block.rows.map((row) => ({
            ...row,
            cells: row.cells.map((cell) => ({
              ...cell,
              content: withBlockControls(cell.content, controls),
            })),
          })),
        };
      default:
        return block;
    }
  });

type BlockPasteIntent = Extract<EditorIntent, { type: "pasteBlocks" }>;
type BlockPasteOptions = {
  intent: BlockPasteIntent;
  mode: EditorIntentMode;
  firstBlockId?: number;
  moved?: true | undefined;
};
type BlockPlacementOptions = {
  intent: Omit<BlockPasteIntent, "annotations">;
  mode: EditorIntentMode;
  firstBlockId?: number | undefined;
};
type AnchoredPlacement = {
  ops: DocumentOp[];
  document: Document;
  caret: TextPosition;
  places: (index: number) => PlacedAnchor | undefined;
};
type AnchoredPlacementOptions = BlockPlacementOptions & {
  /** The blocks without their placeholders. */
  plain: readonly PastedBlock[];
  /** The name prefix, bookmark ids and number of this paste's placeholders. */
  prefix: string;
  ids: ReadonlySet<number>;
  count: number;
};
/**
 * Clipboard blocks placed: once with their placeholders, to find where each
 * lands, and once without, for the operations and the document they leave.
 * Placeholders are zero width and hold no content, so both place the blocks
 * alike; a found place the plain document does not hold refuses. An open
 * fragment holding only placeholders places nothing: the range is cleared and
 * every placeholder lands at the caret. A closed paragraph is placed as one.
 */
const placeAnchoredBlocks = (
  document: Document,
  {
    intent,
    plain,
    prefix,
    ids,
    count,
    mode,
    firstBlockId,
  }: AnchoredPlacementOptions,
): Result<AnchoredPlacement, DocumentOpRefusal> => {
  const story = intent.from.story;
  const paragraphsOf = (from: Document) =>
    storyParagraphs(storyBody(from, story)).map(({ paragraph }) => paragraph);
  const only = intent.blocks.length === 1 ? intent.blocks[0] : undefined;
  if (
    only?.type === "paragraph" &&
    intent.openStart === 1 &&
    intent.openEnd === 1 &&
    holdsOnlyPasteAnchors(only, ids)
  ) {
    // A caret names one gap: two endpoints at one offset may still hold the
    // markers between their gaps, which the clear must take.
    const collapsed =
      idKey(intent.from.blockId) === idKey(intent.to.blockId) &&
      intent.from.offset === intent.to.offset &&
      (intent.from.zeroWidthBefore ?? 0) === (intent.to.zeroWidthBefore ?? 0);
    const cleared = collapsed
      ? Result.ok({ ops: [], selection: intent.from })
      : compileEditorIntent(document, {
          intent: {
            type: "replaceText",
            from: intent.from,
            to: intent.to,
            text: "",
          },
          mode,
          ...(firstBlockId === undefined ? {} : { firstBlockId }),
        });
    if (cleared.isErr()) return Result.err(cleared.error);
    const applied = applyDocumentOps(document, cleared.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    const caret = cleared.value.selection;
    const order = paragraphsOf(applied.value.document).findIndex(
      ({ paraId }) => paraId === caret.blockId,
    );
    const at = { blockId: caret.blockId, offset: caret.offset, order };
    return Result.ok({
      ops: [...cleared.value.ops],
      document: applied.value.document,
      caret,
      places: () => at,
    });
  }
  const placed = compileBlockPlacement(document, {
    intent,
    mode,
    firstBlockId,
  });
  if (placed.isErr()) return Result.err(placed.error);
  const marked =
    placed.value.applied === undefined
      ? applyDocumentOps(document, placed.value.ops)
      : Result.ok(placed.value.applied);
  if (marked.isErr()) return Result.err(marked.error);
  const lost = () =>
    Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
        message: "A pasted note or comment lost its place in the paste.",
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      }),
    );
  // Every placeholder of this paste lands once; one missing or found twice
  // leaves no place that is its own.
  const found = placedPasteAnchors({
    paragraphs: paragraphsOf(marked.value.document),
    prefix,
    count,
  });
  if (found === undefined) return lost();
  const unmarkedPlacement = compileBlockPlacement(document, {
    intent: { ...intent, blocks: plain },
    mode,
    firstBlockId,
  });
  if (unmarkedPlacement.isErr()) return Result.err(unmarkedPlacement.error);
  const unmarked =
    unmarkedPlacement.value.applied === undefined
      ? applyDocumentOps(document, unmarkedPlacement.value.ops)
      : Result.ok(unmarkedPlacement.value.applied);
  if (unmarked.isErr()) return Result.err(unmarked.error);
  const lengths = new Map(
    paragraphsOf(unmarked.value.document).map((paragraph) => [
      paragraph.paraId,
      paragraphLength(paragraph),
    ]),
  );
  for (const { blockId, offset } of found.values()) {
    const length = lengths.get(blockId);
    if (length === undefined || offset > length) return lost();
  }
  return Result.ok({
    ops: [...unmarkedPlacement.value.ops],
    document: unmarked.value.document,
    caret: unmarkedPlacement.value.selection,
    places: (index) => found.get(index),
  });
};

type CopiedAnnotation =
  | { kind: "note"; at: PlacedAnchor; note: Footnote | Endnote }
  | {
      kind: "comment";
      from: PlacedAnchor | undefined;
      at: PlacedAnchor;
      comment: Comment;
    };

/**
 * Clipboard blocks whose note references and comment markers name notes and
 * comments: the blocks are placed with a placeholder for each, then each note
 * and comment is copied under a fresh id where its placeholder landed, the
 * latest place first so that no copy moves the place of another. A comment
 * whose range the clipboard holds only part of is copied as a point.
 */
const compileBlockPaste = (
  document: Document,
  { intent, mode, firstBlockId, moved }: BlockPasteOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const { annotations, ...placement } = intent;
  const anchored = withPasteAnchors(
    intent.blocks,
    bookmarksIn(document.package),
  );
  // Moved directly, the references and markers go with their blocks and keep
  // the notes and comments the source left in place: there is nothing to copy.
  if (anchored.anchors.length === 0 || moved === true)
    return compileBlockPlacement(document, {
      intent: placement,
      mode,
      firstBlockId,
    });
  const refuse = (message: string) =>
    Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
        message,
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      }),
    );
  const story = intent.from.story;
  const notes: (Footnote | Endnote | undefined)[] = [];
  for (const anchor of anchored.anchors) {
    if (anchor.kind !== "note") {
      if (!annotations?.comments.some(({ id }) => id === anchor.id))
        return refuse("A pasted comment marker names no copied comment.");
      notes.push(undefined);
      continue;
    }
    const note = annotations?.notes.find(
      ({ type, id }) => type === anchor.note && id === anchor.id,
    );
    if (note === undefined)
      return refuse("A pasted note reference names no copied note.");
    if (story !== OP_STORIES.MAIN)
      return refuse("Notes are pasted only into the document body.");
    if (note.content.some((block) => block.type !== "paragraph"))
      return refuse("A pasted note holds only paragraphs.");
    notes.push(note);
  }
  const placed = placeAnchoredBlocks(document, {
    intent: { ...placement, blocks: anchored.blocks },
    plain: anchored.plain,
    prefix: anchored.prefix,
    ids: anchored.ids,
    count: anchored.anchors.length,
    mode,
    firstBlockId,
  });
  if (placed.isErr()) return placed;
  const { ops, places, caret } = placed.value;
  let current = placed.value.document;

  const copies: (CopiedAnnotation & { rank: number })[] = [];
  type Ranked = { at: PlacedAnchor; rank: number };
  const ranges = new Map<
    number,
    Partial<Record<"start" | "end" | "reference", Ranked>>
  >();
  for (const [index, anchor] of anchored.anchors.entries()) {
    const at = places(index);
    if (at === undefined) continue;
    const note = notes[index];
    if (anchor.kind === "note") {
      if (note !== undefined)
        copies.push({ kind: "note", at, note, rank: index });
      continue;
    }
    const range = ranges.get(anchor.id) ?? {};
    const ranked = { at, rank: index };
    if (anchor.kind === "commentRangeStart") range.start = ranked;
    else if (anchor.kind === "commentRangeEnd") range.end = ranked;
    else range.reference = ranked;
    ranges.set(anchor.id, range);
  }
  for (const [id, { start, end, reference }] of ranges) {
    const comment =
      annotations?.comments.find((each) => each.id === id) ??
      panic("A pasted comment was checked against the copied comments.");
    const ranged = start !== undefined && end !== undefined;
    const last = ranged ? end : (reference ?? end ?? start);
    if (last === undefined) continue;
    copies.push({
      kind: "comment",
      from: ranged ? start.at : undefined,
      at: last.at,
      comment,
      rank: last.rank,
    });
  }
  // The latest place first; copies at one place in reverse, so that each
  // lands before the ones copied after it.
  copies.sort(
    (left, right) =>
      right.at.order - left.at.order ||
      right.at.offset - left.at.offset ||
      right.rank - left.rank,
  );

  const position = ({ blockId, offset }: PlacedAnchor): TextPosition => ({
    story,
    blockId,
    offset,
  });
  const nextId = (ids: readonly { id: number }[] | undefined) =>
    maxOf(
      (ids ?? []).map(({ id }) => id),
      0,
    ) + 1;
  type FreshContentOptions = {
    content: readonly Paragraph[];
    opType: DocumentOpType;
    /** Revision ids the copy itself takes, which the body leaves free. */
    reserved: readonly number[];
  };
  // A copied body takes what pasted paragraphs take: fresh content-control,
  // revision and bookmark identities, and a fresh id for every paragraph.
  const freshContent = ({
    content,
    opType,
    reserved,
  }: FreshContentOptions): Result<Paragraph[], DocumentOpRefusal> => {
    const identified = identifyClipboardParagraphs({
      document: current,
      paragraphs: content,
      mode: {
        ...mode,
        newIds: {
          ...mode.newIds,
          revision: [...(mode.newIds?.revision ?? []), ...reserved],
        },
      },
      firstBlockId: undefined,
    });
    if (identified.isErr()) return Result.err(identified.error);
    const fresh = freshParagraphIds({
      document: current,
      reserved: mode.reservedBlockIds,
      lease: mode.idLease,
      count: content.length,
      opType,
    });
    if (fresh.isErr()) return Result.err(fresh.error);
    // The identified paragraphs are this body's own copies.
    for (const [index, paragraph] of identified.value.entries())
      paragraph.paraId =
        fresh.value[index] ?? panic("A fresh paragraph id is missing.");
    return Result.ok(identified.value);
  };
  // The revision ids the mode gives that the package leaves free.
  // A comment record holds its id without occupying a revision, so the ids
  // comments hold, a copy just made among them, are left out too.
  const freeRevisionIds = () => {
    const occupied = packageIdOccupancy(current.package).revisions;
    const commented = new Set(
      (current.package.document.comments ?? []).map(({ id }) => id),
    );
    return [
      ...(mode.type === "suggesting" ? [mode.revision.id] : []),
      ...(mode.newIds?.revision ?? []),
    ].filter(
      (candidate) => !occupied.has(candidate) && !commented.has(candidate),
    );
  };
  const out = [...ops];
  for (const copy of copies) {
    let op: DocumentOp;
    if (copy.kind === "note") {
      const content = freshContent({
        content: copy.note.content.filter(
          (block): block is Paragraph => block.type === "paragraph",
        ),
        opType: DOCUMENT_OP_TYPES.ADD_NOTE,
        reserved: [],
      });
      if (content.isErr()) return Result.err(content.error);
      const note =
        copy.note.type === "footnote"
          ? {
              ...cloneModel(copy.note),
              id: nextId(current.package.footnotes),
              content: content.value,
            }
          : {
              ...cloneModel(copy.note),
              id: nextId(current.package.endnotes),
              content: content.value,
            };
      let stamp = {};
      if (mode.type === "suggesting") {
        const [id, ...rest] = freeRevisionIds();
        if (id === undefined)
          return Result.err(
            new DocumentOpRefusal({
              reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
              message: "The pasted notes need more revision ids.",
              opType: DOCUMENT_OP_TYPES.ADD_NOTE,
            }),
          );
        stamp = {
          revision: { ...mode.revision, id },
          newIds: { revision: rest, control: mode.newIds.control ?? [] },
        };
      }
      op = {
        type: DOCUMENT_OP_TYPES.ADD_NOTE,
        at: position(copy.at),
        note,
        ...stamp,
      };
    } else {
      // A copy takes its id as an added comment does: from the lease's
      // comment ids, or past every id the package's annotations carry.
      const fresh = freshCommentId(current, mode.idLease);
      if (fresh.isErr()) return Result.err(fresh.error);
      const commentId = fresh.value;
      const content = freshContent({
        content: copy.comment.content,
        opType: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        reserved: [commentId],
      });
      if (content.isErr()) return Result.err(content.error);
      const {
        parentId: _parent,
        suggestionId: _suggestion,
        ...source
      } = cloneModel(copy.comment);
      // What proposed the comment elsewhere does not propose it here; a
      // suggestion pasting it does, so rejecting that suggestion takes it.
      const proposed =
        mode.type === "suggesting" ? mode.revision.suggestionId : undefined;
      // The comment's markers may cut pasted records, which take new ids.
      op = {
        type: DOCUMENT_OP_TYPES.CREATE_COMMENT,
        newIds: {
          revision: freeRevisionIds().filter((id) => id !== commentId),
          control: mode.newIds?.control ?? [],
        },
        comment: {
          ...source,
          id: commentId,
          content: content.value,
          anchorKind: copy.from === undefined ? "point" : "range",
          ...(proposed === undefined ? {} : { suggestionId: proposed }),
        },
        anchor:
          copy.from === undefined
            ? { kind: "point", at: position(copy.at) }
            : {
                kind: "range",
                from: position(copy.from),
                to: position(copy.at),
              },
      };
    }
    const applied = applyDocumentOp(current, op);
    if (applied.isErr()) return Result.err(applied.error);
    out.push(
      trimAppliedNewIds({ before: current, op, applied: applied.value, story }),
    );
    current = applied.value.document;
  }
  // The caret follows each reference copied at or before it in its paragraph.
  const passed = copies.filter(
    ({ at }) => at.blockId === caret.blockId && at.offset <= caret.offset,
  ).length;
  return Result.ok({
    ops: out,
    selection: { ...caret, offset: caret.offset + passed },
  });
};

/**
 * Clipboard blocks with tables: one step after another, each compiled against
 * the document the last left and, suggesting, taking the revision ids the
 * steps before it left unused.
 */
type PlaceContainerBlockOptions = {
  document: Document;
  story: OpStory;
  /** The paragraph the block goes in before. */
  anchor: string;
  block: Extract<
    PastedBlock,
    { type: "blockSdt" | "blockCustomXml" | "preservedBlock" }
  >;
  mode: EditorIntentMode;
};

/**
 * A suggestion's refusal of a container block it would paste or move. A
 * block-level `w:sdt` takes no revision of its own (CT_SdtBlock records none:
 * an insertion is recorded only on the runs and paragraph marks inside it, so
 * rejecting would leave the empty control standing), which is the format's
 * rule. A block-level `w:customXml` has range records of its own
 * (`w:customXmlInsRangeStart` and its kin) that the model does not carry yet,
 * and kept markup is not read.
 */
const containerBlockSuggestionRefusal = (
  block: { type: BlockContent["type"] },
  message: string,
): DocumentOpRefusal =>
  new DocumentOpRefusal({
    reason:
      block.type === "blockSdt"
        ? DOCUMENT_OP_REFUSAL_REASONS.NO_REVISION_RECORD
        : DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
    message,
    opType: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
  });

/**
 * A block content control, custom XML element or kept block markup put in
 * before `anchor`, in the block list holding it, as its exact list. None has
 * a tracked record the model carries, so a suggestion refuses
 * ({@link containerBlockSuggestionRefusal}).
 */
const placeContainerBlock = ({
  document,
  story,
  anchor,
  block,
  mode,
}: PlaceContainerBlockOptions): Result<
  { op: DocumentOp; document: Document },
  DocumentOpRefusal
> => {
  if (mode.type === "suggesting")
    return Result.err(
      containerBlockSuggestionRefusal(
        block,
        "A suggested paste cannot record a block content control, custom XML element or kept block markup.",
      ),
    );
  const body = storyBody(document, story);
  const located = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(anchor),
  );
  if (located === undefined)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        message: "The clipboard destination paragraph does not exist.",
        opType: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
      }),
    );
  const expected = blockListAt(body.content, located.list);
  const op: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
    story,
    blockId: anchor,
    expected,
    blocks: [
      ...expected.slice(0, located.index),
      block,
      ...expected.slice(located.index),
    ],
  };
  const applied = applyDocumentOps(document, [op]);
  if (applied.isErr()) return Result.err(applied.error);
  return Result.ok({ op, document: applied.value.document });
};

/**
 * A block content control, custom XML element or kept block markup taken out
 * of the block list holding `anchor`, as its exact list; a suggestion has no
 * record for it and refuses.
 */
const removeContainerBlock = ({
  document,
  story,
  anchor,
  block,
  mode,
}: {
  document: Document;
  story: OpStory;
  anchor: string;
  block: BlockContent;
  mode: EditorIntentMode;
}): Result<
  { op: DocumentOp; document: Document; inverse: readonly DocumentOp[] },
  DocumentOpRefusal
> => {
  const refuse = (
    reason: DocumentOpRefusal["reason"],
    message: string,
  ): Result<never, DocumentOpRefusal> =>
    Result.err(
      new DocumentOpRefusal({
        reason,
        message,
        opType: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
      }),
    );
  if (mode.type === "suggesting")
    return Result.err(
      containerBlockSuggestionRefusal(
        block,
        "A suggested move cannot record a block content control, custom XML element or kept block markup.",
      ),
    );
  const body = storyBody(document, story);
  const located = storyParagraphs(body).find(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(anchor),
  );
  if (located === undefined)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
      "The moved range names a paragraph the story does not hold.",
    );
  const expected = blockListAt(body.content, located.list);
  const index = expected.indexOf(block);
  if (index < 0)
    return refuse(
      DOCUMENT_OP_REFUSAL_REASONS.STALE,
      "The moved block is no longer where the range found it.",
    );
  const op: DocumentOp = {
    type: DOCUMENT_OP_TYPES.SET_CONTAINER_BLOCKS,
    story,
    blockId: anchor,
    expected,
    blocks: [...expected.slice(0, index), ...expected.slice(index + 1)],
  };
  const applied = applyDocumentOps(document, [op]);
  if (applied.isErr()) return Result.err(applied.error);
  return Result.ok({
    op,
    document: applied.value.document,
    inverse: applied.value.inverse,
  });
};

const compileBlockPlacement = (
  document: Document,
  { intent, mode, firstBlockId }: BlockPlacementOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const { from, to, openStart, openEnd } = intent;
  const paragraphs = intent.blocks.filter(
    (block): block is Paragraph => block.type === "paragraph",
  );
  if (paragraphs.length === intent.blocks.length)
    return compileEditorIntent(document, {
      intent: {
        type: "replaceFragment",
        from,
        to,
        paragraphs,
        openStart,
        openEnd,
      },
      mode,
      ...(firstBlockId === undefined ? {} : { firstBlockId }),
    });
  // Every pasted paragraph, a table's and a container's among them, takes
  // the identities a clipboard paragraph takes: its copied bookmarks and
  // records renamed. A block content control takes a fresh control id.
  const lastBlock = intent.blocks.at(-1);
  const identified = identifyClipboardContent({
    document,
    paragraphs: intent.blocks.flatMap(blockParagraphs),
    controls: blockControlsOf(intent.blocks),
    // A last paragraph or table keeps its last paragraph for the join or
    // the insertion to identify; inside a container, every paragraph needs
    // its id here.
    identifyLast:
      lastBlock?.type !== "paragraph" && lastBlock?.type !== "table",
    mode,
    firstBlockId,
  });
  if (identified.isErr()) return Result.err(identified.error);
  let next = 0;
  const take = (): Paragraph =>
    identified.value.paragraphs[next++] ??
    panic("An identified clipboard paragraph is missing.");
  const controls = [...identified.value.controls];
  const blocks = intent.blocks.map((block): PastedBlock => {
    const placed = mapPastedBlockParagraphs(block, take);
    if (placed.type === "paragraph" || placed.type === "preservedBlock")
      return placed;
    const [withControls] = withBlockControls([placed], controls);
    return withControls !== undefined && isPastedBlock(withControls)
      ? withControls
      : panic("A pasted block lost its kind.");
  });
  const refuseIdentities = (message: string) =>
    Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
        message,
        opType: DOCUMENT_OP_TYPES.INSERT_TABLE,
      }),
    );
  // A table's own bookmarks and records take no fresh identity here; the
  // paste refuses rather than copy an id the destination may hold.
  if (blocks.some(holdsTableLevelIdentities))
    return refuseIdentities(
      "A pasted table's own bookmarks and tracked records need fresh identities.",
    );
  // The open first and last paragraphs are identified again as they join; a
  // bookmark from one of them to another block would lose its pair.
  const joinedBlocks = new Set<number>();
  if (openStart === 1 && blocks.at(0)?.type === "paragraph")
    joinedBlocks.add(0);
  if (openEnd === 1 && blocks.at(-1)?.type === "paragraph")
    joinedBlocks.add(blocks.length - 1);
  const bookmarkBlocks = new Map<number, Set<number>>();
  for (const [index, block] of blocks.entries())
    for (const paragraph of blockParagraphs(block))
      for (const id of bookmarkIdsIn(paragraph.content)) {
        const holders = bookmarkBlocks.get(id) ?? new Set<number>();
        holders.add(index);
        bookmarkBlocks.set(id, holders);
      }
  for (const holders of bookmarkBlocks.values())
    if (
      holders.size > 1 &&
      [...holders].some((index) => joinedBlocks.has(index))
    )
      return refuseIdentities(
        "A pasted bookmark spanning a joined paragraph and a table needs one identity.",
      );
  const ops: DocumentOp[] = [];
  let current = document;
  let pool =
    mode.type === "suggesting"
      ? [mode.revision.id, ...(mode.newIds.revision ?? [])]
      : [];
  const step = (
    stepIntent: EditorIntent,
  ): Result<TextPosition, DocumentOpRefusal> => {
    let stepMode: EditorIntentMode = mode;
    if (mode.type === "suggesting") {
      const [id, ...rest] = pool;
      if (id === undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.NEEDS_NEW_IDS,
            message: "The pasted blocks need more revision ids.",
            opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
          }),
        );
      stepMode = {
        ...mode,
        revision: { ...mode.revision, id },
        newIds: { ...mode.newIds, revision: rest },
      };
    }
    const compiled = compileEditorIntent(current, {
      intent: stepIntent,
      mode: stepMode,
      ...(firstBlockId === undefined ? {} : { firstBlockId }),
    });
    if (compiled.isErr()) return Result.err(compiled.error);
    const applied =
      compiled.value.applied === undefined
        ? applyDocumentOps(current, compiled.value.ops)
        : Result.ok(compiled.value.applied);
    if (applied.isErr()) return Result.err(applied.error);
    ops.push(...compiled.value.ops);
    current = applied.value.document;
    // A step may take ids beyond the revisions it reports (a split copies a
    // pending property change under a new id), so the package says which
    // remain free.
    const occupied = packageIdOccupancy(current.package).revisions;
    pool = pool.filter((candidate) => !occupied.has(candidate));
    return Result.ok(compiled.value.selection);
  };

  let caret: TextPosition = from;
  if (from.blockId !== to.blockId || from.offset !== to.offset) {
    const cleared = step({ type: "replaceText", from, to, text: "" });
    if (cleared.isErr()) return Result.err(cleared.error);
    caret = cleared.value;
  }
  // An open first paragraph joins the text before the caret.
  const head = blocks.at(0);
  if (openStart === 1 && head?.type === "paragraph") {
    blocks.shift();
    const joined = step({
      type: "replaceFragment",
      from: caret,
      to: caret,
      paragraphs: [head],
      openStart: 1,
      openEnd: 1,
    });
    if (joined.isErr()) return Result.err(joined.error);
    caret = joined.value;
  }
  const last = blocks.at(-1);
  const tail = openEnd === 1 && last?.type === "paragraph" ? last : undefined;
  if (tail !== undefined) blocks.pop();
  // The blocks go in before the caret's paragraph, split at the caret when
  // text that stays stands before it; that paragraph keeps its identity.
  const destination = paragraphAt(current, caret);
  if (destination === undefined)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        message: "The clipboard destination paragraph does not exist.",
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      }),
    );
  const textBefore = holdsContentWithin({
    paragraph: destination,
    from: 0,
    to: caret.offset,
  });
  if (textBefore) {
    // The pasted paragraphs' ids are not in the document yet: the split
    // takes one none of them holds.
    const reserved = new Set(mode.reservedBlockIds);
    for (const block of blocks)
      for (const { paraId } of blockParagraphs(block))
        if (paraId !== undefined) reserved.add(paraId);
    const fresh = freshParagraphIds({
      document: current,
      reserved,
      lease: mode.idLease,
      count: 1,
      opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    });
    if (fresh.isErr()) return Result.err(fresh.error);
    const newBlockId =
      fresh.value[0] ?? panic("A fresh paragraph id is missing.");
    const split = step({ type: "splitParagraph", at: caret, newBlockId });
    if (split.isErr()) return Result.err(split.error);
  }
  const anchor = caret.blockId;
  for (const block of blocks) {
    if (
      block.type === "blockSdt" ||
      block.type === "blockCustomXml" ||
      block.type === "preservedBlock"
    ) {
      const placed = placeContainerBlock({
        document: current,
        story: caret.story,
        anchor,
        block,
        mode,
      });
      if (placed.isErr()) return Result.err(placed.error);
      ops.push(placed.value.op);
      current = placed.value.document;
      continue;
    }
    const placed = step(
      block.type === "paragraph"
        ? {
            type: "insertParagraphs",
            story: caret.story,
            at: { type: "before", blockId: anchor },
            paragraphs: [block],
          }
        : {
            type: "tableBlock",
            operation: {
              type: DOCUMENT_OP_TYPES.INSERT_TABLE,
              story: caret.story,
              blockId: anchor,
              table: block,
              side: "before",
              paragraphAfter: "none",
            },
          },
    );
    if (placed.isErr()) return Result.err(placed.error);
  }
  let selection: TextPosition = {
    story: caret.story,
    blockId: anchor,
    offset: 0,
  };
  // An open last paragraph joins the text after the caret.
  if (tail !== undefined) {
    const joined = step({
      type: "replaceFragment",
      from: selection,
      to: selection,
      paragraphs: [tail],
      openStart: 1,
      openEnd: 1,
    });
    if (joined.isErr()) return Result.err(joined.error);
    selection = joined.value;
  }
  return Result.ok({ ops, selection });
};

type MergedCellParagraphOptions = {
  /** The document the merge left. */
  document: Document;
  op: Extract<TableEditOp, { type: "mergeCells" }>;
  /** Where the merged table was before the merge. */
  located: TableRowLocation;
};

/** The first paragraph of the cell a merge made, which keeps its place. */
const mergedCellParagraph = ({
  document,
  op,
  located,
}: MergedCellParagraphOptions): string | undefined => {
  const leading = tableGrid(located.table, op.type)
    .map(({ rows }) => rows[op.top]?.find(({ start }) => start === op.left))
    .unwrapOr(undefined);
  const table = blockListAt(
    storyBody(document, op.story).content,
    located.list,
  )[located.index];
  if (leading === undefined || table?.type !== "table") return undefined;
  const cell = table.rows[op.top]?.cells[leading.index];
  return cell === undefined
    ? undefined
    : storyParagraphs({ content: cell.content }).at(0)?.paragraph.paraId;
};

/**
 * Zero-width markers and empty leaves never supply the formatting of typed
 * content; for body text, neither does a note reference.
 */
const visibleFormattingSpans = (
  paragraph: Paragraph,
  typed: TypedTextFormatting,
) =>
  leafSpans(paragraph.content).filter(
    ({ node, ancestors, before, after }) =>
      after.offset > before.offset &&
      !ancestors.some(isRemovedRevisionNode) &&
      !(
        typed === TYPED_TEXT_FORMATTING.BODY_TEXT &&
        (node.type === "footnoteRef" || node.type === "endnoteRef")
      ),
  );

/** Capture authored formatting before deletion, using only visible runs. */
const formattingAt = (
  paragraph: Paragraph,
  offset: number,
  typed: TypedTextFormatting,
): TextFormatting | undefined => {
  const spans = visibleFormattingSpans(paragraph, typed);
  const span =
    spans.find(
      ({ before, after }) => before.offset <= offset && offset < after.offset,
    ) ??
    spans.findLast(({ after }) => after.offset <= offset) ??
    spans.find(({ before }) => before.offset > offset);
  return span?.ancestors.findLast((ancestor) => ancestor.type === "run")
    ?.formatting;
};

type AuthoredFormattingOptions = {
  from: TextPosition;
  to: TextPosition;
  typed: TypedTextFormatting;
};

const authoredFormatting = (
  document: Document,
  { from, to, typed }: AuthoredFormattingOptions,
): TextFormatting | undefined => {
  const body = storyBody(document, from.story);
  const location = findStoryParagraph(body, from.blockId);
  if (location === undefined)
    panic(
      "A validated input paragraph must exist when reading authored formatting.",
    );
  const { paragraph } = location;
  const collapsed = from.blockId === to.blockId && from.offset === to.offset;
  if (!collapsed) return formattingAt(paragraph, from.offset, typed);
  const left = visibleFormattingSpans(paragraph, typed).some(
    ({ before }) => before.offset < from.offset,
  );
  if (left) return formattingAt(paragraph, from.offset - 1, typed);
  // A caret at the start can inherit from the preceding physical paragraphs
  // only while their removed breaks join them into this logical paragraph.
  const blocks = blockListAt(body.content, location.list);
  for (let index = location.index - 1; index >= 0; index -= 1) {
    const previous = blocks[index];
    if (previous?.type !== "paragraph" || !joinsFollowingParagraph(previous))
      break;
    if (paragraphVisibleText(previous) !== "")
      return formattingAt(previous, paragraphLength(previous), typed);
  }
  return formattingAt(paragraph, from.offset, typed);
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
  typed: TypedTextFormatting,
) => {
  const authored =
    runProps ?? authoredFormatting(document, { from, to, typed });
  return runPropsPatch === undefined
    ? authored
    : applyFormattingPatch(authored, runPropsPatch);
};

type ReplaceTextIntent = Extract<EditorIntent, { type: "replaceText" }>;

/** Build one replacement's authored runs without dividing its review record. */
const replacementContent = (
  document: Document,
  intent: ReplaceTextIntent,
  typed: TypedTextFormatting,
): Result<ParagraphContent[], DocumentOpRefusal> => {
  const { text, formattingSources, runProps, runPropsPatch } = intent;
  if (text === "") return Result.ok([]);
  if (formattingSources === undefined) {
    const formatting = intentRunFormatting(document, intent, typed);
    return Result.ok([
      {
        type: "run",
        ...(formatting === undefined ? {} : { formatting }),
        content: typedRunContent(text),
      },
    ]);
  }
  const invalidSpans = () =>
    Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OFFSET,
        message: "The formatting spans do not cover the replacement text.",
        opType: DOCUMENT_OP_TYPES.INSERT_TEXT,
      }),
    );
  const runs: { formatting: TextFormatting | undefined; end: number }[] = [];
  for (const { at, end, type } of formattingSources) {
    const previous = runs.at(-1);
    if (
      !Number.isInteger(end) ||
      end <= (previous?.end ?? 0) ||
      end > text.length
    )
      return invalidSpans();
    const paragraph = paragraphAt(document, at);
    if (paragraph === undefined)
      return Result.err(
        new DocumentOpRefusal({
          reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
          message:
            "A source paragraph for replacement formatting does not exist.",
          opType: DOCUMENT_OP_TYPES.INSERT_TEXT,
        }),
      );
    const authored =
      runProps ??
      (type === "caret"
        ? authoredFormatting(document, { from: at, to: at, typed })
        : formattingAt(paragraph, at.offset, typed));
    const formatting =
      runPropsPatch === undefined
        ? authored
        : applyFormattingPatch(authored, runPropsPatch);
    if (
      previous !== undefined &&
      formattingEquals(previous.formatting, formatting)
    )
      previous.end = end;
    else runs.push({ formatting, end });
  }
  if (runs.at(-1)?.end !== text.length) return invalidSpans();
  const content: ParagraphContent[] = [];
  let start = 0;
  for (const { formatting, end } of runs) {
    content.push({
      type: "run",
      ...(formatting === undefined ? {} : { formatting }),
      content: typedRunContent(text.slice(start, end)),
    });
    start = end;
  }
  return Result.ok(content);
};

/** The ending physical mark owns one logical paragraph's properties. */
const createParagraphMarkOwnerLookup = (document: Document) => {
  const stories = new Map<string, Map<string, Paragraph>>();
  return (at: TextPosition) => {
    const key = JSON.stringify(at.story);
    let owners = stories.get(key);
    if (owners === undefined) {
      owners = new Map();
      // A position names a box's paragraph through the story drawing it.
      for (const { paragraphs } of editorParagraphGroupsThroughTextBoxes(
        document,
        at.story,
      )) {
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

type SplitAtOptions = {
  document: Document;
  at: TextPosition;
  newBlockId: string;
  newHalf: SplitHalf;
  /** The revision and identity fields of the split and of the first half's properties. */
  fields: { revision?: RevisionStamp; newIds?: NewIds };
};

/**
 * The ops of a split at `at`. The mark the split creates ends the first half,
 * which starts a logical paragraph of its own and takes the surviving mark's
 * properties; the paragraph's own mark ends the second, which keeps its
 * properties. Where the second half is the new paragraph, the first keeps
 * the paragraph's id, so it takes the surviving mark's properties by a
 * property change of its own. They differ from the paragraph's only where
 * its mark is pending deletion and a later paragraph's mark survives it.
 */
const splitAtOps = ({
  document,
  at,
  newBlockId,
  newHalf,
  fields,
}: SplitAtOptions): (SplitBlockOp | SetParagraphPropsOp)[] => {
  const split = {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at,
    newBlockId,
    newHalf,
    ...fields,
  } as const;
  if (newHalf === SPLIT_HALVES.FIRST)
    return [{ ...split, newParagraph: splitParagraphFields(document, at) }];
  const paragraph = paragraphAt(document, at);
  const survivor = createParagraphMarkOwnerLookup(document)(at);
  const own = paragraph?.formatting;
  const madeSecond = {
    ...split,
    newParagraph: own === undefined ? {} : { formatting: own },
  };
  const existing = paragraphPropertiesOf(own);
  const desired = paragraphPropertiesOf(survivor?.formatting);
  if (survivor === undefined || structurallyEqual(existing, desired))
    return [madeSecond];
  return [
    madeSecond,
    {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      story: at.story,
      blockId: at.blockId,
      patch: Object.fromEntries([
        ...Object.keys(existing ?? {}).map((key) => [key, null]),
        ...Object.entries(desired ?? {}),
      ]),
      whenEmpty:
        survivor.formatting === undefined
          ? EMPTY_PROPERTY_SETS.OMIT
          : EMPTY_PROPERTY_SETS.KEEP,
      ...(fields.revision === undefined ? {} : { revision: fields.revision }),
    },
  ];
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

/**
 * Whether a copied paragraph mark can go beside a suggested paste's
 * insertion of that break: a pending deletion of it holding no cut of its
 * own, which accepting the paste leaves pending as the direct paste does.
 */
const goesBesidePastedBreak = (mark: ParagraphMarkChange): boolean =>
  mark.kind === "del" &&
  mark.resolutionJoin === undefined &&
  mark.splitDate === undefined;

/**
 * The paste's inserted `mark` holding the copied deletion `copied` beside it.
 * A break that cut no record keeps no cut depth, so accepting the paste
 * leaves the deletion as the direct paste copies it.
 */
const besidePastedBreak = (
  mark: ParagraphMarkChange,
  copied: ParagraphMarkChange,
): ParagraphMarkChange => {
  const { resolutionJoin, ...uncut } = mark;
  return {
    ...(resolutionJoin === 0 ? uncut : mark),
    deletion: { info: copied.info },
  };
};

type IdentifyClipboardParagraphsOptions = {
  document: Document;
  paragraphs: readonly Paragraph[];
  mode: EditorIntentMode;
  firstBlockId: number | undefined;
};

const identifyClipboardParagraphs = (
  options: IdentifyClipboardParagraphsOptions,
): Result<Paragraph[], DocumentOpRefusal> =>
  identifyClipboardContent({
    ...options,
    controls: [],
    identifyLast: false,
  }).map(({ paragraphs }) => paragraphs);

type IdentifyClipboardContentOptions = IdentifyClipboardParagraphsOptions & {
  /** The properties of the block content controls holding pasted paragraphs. */
  controls: readonly SdtProperties[];
  /** Whether the last paragraph takes a new id too. */
  identifyLast: boolean;
};

const identifyClipboardContent = ({
  document,
  paragraphs,
  controls,
  identifyLast,
  mode,
  firstBlockId,
}: IdentifyClipboardContentOptions): Result<
  { paragraphs: Paragraph[]; controls: SdtProperties[] },
  DocumentOpRefusal
> => {
  const copy = cloneModel([...paragraphs]);
  const controlCopies = cloneModel([...controls]);
  let occupancy: ReturnType<typeof packageIdOccupancy> | undefined;
  const packageOccupancy = () =>
    (occupancy ??= packageIdOccupancy(document.package));
  const assignedIds = {
    [IDENTITY_SPACES.REVISION]: new Set(mode.newIds?.revision),
    [IDENTITY_SPACES.CONTROL]: new Set(mode.newIds?.control),
  };
  if (mode.type === "suggesting")
    assignedIds[IDENTITY_SPACES.REVISION].add(mode.revision.id);
  const blockIds = new Set<string>();
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
  let destinationMarkers: { ids: Set<number>; names: Set<string> } | undefined;
  const markersInDestination = () => {
    if (destinationMarkers !== undefined) return destinationMarkers;
    const ids = new Set<number>();
    const names = new Set<string>();
    visit(document.package, (value) => {
      if (
        "type" in value &&
        "id" in value &&
        typeof value.id === "number" &&
        isCopiedRangeMarker(value)
      )
        ids.add(value.id);
      if (
        isCopiedRangeStart(value) &&
        "name" in value &&
        typeof value.name === "string"
      )
        names.add(value.name);
    });
    destinationMarkers = { ids, names };
    return destinationMarkers;
  };
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
    const { names } = markersInDestination();
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
      taken: (candidate) => {
        const key = formatParaId(candidate);
        return (
          blockIds.has(key) ||
          packageOccupancy().paragraphs.has(key) ||
          (mode.reservedBlockIds?.has(key) ?? false)
        );
      },
    }).at(0);
    if (fresh === undefined) return undefined;
    const id = formatParaId(fresh);
    blockIds.add(id);
    nextBlock = fresh + 1;
    return id;
  };
  // A leased author mints revision and content-control ids from its lease.
  let leaseExhausted = false;
  const freshId = (space: IdentitySpace | "bookmark", oldId: number) => {
    const key = `${space}:${oldId}`;
    const existing = remapped.get(key);
    if (existing !== undefined) return existing;
    const assigned =
      space === "bookmark" ? markersInDestination().ids : assignedIds[space];
    const taken = (candidate: number) =>
      assigned.has(candidate) ||
      (space !== "bookmark" &&
        (space === IDENTITY_SPACES.REVISION
          ? packageOccupancy().revisions
          : packageOccupancy().controls
        ).has(candidate));
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
        taken,
      }).at(0);
      if (leased === undefined) leaseExhausted = true;
      id = leased ?? MAX_REVISION_ID + 1;
    }
    while (taken(id)) id += 1;
    assigned.add(id);
    remapped.set(key, id);
    nextInSpace.set(space, id + 1);
    return id;
  };
  const remintedControls = new WeakSet<object>();
  const keptMarkers: { keptCut: unknown }[] = [];
  visit([copy, controlCopies], (value) => {
    if ("keptCut" in value) keptMarkers.push(value);
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
  for (const marker of keptMarkers) {
    const cut = marker.keptCut;
    if (!isKeptCut(cut)) continue;
    // Only copied revisions can hold the pasted cut.
    marker.keptCut = remapKeptCutHolders(cut, (id) =>
      remapped.get(slotKey({ space: IDENTITY_SPACES.REVISION, id })),
    );
  }
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
    if (index === copy.length - 1 && !identifyLast) continue;
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
  return Result.ok({ paragraphs: copy, controls: controlCopies });
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
  // Stated properties, as the join gives them: what a list level or style
  // supplies the survivor inherits again from its own numbering.
  const setting = paragraphPropertiesSetting({
    story: intent.story,
    blockId: survivorId,
    from: authoredParagraphFormatting(survivor),
    to: source === undefined ? undefined : authoredParagraphFormatting(source),
  });
  return setting === undefined
    ? undefined
    : {
        ...setting,
        whenEmpty:
          source?.formatting === undefined
            ? EMPTY_PROPERTY_SETS.OMIT
            : EMPTY_PROPERTY_SETS.KEEP,
        ...(revision === undefined ? {} : { revision }),
      };
};

type JoinedRunPropertiesOptions = {
  document: Document;
  /** The operations so far: the run's selected content deleted, none of its paragraphs joined. */
  ops: readonly DocumentOp[];
  story: OpStory;
  run: readonly ParagraphLocation[];
};

/**
 * The paragraph properties a directly joined run of paragraphs has where
 * pending mark deletions join paragraphs to it, and the paragraph kept, which
 * takes them; none where joining the run's paragraphs gives them.
 */
const joinedRunProperties = ({
  document,
  ops,
  story,
  run,
}: JoinedRunPropertiesOptions): Result<
  { blockId: string; formatting: ParagraphFormatting | undefined } | undefined,
  DocumentOpRefusal
> => {
  const first = run.at(0);
  const last = run.at(-1);
  if (first === undefined || last === undefined || first === last)
    return Result.ok(undefined);
  const sourceIn = (source: Document) =>
    joinedPropertiesSource({
      blocks: blockListAt(storyBody(source, story).content, first.list),
      first: first.index,
      last: last.index,
    });
  if (sourceIn(document) === undefined) return Result.ok(undefined);
  const deleted = applyDocumentOps(document, ops);
  if (deleted.isErr()) return Result.err(deleted.error);
  const source = sourceIn(deleted.value.document);
  return Result.ok(
    source === undefined
      ? undefined
      : {
          blockId: last.paragraph.paraId ?? "",
          formatting: authoredParagraphFormatting(source),
        },
  );
};

type JoinPropertyChangeLinkOptions = {
  /** The document before the operation that joined the paragraphs. */
  before: Document;
  /** The document after it. */
  after: Document;
  story: OpStory;
  /** The paragraph whose mark the join deletes. */
  blockId: string;
  /** The paragraph the operation set the joined properties on. */
  propertiesBlockId: string;
};

/**
 * The review op that makes the property change an operation left on the
 * paragraph its tracked join keeps belong to the paragraph-mark deletion the
 * operation recorded, so accepting or rejecting the join decides both.
 * `undefined` when the operation recorded no mark deletion (a mark already
 * pending deletion, an own added mark taken back), the paragraph holds no
 * change the operation made, or the change already belongs to a revision.
 */
export const joinPropertyChangeLink = ({
  before,
  after,
  story,
  blockId,
  propertiesBlockId,
}: JoinPropertyChangeLinkOptions): SetParagraphReviewOp | undefined => {
  const markRevision = (document: Document) => {
    const mark = findStoryParagraph(storyBody(document, story), blockId)
      ?.paragraph.pPrMark;
    return mark?.kind === "del" ? mark.info.id : mark?.deletion?.info.id;
  };
  const recorded = markRevision(after);
  if (recorded === undefined || recorded === markRevision(before))
    return undefined;
  const paragraph = findStoryParagraph(
    storyBody(after, story),
    propertiesBlockId,
  )?.paragraph;
  const change = paragraph?.propertyChanges?.at(0);
  if (
    paragraph === undefined ||
    change === undefined ||
    change.resolvesWith !== undefined ||
    findStoryParagraph(
      storyBody(before, story),
      propertiesBlockId,
    )?.paragraph.propertyChanges?.some(
      ({ info }) => info.id === change.info.id,
    ) === true
  )
    return undefined;
  const expected = reviewFieldsOf(paragraph);
  return {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
    story,
    blockId: propertiesBlockId,
    expected,
    review: {
      ...expected,
      propertyChanges: [{ ...change, resolvesWith: recorded }],
    },
  };
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
    const left = itemBefore(spans, index);
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
/**
 * The paragraphs right before `at` in its block list whose marks are pending
 * deletion, first to last: a pending join runs their words into `at`, which
 * ends it with its own mark. None when `at`'s own mark goes too, pending
 * deletion (beside its insertion or not) or a move away, since the join then
 * runs on past it.
 */
const paragraphsJoinedInto = (
  document: Document,
  at: { story: OpStory; blockId: string },
): string[] => {
  const body = storyBody(document, at.story);
  const located = findStoryParagraph(body, at.blockId);
  if (
    located === undefined ||
    joinsFollowingParagraph(located.paragraph) ||
    located.paragraph.pPrMark?.deletion !== undefined
  )
    return [];
  const list = blockListAt(body.content, located.list);
  const joined: string[] = [];
  for (let index = located.index - 1; index >= 0; index -= 1) {
    const block = list[index];
    if (
      block?.type !== "paragraph" ||
      !joinsFollowingParagraph(block) ||
      block.paraId === undefined
    )
      break;
    joined.push(block.paraId);
  }
  return joined.toReversed();
};

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
  // Range boundaries standing after them neither follow them nor keep them
  // from ending their container.
  const { block: following } = blockFollower(list, last.index);
  if (following?.type === "paragraph" && following.paraId !== undefined)
    return { story, blockId: following.paraId, offset: 0 };
  const preceding = list[first.index - 1];
  return {
    story,
    blockId: last.paragraph.paraId ?? "",
    offset:
      following === undefined && preceding?.type === "paragraph"
        ? paragraphLength(preceding)
        : 0,
  };
};

/**
 * `at` in `content` once the leaves `removed` picks are gone: after the
 * same leaves that stood before it.
 */
const gapWithoutLeaves = (
  content: readonly InlineNode[],
  at: Gap,
  removed: (leaf: InlineNode) => boolean,
): Gap => {
  let offset = 0;
  let zeroWidthBefore = 0;
  for (const span of leafSpans(content)) {
    if (compareGaps(span.after, at) > 0) {
      // A caret inside a run of characters keeps the characters before it.
      if (span.before.offset < at.offset) {
        offset += at.offset - span.before.offset;
        zeroWidthBefore = 0;
      }
      break;
    }
    if (removed(span.node)) continue;
    const width = span.after.offset - span.before.offset;
    if (width > 0) {
      offset += width;
      zeroWidthBefore = 0;
    } else zeroWidthBefore += 1;
  }
  return { offset, zeroWidthBefore };
};

/** A range moved to a target: the range's own edges, the target, and how open the moved content is. */
type RangeMove = {
  from: TextPosition;
  to: TextPosition;
  target: TextPosition;
  openStart: 0 | 1;
  openEnd: 0 | 1;
};
type RangeMoveOptions = {
  intent: RangeMove;
  mode: EditorIntentMode;
  /**
   * The moved content placed at the target, which the deletion of the
   * range has rebased, in the mode the deletion leaves.
   */
  insert: (
    afterDeletion: Document,
    target: TextPosition,
    mode: EditorIntentMode,
  ) => Result<CompiledEditorIntent, DocumentOpRefusal>;
};

/**
 * A move of the text between `from` and `to`: the range is deleted, as
 * `replaceText` with no text deletes it (a suggestion records the
 * deletion), and the moved content goes in where the target stands once
 * the deletion is made. A target inside the range moves nothing.
 */
/**
 * The note bodies a direct deletion pruned as it took their references away,
 * from its inverse: a move puts the references back at its target, so it
 * restores the bodies the deletion removed.
 */
const noteBodyRestorations = (inverse: readonly DocumentOp[]): DocumentOp[] =>
  inverse.flatMap((op): DocumentOp[] =>
    op.type === DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS &&
    (op.parts.footnotes !== undefined || op.parts.endnotes !== undefined)
      ? [
          {
            type: DOCUMENT_OP_TYPES.RESTORE_STORY_PARTS,
            expected: {
              ...(op.expected.footnotes === undefined
                ? {}
                : { footnotes: op.expected.footnotes }),
              ...(op.expected.endnotes === undefined
                ? {}
                : { endnotes: op.expected.endnotes }),
            },
            parts: {
              ...(op.parts.footnotes === undefined
                ? {}
                : { footnotes: op.parts.footnotes }),
              ...(op.parts.endnotes === undefined
                ? {}
                : { endnotes: op.parts.endnotes }),
            },
          },
        ]
      : [],
  );

/**
 * The document with the note bodies of `restorations` back, and the ops that
 * put them back. Each restoration expects the note parts its own deletion
 * left, so they undo the deletions newest first.
 */
const withNoteBodiesRestored = (
  document: Document,
  restorations: readonly DocumentOp[],
): Result<{ document: Document; ops: DocumentOp[] }, DocumentOpRefusal> => {
  if (restorations.length === 0) return Result.ok({ document, ops: [] });
  const ops = restorations.toReversed();
  const restored = applyDocumentOps(document, ops);
  if (restored.isErr()) return Result.err(restored.error);
  return Result.ok({ document: restored.value.document, ops });
};

const compileRangeMove = (
  document: Document,
  { intent, mode, insert }: RangeMoveOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const range = selectedParagraphRuns(document, intent.from, intent.to);
  if (range.isErr()) return Result.err(range.error);
  const sourceParagraph = paragraphAt(document, intent.from);
  const fromGap =
    intent.from.zeroWidthBefore ??
    (sourceParagraph === undefined
      ? 0
      : zeroWidthLeavesAt(sourceParagraph.content, intent.from.offset).length);
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
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(at.blockId),
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
    movedNoteBodies.push(...noteBodyRestorations(applied.value.inverse));
  }
  const restored = withNoteBodiesRestored(afterDeletion, movedNoteBodies);
  if (restored.isErr()) return Result.err(restored.error);
  afterDeletion = restored.value.document;
  deletionOps.push(...restored.value.ops);
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
  const inserted = insert(
    afterDeletion,
    target,
    modeAfterOps(afterDeletion, mode),
  );
  if (inserted.isErr()) return inserted;
  return Result.ok({
    ops: [...deletionOps, ...inserted.value.ops],
    selection: inserted.value.selection,
  });
};

/**
 * A paragraph without the move range markers it carried: a suggested move is
 * recorded as a deletion and an insertion, and a copied marker would pair
 * with no range.
 */
const withoutParagraphMoveAnchors = (paragraph: Paragraph): Paragraph =>
  Object.assign({}, paragraph, {
    content: asParagraphContent(withoutMoveAnchors(paragraph.content)),
  });

/**
 * A block without move range markers, a table's cell paragraphs and a
 * container's among them.
 */
const withoutBlockMoveAnchors = (block: PastedBlock): PastedBlock =>
  mapPastedBlockParagraphs(block, withoutParagraphMoveAnchors);

type BlockMoveIntent = Extract<EditorIntent, { type: "moveBlocks" }>;
type BlockMoveOptions = {
  intent: BlockMoveIntent;
  mode: EditorIntentMode;
  firstBlockId?: number | undefined;
};

/**
 * Blocks, tables among them, moved from a range to a target. The tables
 * standing wholly between the range's paragraphs are removed, the latest
 * first, or in suggesting mode recorded as deleted; the range's text is then
 * deleted and the blocks placed at the rebased target as a paste places them.
 * A suggestion records the move as the deletion of its source and the
 * insertion of its copy, as a moved paragraph is recorded. A target inside
 * the range, a moved table's cell among them, moves nothing.
 */
const compileBlockMove = (
  document: Document,
  { intent, mode, firstBlockId }: BlockMoveOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const { blocks, ...range } = intent;
  const paragraphs = blocks.filter(
    (block): block is Paragraph => block.type === "paragraph",
  );
  if (paragraphs.length === blocks.length)
    return compileEditorIntent(document, {
      intent: { ...range, type: "moveFragment", paragraphs },
      mode,
      ...(firstBlockId === undefined ? {} : { firstBlockId }),
    });
  const story = intent.from.story;
  const body = storyBody(document, story);
  const locations = storyParagraphs(body);
  const indexOf = (at: TextPosition) =>
    locations.findIndex(
      ({ paragraph }) => idKey(paragraph.paraId ?? "") === idKey(at.blockId),
    );
  const first = indexOf(intent.from);
  const last = indexOf(intent.to);
  const start = locations[first];
  const end = locations[last];
  if (start === undefined || end === undefined || intent.to.story !== story)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.BLOCK_NOT_FOUND,
        message: "The moved range names a paragraph the story does not hold.",
        opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
      }),
    );
  if (!sameBlockList(start.list, end.list) || last < first)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.INVALID_OPERATION,
        message: "Moved blocks start and end in one block list.",
        opType: DOCUMENT_OP_TYPES.DELETE_RANGE,
      }),
    );
  const target = indexOf(intent.target);
  if (intent.target.story === story && target > first && target < last)
    return Result.ok({ ops: [], selection: intent.target });
  const covered = blockListAt(body.content, start.list)
    .slice(start.index + 1, end.index)
    .filter(
      (block) =>
        block.type === "table" ||
        block.type === "blockSdt" ||
        block.type === "blockCustomXml" ||
        block.type === "preservedBlock",
    );
  const reserved =
    mode.reservedBlockIds === undefined
      ? {}
      : { reservedBlockIds: mode.reservedBlockIds };
  const ops: DocumentOp[] = [];
  let current = document;
  let currentMode = mode;
  // The note bodies a direct removal pruned with their references, which
  // the moved blocks carry to the target.
  const movedNoteBodies: DocumentOp[] = [];
  // The latest block first, so that no removal moves the place of another.
  for (const table of covered.toReversed()) {
    if (table.type !== "table") {
      const removed = removeContainerBlock({
        document: current,
        story,
        anchor: start.paragraph.paraId ?? "",
        block: table,
        mode: currentMode,
      });
      if (removed.isErr()) return removed;
      ops.push(removed.value.op);
      current = removed.value.document;
      movedNoteBodies.push(...noteBodyRestorations(removed.value.inverse));
      continue;
    }
    const blockId = storyParagraphs({ content: [table] }).at(0)?.paragraph
      .paraId;
    if (blockId === undefined)
      return Result.err(
        new DocumentOpRefusal({
          reason: DOCUMENT_OP_REFUSAL_REASONS.MISSING_BLOCK_ID,
          message: "A moved table holds no paragraph that addresses it.",
          opType: DOCUMENT_OP_TYPES.DELETE_TABLE,
        }),
      );
    const removed = compileEditorIntent(current, {
      intent: {
        type: "tableBlock",
        operation: { type: DOCUMENT_OP_TYPES.DELETE_TABLE, story, blockId },
      },
      mode: currentMode,
    });
    if (removed.isErr()) return removed;
    const applied = applyDocumentOps(current, removed.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    ops.push(...removed.value.ops);
    current = applied.value.document;
    if (mode.type === "editing")
      movedNoteBodies.push(...noteBodyRestorations(applied.value.inverse));
    currentMode = { ...modeAfterOps(current, currentMode), ...reserved };
  }
  const restored = withNoteBodiesRestored(current, movedNoteBodies);
  if (restored.isErr()) return Result.err(restored.error);
  current = restored.value.document;
  ops.push(...restored.value.ops);
  const moved = compileRangeMove(current, {
    intent: range,
    mode: currentMode,
    insert: (afterDeletion, at, insertMode) =>
      compileEditorIntent(afterDeletion, {
        intent: {
          type: "pasteBlocks",
          from: at,
          to: at,
          blocks:
            mode.type === "suggesting"
              ? blocks.map(withoutBlockMoveAnchors)
              : blocks,
          openStart: intent.openStart,
          openEnd: intent.openEnd,
        },
        mode: { ...insertMode, ...reserved },
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
        // A suggestion records a copy at the target and leaves the source's
        // references where they stand.
        ...(mode.type === "editing" ? { moved: true as const } : {}),
      }),
  });
  if (moved.isErr()) return moved;
  // A target at the range's edge, inside it, moves nothing either.
  if (moved.value.ops.length === 0)
    return Result.ok({ ops: [], selection: intent.target });
  return Result.ok({
    ops: [...ops, ...moved.value.ops],
    selection: moved.value.selection,
  });
};

/** Compile one batch; the caller applies it atomically and journals its exact inverse. */
// An intent naming paragraphs of a text box through the story drawing it is
// compiled in the box's own story; what it hands back names the package
// story again, as operations on the box do.
export const compileEditorIntent = (
  document: Document,
  input: CompileEditorIntentOptions,
): Result<CompiledEditorIntent, DocumentOpRefusal> => {
  const routed = textBoxOfIntent(document, input.intent);
  if (routed === undefined) return compileInStory(document, input);
  const compiled = compileInStory(document, {
    ...input,
    intent: addressedTo(input.intent, routed.story, routed.box),
  });
  if (compiled.isErr()) return compiled;
  const { ops, selection, applied } = compiled.value;
  const named: CompiledEditorIntent = {
    ops: ops.map(namingPackageStories),
    selection: namingPackageStories(selection),
  };
  if (applied !== undefined)
    named.applied = {
      ...applied,
      inverse: applied.inverse.map(namingPackageStories),
    };
  return Result.ok(named);
};

type InsertParagraphsIntent = Extract<
  EditorIntent,
  { type: "insertParagraphs" }
>;

type BuildParagraphInsertionOptions = {
  intent: InsertParagraphsIntent;
  ids: readonly string[];
  tracked: Pick<InsertBlocksOp, "revision" | "newIds">;
  mode: EditorIntentMode["type"];
};
const buildParagraphInsertion = ({
  intent,
  ids,
  tracked,
  mode,
}: BuildParagraphInsertionOptions): Result<
  { op: InsertBlocksOp; selection: TextPosition },
  DocumentOpRefusal
> => {
  const blocks = intent.paragraphs.map((paragraph, index) => ({
    ...paragraph,
    paraId: ids[index] ?? panic("A fresh paragraph id is missing."),
  }));
  const first = blocks.at(0);
  if (first === undefined)
    return Result.err(
      new DocumentOpRefusal({
        reason: DOCUMENT_OP_REFUSAL_REASONS.EMPTY_BLOCK_LIST,
        message: "No paragraphs to insert.",
        opType: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      }),
    );
  return Result.ok({
    op: {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: intent.story,
      at: intent.at,
      blocks,
      ...tracked,
      ...(mode === "suggesting" && intent.move === true
        ? { move: true as const }
        : {}),
    },
    selection: { story: intent.story, blockId: first.paraId, offset: 0 },
  });
};

type TrackedBatchOptions = {
  document: Document;
  count: number;
  reservedBlockIds?: ReadonlySet<string>;
};
type TrackedMemberOptions = {
  index: number;
  intent: InsertParagraphsIntent;
  mode: EditorIntentMode;
};
type TrackedInsertionCompilation =
  { type: "unsupported" } | { type: "operation"; op: InsertBlocksOp };
type TrackedInsertionCompiler =
  | { type: "unsupported" }
  | {
      type: "ready";
      compile: (
        options: TrackedMemberOptions,
      ) => Result<TrackedInsertionCompilation, DocumentOpRefusal>;
    };

/**
 * Reserve the selection of fresh paragraph IDs for a private monotone run.
 * The caller uses each index once, in order, and discards the whole run if
 * any step fails or changes the original class of operations.
 */
export const beginTrackedInsertionCompilation = ({
  document,
  count,
  reservedBlockIds,
}: TrackedBatchOptions): TrackedInsertionCompiler => {
  if (
    !Number.isSafeInteger(count) ||
    count <= 0 ||
    count >= ID_KIND_LIMITS.paragraph
  )
    return { type: "unsupported" };
  const ids = takeParagraphIds({
    document,
    count,
    reserved: reservedBlockIds,
    lease: undefined,
  });
  if (ids === undefined) return { type: "unsupported" };
  const newParagraphIds = new Set(ids);
  let nextIndex = 0;
  return {
    type: "ready",
    compile: ({ index, intent, mode }) => {
      const id = ids[nextIndex];
      if (
        index !== nextIndex ||
        id === undefined ||
        intent.story !== OP_STORIES.MAIN ||
        (intent.at.type !== "before" && intent.at.type !== "after") ||
        (!newParagraphIds.has(idKey(intent.at.blockId)) &&
          textBoxOfIntent(document, intent) !== undefined) ||
        intent.paragraphs.length !== 1 ||
        intent.move === true ||
        intent.newList !== undefined ||
        mode.type !== "suggesting" ||
        mode.idLease !== undefined ||
        mode.reservedBlockIds !== reservedBlockIds
      )
        return Result.ok({ type: "unsupported" });
      return buildParagraphInsertion({
        intent,
        ids: [id],
        tracked: { revision: mode.revision, newIds: mode.newIds },
        mode: mode.type,
      }).map(({ op }) => {
        nextIndex += 1;
        return { type: "operation" as const, op };
      });
    },
  };
};

const compileInStory = (
  document: Document,
  {
    intent,
    mode,
    firstBlockId,
    typedFormatting = TYPED_TEXT_FORMATTING.CARET,
    replacementFormattingDocument,
    moved,
  }: CompileEditorIntentOptions,
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
    case "pasteBlocks":
      return compileBlockPaste(document, {
        intent,
        mode,
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
        ...(moved === undefined ? {} : { moved }),
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
      // A copy of moved text is no part of the move: a move range boundary
      // that stands once the replaced content is gone is left out. Tracked,
      // that content stands until it is accepted.
      let pasted = intent.paragraphs;
      const contents = pasted.map(({ content }) => content);
      if (holdsMoveRange(contents)) {
        let standsIn = document;
        if (mode.type === "editing") {
          const cleared = compileEditorIntent(document, {
            intent: {
              type: "replaceText",
              from: intent.from,
              to: intent.to,
              text: "",
            },
            mode,
          });
          if (cleared.isErr()) return cleared;
          const without = applyDocumentOps(document, cleared.value.ops);
          if (without.isErr()) return Result.err(without.error);
          standsIn = without.value.document;
        }
        const settled = withoutRepeatedMoveRangesIn(standsIn, contents);
        pasted = pasted.map((paragraph, index) => {
          const content = settled.at(index);
          return content === undefined || content === paragraph.content
            ? paragraph
            : Object.assign({}, paragraph, { content: [...content] });
        });
      }
      const endpoint = paragraphAt(document, intent.to);
      const destinationFormatting = splitParagraphFields(
        document,
        intent.from,
      ).formatting;
      const endpointFormatting = splitParagraphFields(
        document,
        intent.to,
      ).formatting;
      // Markers and content pending removal keep no paragraph of their own:
      // accepted, a suggested paste leaves the paragraphs the direct one does.
      const keeps = (
        paragraph: Paragraph | undefined,
        from: number,
        to: number,
      ): boolean =>
        paragraph !== undefined && holdsContentWithin({ paragraph, from, to });
      const prefixBoundary =
        intent.openStart === 0 && keeps(destination, 0, intent.from.offset);
      const suffixBoundary =
        intent.openEnd === 0 &&
        endpoint !== undefined &&
        keeps(endpoint, intent.to.offset, paragraphLength(endpoint));
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
        ...pasted,
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
        // A review carrier stands for the paragraph a story keeps after a
        // table shown deleted, and only as that story's last paragraph. A
        // pasted paragraph is followed by the paragraph the paste ends in,
        // so its copy is the plain paragraph it shows.
        delete paragraph.reviewCarrier;
        const openStart = index === 0 && intent.openStart === 1;
        const openEnd = index === normalized.length - 1 && intent.openEnd === 1;
        if (!openStart && !openEnd) continue;
        // The compiler owns these cloned paragraphs. Open edges join the
        // destination mark, keeping their authored inline content.
        delete paragraph.propertyChanges;
        delete paragraph.markPropertyChange;
        delete paragraph.pPrMark;
        delete paragraph.sectionProperties;
        delete paragraph.formatting;
        const inherited = openStart
          ? destinationFormatting
          : endpointFormatting;
        if (inherited !== undefined) paragraph.formatting = inherited;
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
      const tail = paragraphs.at(-1);
      // A suggested paste inserts each pasted paragraph's break, so that
      // mark's insertion is the paste's own. A copied deletion of the break
      // goes beside it, `w:del` beside `w:ins` in `w:pPr/w:rPr`, and stays
      // pending once the paste is accepted, as a pasted inline record does.
      // CT_ParaRPr holds one `w:ins`, the paste's, so a copied added mark
      // (`w:ins`, `w:moveTo`) has no place on the break; an added mark
      // records only a `w:del` beside it, not a move away.
      if (
        mode.type === "suggesting" &&
        paragraphs
          .slice(0, -1)
          .some(
            ({ pPrMark }) =>
              pPrMark !== undefined && !goesBesidePastedBreak(pPrMark),
          )
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A copied added or moved-away paragraph mark has no place beside the tracked paste's insertion of that break.",
            opType: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
          }),
        );
      }
      // The paste ends in the destination's mark, which it keeps rather
      // than inserts. A record copied onto that mark (`w:del` or `w:ins` in
      // `w:pPr/w:rPr`, `w:pPrChange`, `w:rPrChange`) or a `w:sectPr` added to
      // it would stand once the paste is rejected: OOXML nests no mark-level
      // record inside the paste's `w:ins`, and the kept mark has none.
      if (
        mode.type === "suggesting" &&
        tail !== undefined &&
        (tail.sectionProperties !== undefined ||
          tail.pPrMark !== undefined ||
          tail.markPropertyChange !== undefined ||
          (tail.propertyChanges?.length ?? 0) > 0)
      ) {
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE,
            message:
              "A tracked paste cannot put copied paragraph review or a section break on the destination mark it keeps: rejecting the paste would leave them there.",
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
        const copiedMarkChange = paragraph.markPropertyChange;
        if (
          paragraph.pPrMark === undefined &&
          copiedMarkChange === undefined &&
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
        // The paste's own insertion of this break, which a copied deletion
        // of it goes beside.
        const pastedBreak =
          mode.type === "suggesting" &&
          paragraph.pPrMark !== undefined &&
          index < paragraphs.length - 1 &&
          current.pPrMark?.kind === "ins" &&
          current.pPrMark.deletion === undefined
            ? current.pPrMark
            : undefined;
        if (
          (paragraph.pPrMark !== undefined &&
            current.pPrMark !== undefined &&
            pastedBreak === undefined) ||
          (copiedMarkChange !== undefined &&
            current.markPropertyChange !== undefined &&
            !structurallyEqual(current.markPropertyChange, copiedMarkChange))
        ) {
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
        if (
          paragraph.pPrMark !== undefined ||
          copiedChange !== undefined ||
          copiedMarkChange !== undefined
        ) {
          const review = reviewFieldsOf(current);
          if (changes.length > 0) review.propertyChanges = changes;
          if (copiedMarkChange !== undefined)
            review.markPropertyChange = copiedMarkChange;
          if (pastedBreak !== undefined && paragraph.pPrMark !== undefined)
            review.pPrMark = besidePastedBreak(pastedBreak, paragraph.pPrMark);
          else if (paragraph.pPrMark !== undefined)
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
      const compiled = compileRangeMove(document, {
        intent,
        mode,
        insert: (afterDeletion, target, insertMode) =>
          compileEditorIntent(afterDeletion, {
            intent: {
              type: "replaceFragment",
              from: target,
              to: target,
              paragraphs:
                mode.type === "suggesting"
                  ? intent.paragraphs.map(withoutParagraphMoveAnchors)
                  : intent.paragraphs,
              openStart: intent.openStart,
              openEnd: intent.openEnd,
            },
            mode: insertMode,
            ...(firstBlockId === undefined ? {} : { firstBlockId }),
          }),
      });
      if (compiled.isErr()) return compiled;
      ops = compiled.value.ops;
      selection = compiled.value.selection;
      break;
    }
    case "moveBlocks": {
      const compiled = compileBlockMove(document, {
        intent,
        mode,
        ...(firstBlockId === undefined ? {} : { firstBlockId }),
      });
      if (compiled.isErr()) return compiled;
      ops = compiled.value.ops;
      selection = compiled.value.selection;
      break;
    }
    case "table": {
      const operation = intent.operation;
      const located = locateTableRow(document, operation);
      if (located.isErr()) return Result.err(located.error);
      const demand = tableEditParagraphDemand(
        located.value,
        mode.type === "suggesting"
          ? { ...operation, review: "suggesting" }
          : operation,
      );
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
        const end = op.column + (op.count ?? 1);
        const rows = located.value.table.rows.map((row, index) => ({
          ...row,
          cells: row.cells.filter((_cell, cellIndex) => {
            const entry = grid.value.rows[index]?.find(
              (cell) => cell.index === cellIndex,
            );
            return (
              entry !== undefined &&
              (entry.start < op.column || entry.end > end)
            );
          }),
        }));
        const body = storyBody(document, op.story);
        const survivor =
          tableRowAnchor(rows) ??
          storyParagraphs(body).find((paragraph) =>
            sameBlockList(paragraph.list, located.value.list),
          )?.paragraph.paraId ??
          tableRemovalAddress({
            body,
            story: op.story,
            location: located.value,
          })?.anchor.blockId;
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
      // A merge drops an empty cell's paragraph, and a row it closes goes
      // whole: a caret there moves to the merged cell.
      if (
        op.type === DOCUMENT_OP_TYPES.MERGE_CELLS &&
        paragraphAt(applied.value.document, {
          story: op.story,
          blockId,
          offset: 0,
        }) === undefined
      ) {
        const merged = mergedCellParagraph({
          document: applied.value.document,
          op,
          located: located.value,
        });
        if (merged === undefined)
          panic("A merge keeps the paragraphs of its merged cell.");
        blockId = merged;
      }
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
      const { table, rowIndex } = located.value;
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
          tableRemovalAddress({
            body: storyBody(document, op.story),
            story: op.story,
            location: located.value,
          })?.anchor.blockId;
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
      const joined =
        intent.joinedWords === "deleted" && mode.type !== "editing"
          ? paragraphsJoinedInto(document, {
              story: intent.story,
              blockId: first,
            })
          : [];
      return Result.ok({
        ops: [
          {
            type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
            story: intent.story,
            blockIds: [...joined, ...intent.blockIds],
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
      return buildParagraphInsertion({
        intent,
        ids: fresh.value,
        tracked,
        mode: mode.type,
      }).map(({ op, selection: caret }) => ({
        ops: [...newListOps(intent.newList), op],
        selection: caret,
      }));
    }
    case "tableBlock": {
      if (intent.operation.type === DOCUMENT_OP_TYPES.INSERT_TABLE) {
        const { story, blockId, table } = intent.operation;
        let separated = false;
        let trailing = false;
        if (
          intent.operation.anchor === undefined ||
          intent.operation.anchor === "paragraph"
        ) {
          separated = intent.operation.separated === true;
          trailing = intent.operation.paragraphAfter !== "none";
        }
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
        let at: TableInsertionPoint;
        if (intent.operation.anchor === "table")
          at = { type: before ? "beforeTable" : "afterTable", blockId };
        else if (intent.operation.anchor === "containerStart")
          at = { type: "containerStart", blockId };
        else
          at = before
            ? { type: "before", blockId: paragraphId ?? blockId }
            : { type: "after", blockId };
        sequence.push({
          type: DOCUMENT_OP_TYPES.INSERT_TABLE,
          story,
          at,
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
      const body = storyBody(document, op.story);
      const beside = storyParagraphs(body).filter(
        ({ list: other, paragraph }) =>
          sameBlockList(other, list) && paragraph.paraId !== undefined,
      );
      const position = beside.findIndex(
        ({ index }) => index > located.value.index,
      );
      const anchor =
        (beside[position] ?? beside.at(-1))?.paragraph.paraId ??
        tableRemovalAddress({ body, story: op.story, location: located.value })
          ?.anchor.blockId;
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
      const formatting = intentRunFormatting(document, intent, typedFormatting);
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
      // Source XML attributes belong to the existing runs rather than to
      // newly typed text.
      const replacement = replacementContent(
        intent.formattingSources === undefined
          ? document
          : (replacementFormattingDocument ?? document),
        intent,
        typedFormatting,
      );
      if (replacement.isErr()) return Result.err(replacement.error);
      const content = replacement.value;
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
              ...combineRecordedRevisions(steps),
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
        const controls = selectedControlText({
          paragraph,
          from,
          to,
          inside: intent.inside,
        });
        // The plan strikes the control's contents before the outside tail.
        // Write at the end of that retained control, while its struck text
        // still stands there and can receive the opened slice.
        const startControl =
          text === ""
            ? undefined
            : controls.ranges.find(
                ({ start }) =>
                  start.offset === from.offset &&
                  start.zeroWidthBefore ===
                    (from.zeroWidthBefore ??
                      zeroWidthLeavesAt(paragraph.content, from.offset).length),
              );
        const controlSlice =
          startControl === undefined
            ? undefined
            : continuedSlice(
                continuation(paragraph.content, from, startControl.end, {
                  inside: { start: true, end: true },
                  struck: true,
                }),
                content,
              );
        const planned = planTrackedReplacement(document, {
          sourceContainerPolicy: "join",
          cuts: controls.cuts,
          ...(controls.cuts.length === 0
            ? {}
            : {
                retainControls: {
                  ranges: controls.ranges,
                  moveInside: (options) =>
                    controlDeletionPatch({
                      ...options,
                      controls: controls.ids,
                    }),
                },
              }),
          ...(startControl === undefined || controlSlice === undefined
            ? {}
            : {
                retainedControlInsertion: {
                  at: startControl.end,
                  slice: controlSlice,
                },
              }),
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
        const struckInside = strikeInsideControl({
          document,
          planned: planned.value,
          from,
          controls: controls.ids,
          text,
        });
        if (struckInside.isErr()) return Result.err(struckInside.error);
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
        for (const op of struckInside.value) {
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
        // Text replacing a paragraph whose words are exactly a hyperlink's
        // takes a copy of the link, as the exact-text rule gives it.
        const words = paragraph.content.filter(
          (node) => !isRangeBoundary(node) && node.type !== "commentReference",
        );
        const [link] = words;
        const linked =
          words.length === 1 &&
          link?.type === "hyperlink" &&
          link.children.every((child) => child.type === "run")
            ? [withChildren(link, content)]
            : content;
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
              ...linked,
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
      const joinedProperties: {
        blockId: string;
        formatting: ParagraphFormatting | undefined;
      }[] = [];
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
            const segmentsKeeping = (keepsLoneEmptyContainers: boolean) =>
              replacementDeletionSegments({
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
                mode: { type: "editing", keepsLoneEmptyContainers },
              });
            // Empty containers in an otherwise untouched region survive a
            // replacement. A deletion selecting just those records removes them.
            const keeping = segmentsKeeping(true);
            const segments =
              keeping.length === 0 &&
              text === "" &&
              idKey(from.blockId) === idKey(to.blockId)
                ? segmentsKeeping(false)
                : keeping;
            for (const segment of segments.toReversed()) {
              ops.push({
                type: DOCUMENT_OP_TYPES.DELETE_RANGE,
                ...allocationFields,
                from: { story: from.story, blockId, ...segment.from },
                to: { story: from.story, blockId, ...segment.to },
              });
            }
          }
          const properties = joinedRunProperties({
            document,
            ops,
            story: from.story,
            run,
          });
          if (properties.isErr()) return Result.err(properties.error);
          if (properties.value !== undefined)
            joinedProperties.push(properties.value);
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
      const inside = intent.inside;
      const resumed =
        text !== "" && idKey(from.blockId) === idKey(to.blockId)
          ? continuation(paragraph.content, from, to, {
              ...(inside === undefined ? {} : { inside }),
            })
          : undefined;
      const insertion =
        text === ""
          ? undefined
          : ((resumed === undefined
              ? undefined
              : continuedSlice(resumed, content)) ?? {
              openStart: 0,
              openEnd: 0,
              content,
            });
      if (deletes) {
        const deleted = applyDocumentOps(document, ops);
        if (deleted.isErr()) return Result.err(deleted.error);
        let afterDeletion = deleted.value.document;
        for (const { blockId, formatting } of joinedProperties) {
          const setting = paragraphPropertiesSetting({
            story: from.story,
            blockId,
            from: paragraphAt(afterDeletion, { ...from, blockId })?.formatting,
            to: formatting,
          });
          if (setting === undefined) continue;
          const set = applyDocumentOp(afterDeletion, setting);
          if (set.isErr()) return Result.err(set.error);
          ops.push(setting);
          afterDeletion = set.value.document;
        }
        const retirement =
          text === ""
            ? Result.ok(null)
            : planWholeParagraphReplacement({ document, from, to });
        if (retirement.isErr()) return Result.err(retirement.error);
        const ownership = retirement.value;
        if (ownership !== null) {
          const survivingIds = new Set(
            storyParagraphs(storyBody(afterDeletion, from.story)).map(
              ({ paragraph: survivor }) => idKey(survivor.paraId ?? ""),
            ),
          );
          if (ownership.blockIds.some((id) => survivingIds.has(idKey(id))))
            panic(
              "Whole-host replacement must retire its selected source paragraphs.",
            );
          const expired = (document.package.document.comments ?? [])
            .filter(
              ({ id, parentId }) =>
                parentId === undefined &&
                ownership.commentIds.has(id) &&
                !isCommentExhausted({ type: "document", document, id }) &&
                isCommentExhausted({
                  type: "document",
                  document: afterDeletion,
                  id,
                }),
            )
            .map(
              ({ id }) =>
                ({
                  type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
                  id,
                  scope: "thread",
                }) as const,
            );
          if (expired.length > 0) {
            const cleaned = applyDocumentOps(afterDeletion, expired);
            if (cleaned.isErr()) return Result.err(cleaned.error);
            ops.push(...expired);
            afterDeletion = cleaned.value.document;
          }
        }
        at = rangeEndAfterDeletion({
          before: document,
          after: afterDeletion,
          from,
          to,
        });
        if (text === "") {
          const remaining = paragraphAt(afterDeletion, at);
          if (remaining !== undefined) {
            const { selections } = selectedControlText({
              paragraph,
              from,
              to,
              inside,
            });
            for (const control of selections.toReversed()) {
              const slice = emptiedControlSlice({
                source: paragraph,
                remaining,
                control,
              });
              if (slice !== undefined)
                ops.push({
                  type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
                  at,
                  slice,
                  ...allocationFields,
                });
            }
          }
        }
      }
      if (insertion !== undefined)
        ops.push({
          type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
          at,
          slice: insertion,
          ...allocationFields,
        });
      selection = replacementCaret(at, text);
      break;
    }
    case "splitParagraph": {
      const breaksControl = inlineControlAround(
        document,
        intent.to ?? intent.at,
      );
      if (breaksControl !== undefined)
        return Result.err(
          new DocumentOpRefusal({
            reason: DOCUMENT_OP_REFUSAL_REASONS.INSIDE_INLINE_CONTROL,
            message: `A paragraph break cannot go inside the content control ${breaksControl}: it lives in one paragraph.`,
            opType: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          }),
        );
      const newHalf = intent.newHalf ?? SPLIT_HALVES.FIRST;
      if (intent.to === undefined) {
        const split = splitAtOps({
          document,
          at: intent.at,
          newBlockId: intent.newBlockId,
          newHalf,
          fields: tracked,
        });
        if (mode.type === "suggesting") {
          const plan = createTrackedPlan({
            document,
            revision: mode.revision,
            newIds: mode.newIds,
          });
          for (const op of split) {
            const appended = plan.append(op);
            if (appended.isErr()) return Result.err(appended.error);
          }
          ops = plan.ops;
        } else ops = split;
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
        for (const op of splitAtOps({
          document: plan.document(),
          at,
          newBlockId: intent.newBlockId,
          newHalf,
          fields: { newIds: mode.newIds, revision: mode.revision },
        })) {
          const split = plan.append(op);
          if (split.isErr()) return Result.err(split.error);
        }
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
        ...splitAtOps({
          document: deleted.value.document,
          at: deletion.value.selection,
          newBlockId: intent.newBlockId,
          newHalf,
          fields: allocationFields,
        }),
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
      let cancelledAddedMark = false;
      const join = {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: intent.story,
        blockId: intent.blockId,
        nextBlockId: intent.nextBlockId,
        sectionBoundary: SECTION_BOUNDARY_POLICIES.REMOVE,
      } as const;
      if (mode.type === "suggesting") {
        const plan = createTrackedPlan({
          document,
          revision: mode.revision,
          newIds: mode.newIds,
        });
        // A mark a pending deletion already took already joins the two
        // paragraphs as the view shows them; the suggestion records only the
        // properties the join gives, under that deletion.
        if (paragraph?.pPrMark?.kind !== "del") {
          const joined = plan.join(join);
          if (joined.isErr()) return Result.err(joined.error);
          cancelledAddedMark = joined.value === "cancelled-added-mark";
        } else {
          const adjacent = adjacentJoinRefusal(document, join);
          if (adjacent !== undefined) return Result.err(adjacent);
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
        const link =
          alignment === undefined
            ? undefined
            : joinPropertyChangeLink({
                before: document,
                after: plan.document(),
                story: intent.story,
                blockId: intent.blockId,
                propertiesBlockId: alignment.blockId,
              });
        ops = link === undefined ? plan.ops : [...plan.ops, link];
      } else {
        ops = [
          {
            ...join,
            ...(intent.survivor === undefined
              ? {}
              : { survivor: intent.survivor }),
          },
        ];
      }
      let offset = 0;
      if (
        (mode.type === "editing" || cancelledAddedMark) &&
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
    // A thread whose every covered character the direct edit removed goes
    // with it, as a deleted block's does: the save drops an exhausted thread,
    // so keeping it would show a comment the reopened file no longer has.
    const beforeRetirement = current;
    const retiredIds = new Set<number>();
    for (const { id, parentId } of document.package.document.comments ?? []) {
      if (
        parentId !== undefined ||
        isCommentExhausted({ type: "document", document, id }) ||
        !isCommentExhausted({ type: "document", document: current, id })
      )
        continue;
      const retire = {
        type: DOCUMENT_OP_TYPES.DELETE_COMMENT,
        id,
        scope: "thread",
      } as const;
      const retired = applyDocumentOp(current, retire);
      if (retired.isErr()) return Result.err(retired.error);
      compact.push(retire);
      steps.push(retired.value);
      current = retired.value.document;
      retiredIds.add(id);
    }
    for (const removal of linksEmptiedByRetirement({
      before: beforeRetirement,
      after: current,
      story,
      retired: retiredIds,
    })) {
      const removed = applyDocumentOp(current, removal);
      if (removed.isErr()) return Result.err(removed.error);
      compact.push(removal);
      steps.push(removed.value);
      current = removed.value.document;
    }
    // The caret stands after the leaves before it; the retired threads'
    // markers among them are gone.
    const caretParagraph =
      retiredIds.size === 0
        ? undefined
        : paragraphAt(beforeRetirement, selection);
    if (caretParagraph !== undefined) {
      const remaining = new Set(
        (current.package.document.comments ?? []).map(({ id }) => id),
      );
      selection = {
        ...selection,
        ...gapWithoutLeaves(
          caretParagraph.content,
          {
            offset: selection.offset,
            zeroWidthBefore:
              selection.zeroWidthBefore ??
              defaultInsertionGap(caretParagraph.content, selection.offset)
                .zeroWidthBefore,
          },
          (leaf) =>
            (leaf.type === "commentRangeStart" ||
              leaf.type === "commentRangeEnd" ||
              leaf.type === "commentReference") &&
            !remaining.has(leaf.id),
        ),
      };
    }
    ops = compact;
    // Typed text hands its application back; other intents apply again.
    if (intent.type === "replaceText")
      handedBack = {
        ops: [...compact],
        result: {
          ...combineEdits(document, steps),
          ...combineRecordedRevisions(steps),
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
  const ids = takeParagraphIds({ document, reserved, lease, count });
  if (ids === undefined)
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
  return Result.ok(ids);
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
 * The name of the inline content control a position stands inside, with its
 * text on both sides: its alias, tag or id. `undefined` at a control's edge
 * or outside every control.
 */
const inlineControlAround = (
  document: Document,
  at: TextPosition,
): string | undefined => {
  const paragraph = paragraphAt(document, at);
  if (paragraph === undefined) return undefined;
  const control = continuation(paragraph.content, at, at).containers.findLast(
    (container) => container.type === "inlineSdt",
  );
  if (control?.type !== "inlineSdt") return undefined;
  const { alias, tag, id } = control.properties;
  if (alias !== undefined) return `"${alias}"`;
  if (tag !== undefined) return `tagged "${tag}"`;
  return id === undefined ? "here" : `with id ${String(id)}`;
};

type StrikeInsideControlOptions = {
  document: Document;
  planned: readonly DocumentOp[];
  from: TextPosition;
  controls: ReadonlySet<number>;
  text: string;
};

type SelectedControlTextOptions = {
  paragraph: Paragraph;
  from: TextPosition;
  to: TextPosition;
  inside: { start?: boolean; end?: boolean } | undefined;
};

type SelectedControl = {
  id: number;
  start: TextPosition;
  end: TextPosition;
};

/** The fully selected controls whose outer boundaries the selection leaves in place. */
const selectedControlText = ({
  paragraph,
  from,
  to,
  inside,
}: SelectedControlTextOptions) => {
  const ids = new Set<number>();
  const selections: SelectedControl[] = [];
  const ranges: { start: TextPosition; end: TextPosition }[] = [];
  if (inside === undefined || idKey(from.blockId) !== idKey(to.blockId))
    return { ids, cuts: [], ranges, selections };
  const cuts: Gap[] = [];
  const start = {
    offset: from.offset,
    zeroWidthBefore:
      from.zeroWidthBefore ??
      zeroWidthLeavesAt(paragraph.content, from.offset).length,
  };
  const end = { offset: to.offset, zeroWidthBefore: to.zeroWidthBefore ?? 0 };
  const spans = leafSpans(paragraph.content);
  const first = spans.find(({ after }) => compareGaps(after, start) > 0);
  const last = spans.findLast(({ before }) => compareGaps(before, end) < 0);
  const edgeControl = (span: (typeof spans)[number] | undefined) =>
    span?.ancestors.findLast((node) => node.type === "inlineSdt");
  const atStart = inside.start === true ? edgeControl(first) : undefined;
  const atEnd = inside.end === true ? edgeControl(last) : undefined;
  for (const control of new Set([atStart, atEnd])) {
    if (control?.properties.id === undefined) continue;
    const inControl = ({ ancestors }: (typeof spans)[number]) =>
      ancestors.includes(control);
    const controlStart = spans.find(inControl)?.before;
    const controlEnd = spans.findLast(inControl)?.after;
    if (
      controlStart === undefined ||
      controlEnd === undefined ||
      compareGaps(start, controlStart) > 0 ||
      compareGaps(controlEnd, end) > 0
    )
      continue;
    ids.add(control.properties.id);
    const selected = {
      id: control.properties.id,
      start: { ...from, ...controlStart },
      end: { ...from, ...controlEnd },
    };
    selections.push(selected);
    const cut = control === atStart ? controlEnd : controlStart;
    if (compareGaps(start, cut) >= 0 || compareGaps(cut, end) >= 0) continue;
    // Stage this piece first only if it cannot retract any already-added
    // content, which would shift the source positions of the other pieces.
    if (
      spans.some(
        (span) => inControl(span) && span.ancestors.some(isTrackedWrapper),
      )
    )
      continue;
    ranges.push(selected);
    if (!cuts.some((existing) => compareGaps(existing, cut) === 0))
      cuts.push(cut);
  }
  return {
    ids,
    cuts: cuts.toSorted(compareGaps).map((cut) => Object.assign({}, from, cut)),
    ranges,
    selections,
  };
};

type EmptiedControlSliceOptions = {
  source: Paragraph;
  remaining: Paragraph;
  control: SelectedControl;
};

/** Restore only a control that the deletion emptied and whose edge was kept. */
const emptiedControlSlice = ({
  source,
  remaining,
  control,
}: EmptiedControlSliceOptions) => {
  const sameControl = (node: InlineNode) =>
    node.type === "inlineSdt" && node.properties.id === control.id;
  if (
    leafSpans(remaining.content).some(
      ({ node, ancestors }) => sameControl(node) || ancestors.some(sameControl),
    )
  )
    return undefined;
  const resumed = continuation(source.content, control.start, control.end, {
    inside: { start: true, end: true },
  });
  const index = resumed.containers.findIndex(sameControl);
  if (index < 0) return undefined;
  return continuedSlice(
    {
      containers: resumed.containers.slice(0, index + 1),
      openStart: Math.min(resumed.openStart, index),
      openEnd: Math.min(resumed.openEnd, index),
    },
    [],
  );
};

/** What a tracked deletion may hold: no other tracked insertion or deletion. */
const isTrackedRunContent = (
  node: InlineSdt["content"][number],
): node is TrackedRunContent =>
  node.type !== "insertion" &&
  node.type !== "deletion" &&
  node.type !== "moveFrom" &&
  node.type !== "moveTo";

/**
 * A tracked replacement that selects the text from inside a control's edge
 * strikes that text inside the control. A range that takes both outer edges
 * leaves the control inside the deletion.
 */
const strikeInsideControl = ({
  document,
  planned,
  from,
  controls,
  text,
}: StrikeInsideControlOptions): Result<
  readonly DocumentOp[],
  DocumentOpRefusal
> => {
  if (controls.size === 0) return Result.ok(planned);
  // A deletion is the whole plan; a replacement's text follows it.
  const split =
    text === ""
      ? planned.length
      : planned.findIndex((op) => op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT);
  if (split <= 0) return Result.ok(planned);
  const deletion = planned.slice(0, split);
  const applied = applyDocumentOps(document, deletion);
  if (applied.isErr()) return Result.err(applied.error);
  const patch = controlDeletionPatch({
    document: applied.value.document,
    start: from,
    controls,
    operations: deletion,
  });
  return Result.ok(
    patch === undefined
      ? planned
      : [...deletion, patch, ...planned.slice(split)],
  );
};

type ControlDeletionPatchOptions = {
  document: Document;
  start: TextPosition;
  operations: readonly DocumentOp[];
  controls: ReadonlySet<number>;
};

const controlDeletionPatch = ({
  document,
  start,
  operations,
  controls,
}: ControlDeletionPatchOptions) => {
  const struck = new Set(
    operations.flatMap((op) =>
      op.type === DOCUMENT_OP_TYPES.DELETE_RANGE && op.revision !== undefined
        ? [op.revision.id]
        : [],
    ),
  );
  const paragraph = paragraphAt(document, start);
  if (paragraph?.paraId === undefined) return undefined;
  let swapped = false;
  const markControlText = (node: InlineNode): InlineNode => {
    if (
      node.type !== "deletion" ||
      !struck.has(node.info.id) ||
      node.content.length !== 1
    ) {
      const children = childNodes(node);
      if (children === undefined) return node;
      const next = children.map(markControlText);
      return next.some((child, index) => child !== children.at(index))
        ? rebuildNode(node, next)
        : node;
    }
    const [control] = node.content;
    if (
      control?.type !== "inlineSdt" ||
      control.properties.id === undefined ||
      !controls.has(control.properties.id) ||
      !control.content.every(isTrackedRunContent)
    )
      return node;
    swapped = true;
    return {
      ...control,
      content: [
        {
          ...node,
          content: control.content,
          resolutionJoins: {
            ...(node.resolutionJoins ?? { before: 0, after: 0, remove: 0 }),
            holdsControlText: true,
          },
        },
      ],
    };
  };
  const content = asParagraphContent(paragraph.content.map(markControlText));
  return swapped
    ? ({
        type: DOCUMENT_OP_TYPES.REPLACE_INLINE,
        story: start.story,
        blockId: paragraph.paraId,
        expected: paragraph.content,
        content,
      } satisfies DocumentOp)
    : undefined;
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
  // Text replacing words of one hyperlink, or typed between two of its
  // leaves, stays in the link; a range reaching past the link, or a caret at
  // its edge, leaves the text outside. A range holding every leaf of the
  // link leaves no piece of it to meet, so there the text takes a copy of
  // the link, which meets none. Tracked, the text continues the pieces of
  // the link either side of it, the struck one included: a tracked record
  // cannot stand inside a link, so the insertion records that it joins them.
  const link = caret
    ? left?.ancestors[containers.length]
    : last?.ancestors[containers.length];
  let copiesLink = false;
  if (link?.type === "hyperlink") {
    const ofLink = ({ ancestors }: (typeof spans)[number]) =>
      ancestors.includes(link);
    const inRange = ({ before, after }: (typeof spans)[number]) =>
      compareGaps(before, end) < 0 && compareGaps(after, start) > 0;
    const covered = ({ before, after }: (typeof spans)[number]) =>
      compareGaps(start, before) <= 0 && compareGaps(after, end) <= 0;
    // A range marker the range takes beside the link's text is no text
    // outside it: the deletion keeps it standing where it was.
    const marker = ({ node }: (typeof spans)[number]) =>
      isParagraphContent(node) && isRangeBoundary(node);
    const keepsLink = caret
      ? right !== undefined && ofLink(right)
      : spans.filter(inRange).every((span) => ofLink(span) || marker(span));
    if (keepsLink) {
      containers.push(link);
      copiesLink = spans.filter(ofLink).every(covered);
    }
  }
  // Each open end merges with the containers the source holds on that side;
  // at a container's own edge that side has none to meet, and a copied link
  // meets none.
  const continued = (side: (typeof spans)[number] | undefined) => {
    let depth = 0;
    while (
      depth < containers.length &&
      !(copiesLink && containers[depth]?.type === "hyperlink") &&
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
      : leavesKeptByDirectDeletion({
          spans,
          from: start,
          to: end,
          // Only a direct replacement with text gets here, which keeps an
          // empty container standing alone.
          keepsLoneEmptyContainers: true,
        }).at(-1);
  // The replacement goes past the ends of comments the range reached (see
  // `pastEndedComments`): it meets the last of them and the leaf after it,
  // so a link ending before them is one the text takes a copy of. The
  // comment goes with the words a direct edit took, and with it the link
  // piece holding only its start (`linksEmptiedByRetirement`).
  const passedEnd =
    caret || containers.length === 0 ? end : pastEndedComments(items, end);
  const passed =
    compareGaps(passedEnd, end) === 0
      ? undefined
      : spans.find(({ after }) => compareGaps(after, passedEnd) === 0);
  return {
    containers,
    openStart: continued(
      passed ?? (struck && last !== undefined ? last : (kept ?? outerLeft)),
    ),
    openEnd: continued(
      passed === undefined
        ? outerRight
        : spans.find(({ after }) => compareGaps(after, passed.after) > 0),
    ),
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
      case "hyperlink":
        wrapped = [withChildren(container, wrapped)];
        break;
      default:
        return undefined;
    }
  }
  return { openStart, openEnd, content: wrapped };
};
