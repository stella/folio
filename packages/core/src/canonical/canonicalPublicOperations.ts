/** Public batches compile against the read projection; only the canonical journal publishes. */
import {
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import { Mark, type Node as PMNode, type ResolvedPos } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import { commentAuthorIssue, trackedChangeAuthorIssue } from "@stll/docx-core";
import {
  type BlockContent,
  CONTENT_CONTROL_LOCK,
  lockedControlMessage,
  MAX_REVISION_ID,
  paragraphNumberingReference,
  type Table,
} from "@stll/docx-core/model";
import {
  applyDocumentOps,
  BATCH_ANNOTATION_KINDS,
  type BatchBlockAccess,
  type BatchOrderAddress,
  type BlockInsertionPoint,
  combineRecordedRevisions,
  compileEditorIntent,
  contentRank,
  createEditorIntentIdAllocator,
  deletionRemovesBlock,
  DOCUMENT_OP_REFUSAL_REASONS,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  DocumentOpRefusal,
  documentStories,
  type EditorIntent,
  type FinalBreakCandidate,
  type FinalBreakCarrier,
  finalBreakCarrier,
  findStoryBody,
  formatParaId,
  freshCommentId,
  groupInsertionRuns,
  ID_KIND_LIMITS,
  ID_LEASE_KINDS,
  type IdBlock,
  idKey,
  type IdLease,
  introducedDocumentIds,
  isBlockBoundary,
  isCommentAnchorCarrierType,
  joinPropertyChangeLink,
  type LeasedId,
  leasedIdCensus,
  locateParagraphThroughTextBoxes,
  type NewListInstance,
  noteDeletionOps,
  OP_STORIES,
  type OpStory,
  orderAddressingBeforeRemoval,
  orderBatchOperations,
  packageResourcesOpOf,
  type ParagraphAbsorption,
  paragraphHoldsContent,
  paragraphHoldsContentOnceAccepted,
  paragraphLength,
  type ParagraphPropsPatch,
  paragraphsThroughTextBoxes,
  pendingSuggestionIdsIn,
  type RecordedRevisionIds,
  revisionSpan,
  type RunPropsPatch,
  SPLIT_HALVES,
  type TableInsertionPoint,
  type TextPosition,
  TYPED_TEXT_FORMATTING,
} from "@stll/docx-core/ops";

import { SUGGESTED_SUPPORTED_OPERATION_TYPES } from "../ai-edits/applyTypes";
import { BatchClaims } from "../ai-edits/batch-claims";
import { describeCharacterSplit } from "../ai-edits/character-boundaries";
import {
  buildCleanBlockText,
  resolveCleanTextRange,
} from "../ai-edits/clean-text";
import { directJoinSurvivor, directJoinTakes } from "../ai-edits/directJoin";
import {
  createExistingRevisionGuard,
  EXISTING_REVISION_CONFLICT_MESSAGE,
} from "../ai-edits/existingRevisions";
import {
  hasInlineEmphasis,
  type InlineEmphasisRun,
  parseInlineEmphasisRuns,
  stripInlineEmphasisMarkers,
} from "../ai-edits/inline-emphasis";
import { compileMarkdownInsertion } from "../ai-edits/markdownInsertion";
import {
  BACKGROUND_MARK_NAMES,
  changesFromSegments,
  cutAroundRepeatedSeparators,
  planReplacementChanges,
  planTextChanges,
  touchedBackgroundStretches,
} from "../ai-edits/minimal-replacement";
import { createBatchListMinter } from "../ai-edits/newListNumbering";
import {
  cutsIntoNoteReference,
  endsWithNoteReference,
  noteReferenceEditsAllowed,
  type NoteReferenceLabels,
  noteReferenceSpansWithin,
  segmentsAroundNoteReferences,
} from "../ai-edits/note-references";
import {
  firstParagraphJoiningTable,
  insertionAfterPendingBreak,
  isDeletedPPrMark,
  joinStopsAtTableAfter,
  PENDING_BREAK_TABLE_REFUSAL,
  PENDING_JOIN_INSERTION_REFUSAL,
  type PendingJoinParagraphs,
  pendingJoinParagraphs,
  pendingMarkEditOf,
  pendingMarkEditRefusal,
  pendingParagraphPropertiesRefusal,
} from "../ai-edits/pendingJoin";
import { addRetiredCommentReceipts } from "../ai-edits/retired-comment-receipts";
import {
  collectNoteReferenceLabels,
  hashFolioAIBlockContent,
  hashFolioAIBlockText,
} from "../ai-edits/snapshot";
import {
  isInsertedPPrMark,
  snapshotBatchClaim,
  type SnapshotBatchOperation,
} from "../ai-edits/snapshotBatchFacts";
import {
  mergesRemoveRow,
  type TableRectangle,
  tableRectanglesOverlap,
} from "../ai-edits/table-mutation-plan";
import {
  applyTableRowInsertion,
  describeUnplacedCellTexts,
  findTableColumnInsertion,
  findTableRowInsertion,
  getInsertedPhysicalColumnRows,
  getTableColumnCoordinateKey,
  splitCellParagraphTexts,
} from "../ai-edits/table-row-column-mutations";
import {
  findEnclosingTableCell,
  findEnclosingTableRow,
  findOutermostTableBoundary,
  tableRectangleCutsMergedCell,
} from "../ai-edits/table-targets";
import { tableFromTemplate } from "../ai-edits/table-template";
import {
  supportsTableInsertionMode,
  type TableInsertionOperation,
} from "../ai-edits/tableInsertionModes";
import {
  buildSignatureTableNode,
  buildTableNode,
} from "../ai-edits/tableNodes";
import type {
  FolioAIBlockParagraphProperties,
  FolioAIEditNormalization,
  FolioAIEditParagraphIdentity,
  FolioAIEditSkippedOperation,
  FolioAIEditSkipReason,
  FolioAIListNumbering,
  FolioAIParagraphIndentation,
  FolioAIParagraphSpacing,
} from "../ai-edits/types";
import {
  createOperationHeadingStyles,
  findUndefinedOperationReference,
  styleDefinitionLookup,
} from "../ai-edits/undefinedReferences";
import { wordDiffSessionFromOptions } from "../ai-edits/word-diff";
import type { FolioContentStatedNumbering } from "../compare/content-types";
import {
  type ApplyFolioDocumentOperationsOptions,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationMode,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStory,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceiptsForStory,
  parseFolioDocumentOperationBatch,
  withDistinctSuggestionIds,
} from "../document-operations";
import { paragraphFormattingWithAuthoredIndentation } from "../docx/paragraphPropertySource";
import { canonicalOperationStagingAccess } from "../internal/canonicalOperationStaging";
import {
  expectRunPropertyChangeMarkAttrs,
  expectTableAttrs,
} from "../prosemirror/attrs";
import { tableAnchor } from "../prosemirror/canonicalTableCommands";
import { paragraphEndsItsContainer } from "../prosemirror/containerFinalParagraph";
import {
  standaloneTableRowsFromProseMirror,
  tableAttrsToFormatting,
} from "../prosemirror/conversion/fromProseDoc";
import { insertionGap } from "../prosemirror/insertionGaps";
import { isBlockBoundaryNode } from "../prosemirror/keptRangeBoundaries";
import { ensureParaIdsInDoc } from "../prosemirror/paraIds";
import { forEachAnnotationIdInDoc } from "../prosemirror/plugins/revisionIds";
import { surveyAcceptedReplacedAnnotations } from "../prosemirror/replacedAnnotations";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import {
  insertedColumnWidth,
  omittedSlotTakingColumn,
} from "../prosemirror/tableGridMutation";
import { paragraphGoes, tableGoes } from "../prosemirror/tableRunIn";
import {
  CANONICAL_CONDITION,
  CANONICAL_RULE,
  type CanonicalRefusalGround,
} from "../types/canonicalCapabilities";
import type {
  Document,
  Paragraph,
  ParagraphFormatting,
  Run,
  StyleDefinitions,
} from "../types/document";
import { readModel } from "../types/documentModel";
import {
  batchCommentReferenceShift,
  mapIntentPositions,
} from "./batchCommentReferences";
import { createCanonicalBatchReview } from "./canonicalBatchReview";
import {
  canonicalCommentBody,
  compileCanonicalComments,
} from "./canonicalComments";
import { compilerRefusalSkip } from "./canonicalCompilerRefusal";
import {
  insertedCommentRange,
  plainInsertedParagraphIds,
} from "./canonicalInsertedCommentRange";
import {
  planReplayParagraphIds,
  startParagraphAdmissionCursor,
} from "./canonicalInsertionReplayIds";
import { addJoinedParagraphReceipts } from "./canonicalJoinedReceipts";
import { lockViolationBetween } from "./canonicalLockGuard";
import { groundForDocumentOpRefusal } from "./canonicalRefusalGrounds";
import { keepStandingRevisionReceipts } from "./canonicalRevisionReceipts";
import {
  CANONICAL_ADMISSION,
  type CanonicalCommit,
  type CanonicalProjection,
  type CanonicalSession,
} from "./canonicalSession";
import { canonicalTrackedBatchFidelityError } from "./canonicalTrackedBatchFidelity";
import { findCommentAnchorParagraph } from "./commentAnchorParagraph";
import { controlEdgesAt } from "./controlEdges";
import { indentationAsRead, spacingSidesAsRead } from "./paragraphPatchAsRead";
import { replacementRunSources } from "./replacementRunSources";

/**
 * The table row or column a deletion removes, as the batch read the table.
 * Claims of one table's row (or column) deletions do not overlap, since each
 * removes its own cells; a second deletion of the same row or column has
 * nothing left of its own to remove and is a no-op, as in the default engine.
 */
type TableLineRemoval = {
  axis: "row" | "column";
  tablePosition: number;
  index: number;
};

const tableLineRemovalKey = ({
  axis,
  tablePosition,
  index,
}: TableLineRemoval) => `${axis}:${tablePosition}:${index}`;

type CanonicalBatchPlan = {
  claim: ReturnType<typeof snapshotBatchClaim>;
  address: BatchOrderAddress;
  removes: TableLineRemoval | null;
  /** The cells a merge or split reshapes, as the batch read its table. */
  cells?: TableRectangle;
  block: number;
  /**
   * The paragraph the edit is addressed through as it runs: `block`, or
   * the one where the insertion lands. A table placement uses a paragraph
   * from the table, so it runs before any removal of that paragraph.
   */
  placedAgainst: number;
  insertionRun: number | null;
};

type CanonicalBatchPlanningOptions = {
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
  from: number;
  to: number;
  projection: CanonicalProjection;
  tracked: boolean;
  inputIndex: number;
};

const canonicalBatchPlan = ({
  operation,
  intents,
  from,
  to,
  projection,
  tracked,
  inputIndex,
}: CanonicalBatchPlanningOptions): CanonicalBatchPlan => {
  const blockId =
    operation.type === "replaceRange" ||
    operation.type === "formatRange" ||
    operation.type === "commentOnRange"
      ? operation.range.blockId
      : operation.blockId;
  const paragraph =
    projection.paragraph(blockId) ??
    panic("A resolved batch operation lost its source paragraph", {
      id: operation.id,
    });
  const doc = projection.doc;
  const block = paragraph.start - 1;
  const claim = (item: SnapshotBatchOperation) =>
    snapshotBatchClaim({ item, doc, producesTrackedChanges: tracked });
  const content = (
    item: SnapshotBatchOperation,
    position = from,
  ): CanonicalBatchPlan => ({
    claim: claim(item),
    block,
    placedAgainst: block,
    address: { type: "content", position, rank: 1, inputIndex },
    removes: null,
    insertionRun: null,
  });
  /**
   * The block gap an insertion at `from` lands in, as every other claim names
   * block boundaries: past the paragraph's node, or before it, when `from` is
   * a text position at its end or start.
   */
  const gapAt = (side: "before" | "after") => {
    const $from = doc.resolve(from);
    if (!$from.parent.isTextblock) return from;
    return side === "after" ? $from.after() : $from.before();
  };
  const row = () =>
    findEnclosingTableRow(doc, block) ??
    panic("A resolved table operation lost its source row", {
      id: operation.id,
    });
  const insertionAnchor = (anchorId: string) => {
    const anchor = projection.paragraph(anchorId);
    if (anchor === undefined)
      return panic("An insertion lost its placement paragraph");
    return anchor.start - 1;
  };
  const cell = () =>
    findEnclosingTableCell(doc, block) ??
    panic("A resolved cell operation lost its source cell", {
      id: operation.id,
    });
  switch (operation.type) {
    case "replaceInBlock":
    case "replaceRange":
      return content({ type: "text", block, from, to });
    case "splitBlock":
      return content({ type: "split", block, from, to });
    case "formatRange":
    case "commentOnRange":
    case "commentOnBlock":
      return {
        claim: claim({ type: "annotation", block, from, to }),
        block,
        placedAgainst: block,
        address: {
          type: "annotation",
          kind:
            operation.type === "formatRange"
              ? BATCH_ANNOTATION_KINDS.FORMATTING
              : BATCH_ANNOTATION_KINDS.COMMENT,
          position: from,
          inputIndex,
        },
        removes: null,
        insertionRun: null,
      };
    case "replaceBlock":
      return content(
        intents.some((intent) => intent.type === "replaceText")
          ? { type: "rewriteBlock", block }
          : { type: "paragraphProperties", block },
      );
    case "setBlockParagraphProperties":
      return content({ type: "paragraphProperties", block });
    case "mergeBlockWithNext": {
      const join =
        intents.find((intent) => intent.type === "joinParagraphs") ??
        panic("A resolved merge lost its join intent");
      const next =
        projection.paragraph(join.nextBlockId) ??
        panic("A resolved merge lost its second paragraph");
      return content({
        type: "merge",
        block,
        next: next.start - 1,
        joinsNow: !tracked,
        setsProperties: operation.mergedParagraphProperties !== undefined,
      });
    }
    case "deleteBlock":
      return content({
        type: "deleteBlock",
        block,
        end: block + paragraph.node.nodeSize,
        node: paragraph.node,
      });
    case "insertBeforeBlock":
    case "insertAfterBlock": {
      // It executes at its text position (the address) and claims the gap.
      const plan = content({
        type: "insertion",
        at: gapAt(operation.type === "insertAfterBlock" ? "after" : "before"),
      });
      // Matching the native owner: redirected or commented insertions do not
      // share a provisional insertion run with their neighbours.
      const hasComment = operation.comment !== undefined;
      const emptyPageBreak =
        tracked && operation.pageBreakBefore === true && operation.text === "";
      // Put after a table, it is placed before the paragraph that follows.
      const placement = intents.find(
        (intent) => intent.type === "insertParagraphs",
      )?.at;
      return {
        ...plan,
        placedAgainst:
          placement?.type === "before" ||
          placement?.type === "after" ||
          placement?.type === "beforeTable" ||
          placement?.type === "afterTable"
            ? insertionAnchor(placement.blockId)
            : block,
        insertionRun: hasComment || emptyPageBreak ? null : from,
      };
    }
    case "insertTable":
    case "insertSignatureTable": {
      const plan = content({
        type: "insertion",
        at: gapAt(operation.position ?? "after"),
      });
      const insertion = intents.find(
        (intent) => intent.type === "tableBlock",
      )?.operation;
      // Placed beside a table, or a paragraph other than the one it names
      // (before the join that paragraph is part of), it needs that block
      // standing when it runs.
      return insertion?.type === DOCUMENT_OP_TYPES.INSERT_TABLE &&
        insertion.anchor !== "containerStart"
        ? { ...plan, placedAgainst: insertionAnchor(insertion.blockId) }
        : plan;
    }
    case "deleteTable": {
      const target = row();
      return content({
        type: "deleteTable",
        position: target.tablePosition,
        node: target.table,
      });
    }
    case "deleteTableRow": {
      const target = row();
      return {
        ...content({
          type: "deleteTableRow",
          deletion: target,
          row: target.table.child(target.rowIndex),
          table: target.table,
        }),
        removes: {
          axis: "row",
          tablePosition: target.tablePosition,
          index: target.rowIndex,
        },
      };
    }
    case "deleteTableColumn": {
      const target = row();
      const column = cell().leftColumnIndex;
      const plan = content({
        type: "deleteTableColumn",
        deletion: { tablePosition: target.tablePosition, columnIndex: column },
      });
      return {
        ...plan,
        address: {
          type: "column",
          tablePosition: target.tablePosition,
          columnIndex: column,
          insertion: false,
          inputIndex,
        },
        removes: {
          axis: "column",
          tablePosition: target.tablePosition,
          index: column,
        },
      };
    }
    case "insertTableColumn": {
      const insertion =
        findTableColumnInsertion({
          doc,
          blockFrom: block,
          position: operation.position ?? "after",
        }) ?? panic("A resolved column insertion lost its source gap");
      const plan = content({
        type: "tableStructure",
        axis: "column",
        target: {
          type: "tableStructure",
          tablePosition: insertion.tablePosition,
        },
      });
      return {
        ...plan,
        address: {
          type: "column",
          tablePosition: insertion.tablePosition,
          columnIndex: insertion.columnIndex,
          insertion: true,
          inputIndex,
        },
      };
    }
    case "insertTableRow":
      return content({
        type: "tableStructure",
        axis: "row",
        target: { type: "tableStructure", tablePosition: row().tablePosition },
      });
    case "mergeTableCells":
    case "splitTableCell": {
      const target = row();
      const source = cell();
      const merge = intents.find(
        (intent) =>
          intent.type === "table" &&
          intent.operation.type === DOCUMENT_OP_TYPES.MERGE_CELLS,
      );
      const rectangleOfMerge = () => {
        if (
          merge?.type !== "table" ||
          merge.operation.type !== DOCUMENT_OP_TYPES.MERGE_CELLS
        )
          return panic("A resolved cell merge lost its rectangle");
        return merge.operation;
      };
      const rectangle =
        operation.type === "mergeTableCells"
          ? rectangleOfMerge()
          : {
              left: source.leftColumnIndex,
              top: source.topRowIndex,
              right: source.rightColumnIndex,
              bottom: source.bottomRowIndex,
            };
      const plan = content({
        type: "tableStructure",
        axis: "cell",
        target: {
          type:
            operation.type === "mergeTableCells" ? "mergeCells" : "splitCell",
          tablePosition: target.tablePosition,
          rectangle,
        },
      });
      return {
        ...plan,
        cells: rectangle,
        address: {
          type: "cell",
          tablePosition: target.tablePosition,
          bottom: rectangle.bottom,
          right: rectangle.right,
          inputIndex,
        },
      };
    }
    default: {
      const unreachable: never = operation;
      return unreachable;
    }
  }
};

type PlannedOperation = {
  operation: FolioDocumentOperation;
  plan: CanonicalBatchPlan;
  joinsParagraphBefore?: true;
  collapsesRunStart?: true;
  retiresRunBreak?: RunCarrier;
  followsTableInsertion?: true;
  removesParagraphBeforeBlock?: true;
};

/** What precedes a run of paragraphs a batch deletes (`finalBreakCarrier`). */
type RunCarrier = Exclude<FinalBreakCarrier<unknown>["type"], "none">;

const ROW_SHIFTED_CELL_SHAPE =
  "The batch inserts or deletes a row that shifts the cells this merges or splits.";

/**
 * The merges and splits whose cells a batch row edit would shift or remove.
 * They name their cells by row as the batch read the table and run after
 * row edits, so a row edit before the end of the rectangle changes its cells.
 */
const rowShiftedCellShapes = <
  T extends PlannedOperation & { intents: readonly EditorIntent[] },
>(
  items: readonly T[],
): T[] => {
  const boundaries = items.flatMap(({ plan: { claim, removes }, intents }) => {
    if (removes?.axis === "row")
      return [{ table: removes.tablePosition, at: removes.index }];
    return claim.type === "tableStructure" && claim.axis === "row"
      ? intents.flatMap((intent) =>
          intent.type === "tableRow" &&
          intent.operation.type === DOCUMENT_OP_TYPES.INSERT_ROW
            ? [{ table: claim.table, at: intent.operation.at }]
            : [],
        )
      : [];
  });
  return items.filter(({ plan: { address } }) =>
    address.type === "cell"
      ? boundaries.some(
          ({ table, at }) =>
            table === address.tablePosition && at < address.bottom,
        )
      : false,
  );
};

const COLUMN_EDITED_CELL_SHAPE =
  "The batch inserts or deletes a column in the table this merges or splits.";

/**
 * The merges and splits in a table the batch inserts or deletes a column of,
 * whatever their order in the request. They name their cells by grid column
 * as the batch read the table, and a column edit moves every column after it
 * and the cells it runs through, so, as `planTableMutations` refuses them on
 * the legacy engine, they are refused in either mode.
 */
const columnEditedCellShapes = <
  T extends PlannedOperation & { intents: readonly EditorIntent[] },
>(
  items: readonly T[],
): T[] => {
  const tables = new Set(
    items.flatMap(({ plan: { address, claim }, intents }) => {
      if (claim.type === "tableStructure" && claim.axis === "column")
        return [claim.table];
      // Removing the final column compiles to a table deletion. Its public
      // address still names the column and claims the same cell geometry.
      return address.type === "column" &&
        intents.some(
          (intent) =>
            intent.type === "tableBlock" &&
            intent.operation.type === DOCUMENT_OP_TYPES.DELETE_TABLE,
        )
        ? [address.tablePosition]
        : [];
    }),
  );
  if (tables.size === 0) return [];
  return items.filter(
    ({ operation, plan: { address } }) =>
      (operation.type === "mergeTableCells" ||
        operation.type === "splitTableCell") &&
      address.type === "cell" &&
      tables.has(address.tablePosition),
  );
};

const ROW_CLOSING_CELL_SHAPE =
  "Another merge the batch makes in this table removes a row, which would move the cells this merges or splits.";

type RowClosingCellShapesOptions<T> = {
  /** The batch's operations, in the order it requests them. */
  items: readonly T[];
  doc: PMNode;
};

/**
 * The merges and splits whose cells a merge of the batch in their table
 * would move. A merge that leaves a row nothing it starts removes the row,
 * and the batch names every merge and split by row as it read the table, so
 * as `planTableMutations` does: the first merge stays, a later one whose
 * cells and the earlier ones' would empty a row is refused, and so is a
 * split in a table whose merges empty a row, whatever the order.
 */
const rowClosingCellShapes = <T extends PlannedOperation>({
  items,
  doc,
}: RowClosingCellShapesOptions<T>): T[] => {
  const kept = new Map<number, TableRectangle[]>();
  const refused: T[] = [];
  for (const item of items) {
    const { operation, plan } = item;
    if (operation.type !== "mergeTableCells" || plan.address.type !== "cell")
      continue;
    const cells =
      plan.cells ?? panic("A resolved cell merge lost its rectangle");
    const table = plan.address.tablePosition;
    const earlier = kept.get(table) ?? [];
    const rectangles = [...earlier, cells];
    if (
      earlier.length > 0 &&
      mergesRemoveRow({ doc, tablePosition: table, rectangles })
    ) {
      refused.push(item);
      continue;
    }
    kept.set(table, rectangles);
  }
  for (const item of items) {
    const { address } = item.plan;
    if (item.operation.type !== "splitTableCell" || address.type !== "cell")
      continue;
    const merges = kept.get(address.tablePosition);
    if (
      merges !== undefined &&
      mergesRemoveRow({
        doc,
        tablePosition: address.tablePosition,
        rectangles: merges,
      })
    )
      refused.push(item);
  }
  return refused;
};

const OVERLAPPING_CELL_SHAPE =
  "An earlier merge or split of the batch reshapes cells this merges or splits.";

/**
 * The merges and splits whose cells an earlier one of the batch in their
 * table reshapes too. Each names its cells as the batch read the table, so
 * once the other has run its rectangle can cut the cell that one made; as
 * `planTableMutations` does, the first stays and a later one is refused,
 * whatever the mode.
 */
const overlappingCellShapes = <T extends PlannedOperation>(
  items: readonly T[],
): T[] => {
  const kept: { table: number; cells: TableRectangle }[] = [];
  return items.filter(({ plan: { address, cells } }) => {
    if (address.type !== "cell" || cells === undefined) return false;
    const table = address.tablePosition;
    if (
      kept.some(
        (other) =>
          other.table === table && tableRectanglesOverlap(other.cells, cells),
      )
    )
      return true;
    kept.push({ table, cells });
    return false;
  });
};

/**
 * Whether a deletion of a paragraph that a pending join runs words into
 * takes those words too (`joinedWords`): tracked it does, as the joined
 * paragraph carries the deleted one's id; a suggestion stays rejectable on
 * its own and takes only the paragraph's own words. The deletion and the
 * batch order both read this rule, so an edit of the joined words runs
 * before the deletion only when the deletion takes them.
 */
const deletionTakesJoinedWords = (mode: FolioDocumentOperationMode): boolean =>
  mode === "tracked-changes";

type CanonicalBatchOrderOptions<T extends PlannedOperation> = {
  items: readonly T[];
  doc: PMNode;
  mode: FolioDocumentOperationMode;
};

/**
 * A paragraph's mark, as a block an operation addresses or removes: no
 * block position is negative.
 */
const markOf = (block: number): number => -block - 1;

/** Below every `contentRank`, so it runs first at its position. */
const RUN_START_COLLAPSE_RANK = -1;

/**
 * The paragraphs that go with the row or table a deletion naming the one at
 * `block` removes, those of its nested tables included.
 */
const removedTableParagraphs = (
  doc: PMNode,
  type: "deleteTable" | "deleteTableRow",
  block: number,
): number[] => {
  const row = findEnclosingTableRow(doc, block);
  if (row === null) return panic("A planned table deletion lost its row");
  const removed =
    type === "deleteTable"
      ? { node: row.table, position: row.tablePosition }
      : { node: row.table.child(row.rowIndex), position: row.rowPosition };
  const blocks: number[] = [];
  removed.node.descendants((child, offset) => {
    if (child.type.name !== "paragraph") return true;
    blocks.push(removed.position + 1 + offset);
    return false;
  });
  return blocks;
};

const orderCanonicalBatch = <T extends PlannedOperation>({
  items,
  doc,
  mode,
}: CanonicalBatchOrderOptions<T>): T[] => {
  const deletionOf = new Map<number, T>();
  const directTableDeletionOf = new Map<number, T>();
  for (const item of items) {
    const { claim, block } = item.plan;
    if (claim.type === "deleteBlock") deletionOf.set(block, item);
    if (
      mode === "direct" &&
      claim.type === "tableRemoval" &&
      claim.axis === "table"
    )
      directTableDeletionOf.set(claim.table, item);
  }
  const precedingTableDeletionOf = new Map<T, T>();
  if (directTableDeletionOf.size > 0) {
    for (const item of deletionOf.values()) {
      if (!item.removesParagraphBeforeBlock) continue;
      const at = doc.resolve(item.plan.block);
      let position = at.pos;
      for (let index = at.index() - 1; index >= 0; index--) {
        position -= at.parent.child(index).nodeSize;
        const tableDeletion = directTableDeletionOf.get(position);
        if (tableDeletion === undefined) break;
        precedingTableDeletionOf.set(item, tableDeletion);
      }
    }
  }
  const deleted = new Set(deletionOf.keys());
  const joinedAway = new Set(
    items.flatMap(({ plan }) =>
      plan.claim.type === "merge" && plan.claim.joinsNow ? [plan.block] : [],
    ),
  );
  const deletedSplits = new Set(
    items.flatMap(({ operation, plan }) => {
      if (
        operation.type !== "deleteBlock" ||
        !isInsertedPPrMark(doc.nodeAt(plan.block)?.attrs["pPrMark"])
      )
        return [];
      const at = doc.resolve(plan.block);
      const next = at.parent.maybeChild(at.index() + 1);
      if (next?.type.name !== "paragraph") return [];
      return items.some(
        ({ operation: other, plan: otherPlan }) =>
          other.type === "insertTable" &&
          other.position !== "before" &&
          otherPlan.block === plan.block &&
          otherPlan.claim.type === "insertion" &&
          otherPlan.claim.at ===
            plan.block + (doc.nodeAt(plan.block)?.nodeSize ?? 0),
      )
        ? [plan.block]
        : [];
    }),
  );
  /**
   * The batch's deletion of the first of the paragraphs it deletes in a row
   * right before the paragraph at `block`, if it deletes the one before it,
   * and what precedes the run, as `finalBreakCarrier` reads it: `kept` when
   * a paragraph the batch keeps stands there to take the break, `runStart`
   * at the container's start or another block. Tracked, a paragraph a
   * pending deletion already takes whole stands in the run.
   */
  const deletedRunBeforeBlock = (
    block: number,
    tracked: boolean,
  ): { first: T; carrier: RunCarrier } | undefined => {
    const at = doc.resolve(block);
    let position = block;
    let first: T | undefined;
    for (let index = at.index() - 1; index >= 0; index--) {
      const previous = at.parent.child(index);
      position -= previous.nodeSize;
      const deletion = deletionOf.get(position);
      if (deletion === undefined && tracked && goesWholeOnAcceptance(previous))
        continue;
      if (deletion === undefined) {
        if (first === undefined) return undefined;
        return {
          first,
          carrier: previous.type.name === "paragraph" ? "kept" : "runStart",
        };
      }
      first = deletion;
    }
    return first === undefined ? undefined : { first, carrier: "runStart" };
  };
  const address = (item: T): BatchOrderAddress => {
    const { operation, plan } = item;
    if (plan.address.type !== "content") return plan.address;
    const planned = plan.address;
    /** This operation's address right after `deletion`, `steps` ranks on. */
    const following = (deletion: T, steps: number): BatchOrderAddress => {
      const before = address(deletion);
      return before.type === "content"
        ? { ...planned, position: before.position, rank: before.rank + steps }
        : planned;
    };
    // The paragraph between tables can go whole only after the batch has
    // removed the tables before it. Tables after it already run first.
    const precedingTableDeletion = precedingTableDeletionOf.get(item);
    if (precedingTableDeletion !== undefined)
      return following(precedingTableDeletion, 1);
    // A direct merge into a paragraph the batch deletes runs after that
    // deletion, and after a deletion of the container's last paragraph that
    // joins across it, so it joins what then follows.
    const deletion =
      plan.claim.type === "merge" && plan.claim.joinsNow
        ? deletionOf.get(plan.claim.next)
        : undefined;
    if (deletion !== undefined) return following(deletion, 2);
    const run =
      operation.type === "deleteBlock"
        ? deletedRunBeforeBlock(plan.block, item.retiresRunBreak !== undefined)
        : undefined;
    // A direct deletion that joins the paragraph before it as it runs goes
    // after the batch's deletions of the paragraphs right before it, so it
    // joins the one that stays; so does a tracked one retiring the break of
    // the kept paragraph before them. Retiring the break of a run back to the
    // container's start or a table, it runs before the run's first paragraph
    // is deleted, and takes that paragraph's properties.
    if (
      run?.carrier === "kept" &&
      (item.joinsParagraphBefore === true || item.retiresRunBreak === "kept")
    )
      return following(run.first, 1);
    if (run?.carrier === "runStart" && item.retiresRunBreak === "runStart")
      return following(run.first, -1);
    // A final deletion collapsing a run back to its start runs first at the
    // paragraph before it, ahead of that paragraph's own deletion.
    if (item.collapsesRunStart === true)
      return { ...planned, rank: RUN_START_COLLAPSE_RANK };
    const hasDeletedSplit = deletedSplits.has(plan.block);
    return {
      ...plan.address,
      position:
        hasDeletedSplit &&
        operation.type === "deleteBlock" &&
        plan.claim.type === "deleteBlock"
          ? plan.claim.end
          : plan.address.position,
      rank: contentRank(operation, { hasDeletedSplit }),
    };
  };
  const access = (item: T): BatchBlockAccess<number> => {
    const { operation, plan } = item;
    const claim = plan.claim;
    if (claim.type === "merge")
      return {
        addresses: [plan.block],
        removes: claim.joinsNow
          ? directJoinTakes(plan.block, claim.next)
          : [claim.next],
      };
    if (operation.type === "deleteTable" || operation.type === "deleteTableRow")
      return {
        addresses: [],
        removes: removedTableParagraphs(doc, operation.type, plan.block),
      };
    if (plan.address.type === "column" || plan.address.type === "cell")
      return { addresses: [], removes: [] };
    if (operation.type === "insertBeforeBlock" && joinedAway.has(plan.block))
      return { addresses: [], removes: [] };
    // An edit of a paragraph's properties addresses its mark too.
    if (claim.type === "paragraphProperties")
      return { addresses: [plan.block, markOf(plan.block)], removes: [] };
    // An insertion needs the block it is placed against, as well as the
    // one it names, still standing when it runs.
    if (claim.type !== "deleteBlock")
      return {
        addresses:
          plan.placedAgainst === plan.block
            ? [plan.block]
            : [plan.block, plan.placedAgainst],
        removes: [],
      };
    const at = doc.resolve(plan.block);
    const paragraph =
      doc.nodeAt(plan.block) ??
      panic("A planned deletion lost its source paragraph");
    const before: FinalBreakCandidate<number>[] = [];
    let position = plan.block;
    for (let index = at.index() - 1; index >= 0; index--) {
      const previous = at.parent.child(index);
      position -= previous.nodeSize;
      if (previous.type !== paragraph.type) break;
      before.push({ paragraph: position, deleted: deleted.has(position) });
    }
    const appended = items.some(
      ({ operation: other, plan: otherPlan }) =>
        otherPlan.block === plan.block &&
        (other.type === "insertAfterBlock" ||
          (other.type === "insertTable" && other.position !== "before")),
    );
    // A paragraph the batch puts right before it stays and takes its break.
    const prepended = items.some(
      ({ operation: other, plan: otherPlan }) =>
        otherPlan.block === plan.block && other.type === "insertBeforeBlock",
    );
    const carrier = finalBreakCarrier(before);
    const removes = deletionRemovesBlock({
      endsContainer:
        !appended && paragraphEndsItsContainer(at, paragraph.type.name),
      carrier: prepended ? "kept" : carrier.type,
    });
    // Resolved from before it, the deletion retires the break of the
    // paragraph the batch keeps there: a property edit of that paragraph
    // runs first, as one at a time.
    const takesCarrierMark =
      !prepended &&
      carrier.type === "kept" &&
      plan.address.type === "content" &&
      plan.address.position < plan.block;
    // Tracked, it takes the paragraphs a pending join runs into it too
    // (`joinedWords`): an edit of them runs first, as one at a time.
    const joined: number[] = [];
    if (
      deletionTakesJoinedWords(mode) &&
      !isDeletedPPrMark(paragraph.attrs["pPrMark"])
    )
      for (const { paragraph: previous } of before) {
        if (!isDeletedPPrMark(doc.nodeAt(previous)?.attrs["pPrMark"])) break;
        joined.push(previous);
      }
    return {
      addresses: [],
      removes: [
        ...(removes || item.followsTableInsertion === true ? [plan.block] : []),
        ...(takesCarrierMark ? [markOf(carrier.paragraph)] : []),
        ...joined,
      ],
    };
  };
  const runs = groupInsertionRuns(
    orderBatchOperations(items, address),
    ({ plan }) => plan.insertionRun,
    ({ operation }) => operation.type === "setBlockParagraphProperties",
  );
  return orderAddressingBeforeRemoval(runs, (run) => ({
    addresses: run.flatMap((item) => access(item).addresses),
    removes: run.flatMap((item) => access(item).removes),
  })).flat();
};

type CanonicalPublicOperationRefusal = NonNullable<
  FolioAIEditSkippedOperation["canonicalRefusal"]
>;

/**
 * A refused resolution under `reason`, with the skip's detail, and the
 * ground the skip reports where the reason alone does not say why.
 */
type ResolutionRefusal = (
  reason: FolioAIEditSkipReason,
  message?: string,
  ground?: CanonicalRefusalGround,
) => Result<never, CanonicalPublicOperationResolutionError>;

/** A new public kind cannot bypass a canonical compiler/refusal decision. */
export const CANONICAL_PUBLIC_OPERATION_DISPOSITIONS: {
  readonly replaceInBlock: "compile";
  readonly replaceRange: "compile";
  readonly replaceBlock: "compile";
  readonly splitBlock: "compile";
  readonly formatRange: "compile";
  readonly mergeBlockWithNext: "compile";
  readonly setBlockParagraphProperties: "compile";
  readonly insertAfterBlock: "compile";
  readonly insertBeforeBlock: "compile";
  readonly deleteBlock: "compile";
  readonly commentOnBlock: "compile";
  readonly commentOnRange: "compile";
  readonly insertTable: "compile";
  readonly insertSignatureTable: "compile";
  readonly deleteTable: "compile";
  readonly insertTableRow: "compile";
  readonly deleteTableRow: "compile";
  readonly insertTableColumn: "compile";
  readonly deleteTableColumn: "compile";
  readonly mergeTableCells: "compile";
  readonly splitTableCell: "compile";
} = {
  replaceInBlock: "compile",
  replaceRange: "compile",
  replaceBlock: "compile",
  splitBlock: "compile",
  formatRange: "compile",
  mergeBlockWithNext: "compile",
  setBlockParagraphProperties: "compile",
  insertAfterBlock: "compile",
  insertBeforeBlock: "compile",
  deleteBlock: "compile",
  commentOnBlock: "compile",
  commentOnRange: "compile",
  insertTable: "compile",
  insertSignatureTable: "compile",
  deleteTable: "compile",
  insertTableRow: "compile",
  deleteTableRow: "compile",
  insertTableColumn: "compile",
  deleteTableColumn: "compile",
  mergeTableCells: "compile",
  splitTableCell: "compile",
} as const satisfies Record<FolioDocumentOperation["type"], "compile">;

/**
 * A request the operation vocabulary does not take, or names a place the
 * document has no address for: refused as the request, not as a capability
 * the session lacks.
 */
const INVALID_REQUEST: CanonicalRefusalGround = {
  condition: CANONICAL_CONDITION.invalidRequest,
};
/** A place the operation named that an earlier change of its batch moved on. */
const STALE: CanonicalRefusalGround = { condition: CANONICAL_CONDITION.stale };
/** A shape the session built itself and then could not read back: a defect. */
const INTERNAL: CanonicalRefusalGround = {
  condition: CANONICAL_CONDITION.internal,
};

/**
 * The public reason a batch the session refused to stage reports: a no-op
 * for one that changes nothing, a content control's lock as the lock, a
 * document that moved on as a stale range, and any other refusal as a block
 * the session does not take, under the session's own ground.
 */
const stagingRefusalReason = (
  reason: "refused" | "noChange",
  ground: CanonicalRefusalGround,
): FolioAIEditSkipReason => {
  if (reason === "noChange") return "noopOperation";
  if (ground.rule === CANONICAL_RULE.contentControlLock)
    return "contentControlLock";
  if (ground.condition === CANONICAL_CONDITION.stale) return "staleRange";
  return "unsupportedBlock";
};

export type CanonicalPublicOperationOptions = Omit<
  ApplyFolioDocumentOperationsOptions,
  "view" | "createUndoHandle" | "createCommentId"
>;

type CommentOperation = {
  [
    Operation in FolioDocumentOperation as Operation["type"]
  ]: "comment" extends keyof Operation ? Operation : never;
}[FolioDocumentOperation["type"]];

type PreparedOperation = {
  intents: readonly EditorIntent[];
  normalizations: readonly FolioAIEditNormalization[];
  from: number;
  to: number;
  /** Deleted source spans of a tracked text replacement, in source order. */
  deletionSpans?: readonly { from: number; to: number }[];
  /**
   * A deletion of the whole paragraph by its identity, with a block after
   * it: an insertion beside that paragraph keeps its own anchor.
   */
  removesParagraphBeforeBlock?: true;
  /**
   * Resolved beside a paragraph no paragraph follows, counting on the
   * batch's other half there: the deletion of that paragraph, or an
   * insertion after it. Without that half it resolves again on its own.
   */
  besideDeletion?: true;
  /**
   * A direct deletion of a paragraph no paragraph follows, which joins the
   * paragraph before it: the one before it when the deletion runs, which an
   * insertion the batch order runs first may have put there.
   */
  joinsParagraphBefore?: true;
  /**
   * A tracked deletion of a paragraph no paragraph follows, which retires
   * the break before it: the break of the paragraph before it when the
   * deletion runs, which an insertion the batch order runs first may have
   * put there.
   */
  retiresBreakBefore?: true;
  /**
   * A deletion of a container's final paragraph after a table the batch
   * puts a paragraph after: that insertion runs first, and the deletion takes
   * the paragraph it put there as the one before it.
   */
  followsTableInsertion?: true;
  /**
   * A direct deletion of a container's final paragraph whose batch deletes
   * every paragraph before it back to the container's start or another
   * block (`finalBreakCarrier`'s `runStart`): it runs first there, while the
   * paragraph before it stands, and joins into it, which keeps its identity
   * and properties; that paragraph's own deletion follows.
   */
  collapsesRunStart?: true;
  /**
   * A tracked deletion of a container's final paragraph whose batch deletes
   * the paragraphs right before it. With a kept one before them, it runs
   * after their deletions, so it retires the break before them; back to the
   * container's start or another block, it runs before the first one's
   * deletion, retiring that paragraph's break and taking its properties.
   */
  retiresRunBreak?: RunCarrier;
  /**
   * A direct deletion of a paragraph that another paragraph follows, which
   * the batch may delete first: then it deletes a final paragraph as it runs.
   */
  mayEndContainer?: true;
  /**
   * A deletion of a paragraph that stands between two tables once the
   * batch's blocks land (`deletedParagraphPlacement`): it takes the words
   * only, in every mode, and nothing as it runs takes the paragraph whole.
   */
  keepsTablesApart?: true;
  /**
   * The paragraph whose stated properties a block insertion copies, and the
   * paragraphs it inserts from a given copy. The batch runs in its execution
   * order, so the copy is read when the insertion runs: an edit of that
   * paragraph's properties the order runs first is in it.
   */
  inheritsFrom?: {
    blockId: string;
    /** The paragraph's own break was already pending removal when read. */
    markDeleted: boolean;
    /**
     * The insertion goes right after the paragraph it copies, which keeps its
     * own properties even when the batch joins it to the next one first.
     */
    followsAnchor: boolean;
    /** The operation names this paragraph, not a paragraph a join reaches. */
    namesSource: boolean;
    paragraphs: (formatting: ParagraphFormatting | undefined) => Paragraph[];
  };
} & (
  | { type: "plain"; operation: FolioDocumentOperation }
  | {
      type: "commented";
      operation: CommentOperation;
      comment: {
        text: string;
        from: TextPosition;
        to: TextPosition;
        /**
         * The comment covers the paragraphs the operation inserts, whose
         * identities the compiler gives: its range is read after compiling.
         */
        inserted?: true;
      };
    }
);

type ResolveCommentOptions = {
  operation: FolioDocumentOperation;
  from: TextPosition;
  to: TextPosition;
};

const resolveComment = ({ operation, from, to }: ResolveCommentOptions) => {
  if ("comment" in operation && operation.comment !== undefined) {
    return {
      type: "commented",
      operation,
      comment: { text: operation.comment.text, from, to },
    } as const satisfies Pick<
      Extract<PreparedOperation, { type: "commented" }>,
      "type" | "operation" | "comment"
    >;
  }
  return { type: "plain", operation } as const;
};

/**
 * The two paragraphs a direct operation's intents join, if one paragraph's
 * content ends up in another: a deletion that keeps the paragraph before it
 * leaves nothing of the deleted paragraph there, and joins nothing.
 */
const joinedParagraphsOf = (
  operationType: FolioDocumentOperation["type"],
  intents: readonly EditorIntent[],
): FolioAIEditParagraphIdentity | undefined => {
  const join = intents.find(
    (intent): intent is Extract<EditorIntent, { type: "joinParagraphs" }> =>
      intent.type === "joinParagraphs",
  );
  if (join === undefined) return undefined;
  if (operationType === "deleteBlock" && join.survivor === "first")
    return undefined;
  return join.survivor === "first"
    ? {
        type: "joined",
        blockId: join.blockId,
        retiredBlockId: join.nextBlockId,
      }
    : {
        type: "joined",
        blockId: join.nextBlockId,
        retiredBlockId: join.blockId,
      };
};

const SPACING_FIELDS = [
  "spaceBefore",
  "spaceAfter",
  "lineSpacing",
  "lineSpacingRule",
  "beforeAutospacing",
  "afterAutospacing",
] as const satisfies readonly (keyof FolioAIParagraphSpacing)[];
const INDENTATION_FIELDS = [
  "indentLeft",
  "indentRight",
  "indentFirstLine",
  "hangingIndent",
] as const satisfies readonly (keyof FolioAIParagraphIndentation)[];

/**
 * The stated paragraph properties an operation writes: each given property
 * as it says, a spacing or indentation object as the whole direct child, so
 * an attribute it leaves out is removed, and `null` clears a property back
 * to what the style gives. The patch states the paragraph as the parser reads
 * the saved file back: the spacing sides the paragraph states are marked
 * explicit by the parser's own rule, and an indentation is held as read.
 */
const statedParagraphPatch = (
  properties: FolioAIBlockParagraphProperties,
): ParagraphPropsPatch => {
  const patch: Record<string, unknown> = {};
  if (properties.styleId !== undefined) {
    patch["styleId"] = properties.styleId;
    // The numbering the old style supplied is no longer the paragraph's:
    // list normalization derives it again from the style the patch sets.
    patch["numPrFromStyle"] = null;
  }
  if (properties.outlineLevel !== undefined)
    patch["outlineLevel"] = properties.outlineLevel;
  if (properties.alignment !== undefined)
    patch["alignment"] = properties.alignment;
  // Stated numbering as given; inheriting removes the stated value.
  const { numbering } = properties;
  if (numbering !== undefined && numbering.kind !== "newList")
    patch["numPr"] = numbering.kind === "inherit" ? null : numbering;
  const { spacing, indentation } = properties;
  if (spacing !== undefined) {
    for (const field of SPACING_FIELDS)
      patch[field] = spacing === null ? null : (spacing[field] ?? null);
    patch["spacingExplicit"] =
      spacing === null ? null : (spacingSidesAsRead(spacing) ?? null);
  }
  if (indentation !== undefined) {
    const read = indentation === null ? null : indentationAsRead(indentation);
    for (const field of INDENTATION_FIELDS)
      patch[field] = read === null ? null : (read[field] ?? null);
  }
  // Every key is a ParagraphFormatting field set to its value or null.
  return patch as ParagraphPropsPatch;
};

type ResolveParagraphPropertiesOptions = {
  mintList: ReturnType<typeof createBatchListMinter>;
  operation: Extract<
    FolioDocumentOperation,
    { type: "setBlockParagraphProperties" }
  >;
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>;
  projection: CanonicalProjection;
  refusal: ResolutionRefusal;
};

const resolveParagraphProperties = ({
  operation,
  paragraph,
  projection,
  mintList,
  refusal,
}: ResolveParagraphPropertiesOptions) => {
  const requested = operation.properties.numbering;
  const stated = statedNumbering(requested, mintList);
  const patch = statedParagraphPatch({
    ...operation.properties,
    ...(stated.numbering === undefined ? {} : { numbering: stated.numbering }),
  });
  if (Object.keys(patch).length === 0) return refusal("emptyOperation");
  const at = projection.addressAt(paragraph.start);
  if (at.isErr())
    return refusal("unsupportedBlock", undefined, at.error.ground);
  const intents: EditorIntent[] = [
    {
      type: "formatParagraph",
      at: at.value,
      patch,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
  return Result.ok({
    type: "plain" as const,
    operation,
    intents,
    normalizations: [],
    from: paragraph.start,
    to: paragraph.start + paragraph.node.content.size,
  });
};

/**
 * The intent that states `properties` on one paragraph by its identity,
 * none when they state nothing.
 */
const statedPropertiesIntents = ({
  story,
  blockId,
  properties,
  mintList,
}: {
  story: OpStory;
  blockId: string;
  properties: FolioAIBlockParagraphProperties | undefined;
  mintList: ReturnType<typeof createBatchListMinter>;
}): EditorIntent[] => {
  if (properties === undefined) return [];
  const stated = statedNumbering(properties.numbering, mintList);
  const patch = statedParagraphPatch({
    ...properties,
    ...(stated.numbering === undefined ? {} : { numbering: stated.numbering }),
  });
  if (Object.keys(patch).length === 0) return [];
  return [
    {
      type: "formatParagraph",
      at: { story, blockId, offset: 0 },
      patch,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
};

type BlockInsertion = Extract<
  FolioDocumentOperation,
  { type: "insertAfterBlock" | "insertBeforeBlock" }
>;

const isBlockInsertion = (
  operation: FolioDocumentOperation,
): operation is BlockInsertion =>
  operation.type === "insertAfterBlock" ||
  operation.type === "insertBeforeBlock";

type TableInsertion = TableInsertionOperation;

const placesBlock = (
  operation: FolioDocumentOperation,
): operation is BlockInsertion | TableInsertion =>
  isBlockInsertion(operation) ||
  operation.type === "insertTable" ||
  operation.type === "insertSignatureTable";

/** The side of the block it names that an insertion puts its block on. */
const placementSide = (
  operation: BlockInsertion | TableInsertion,
): "before" | "after" => {
  switch (operation.type) {
    case "insertBeforeBlock":
      return "before";
    case "insertAfterBlock":
      return "after";
    case "insertTable":
    case "insertSignatureTable":
      return operation.position ?? "after";
    default: {
      const unreachable: never = operation;
      return unreachable;
    }
  }
};

const placesTable = (operation: BlockInsertion | TableInsertion): boolean =>
  operation.type === "insertTable" || operation.type === "insertSignatureTable";

/**
 * Whether the block a batch puts on `side` of a block, of `placed` beside it
 * in input order, that lands furthest from it is a table: those land in that
 * order, so the last put after it is right before the block after, and the
 * first put before it right after the block before.
 */
const outermostPutIsTable = (
  placed: readonly (BlockInsertion | TableInsertion)[],
  side: "before" | "after",
): boolean => {
  const onSide = placed.filter(
    (operation) => placementSide(operation) === side,
  );
  const outermost = side === "after" ? onSide.at(-1) : onSide.at(0);
  return outermost !== undefined && placesTable(outermost);
};

type DeletedParagraphPlacementOptions = {
  /** The batch's insertions beside the paragraph, in input order. */
  beside: readonly (BlockInsertion | TableInsertion)[];
  /** Whether a table precedes the paragraph if nothing is put right before it. */
  tableBefore: boolean;
  /** Whether a table follows the paragraph if nothing is put right after it. */
  tableAfter: boolean;
  /**
   * What follows the paragraph as the batch runs, before its blocks beside
   * it land: a paragraph, another block, or nothing.
   */
  after: "paragraph" | "block" | "none";
};

/**
 * Where a batch's blocks beside a paragraph it deletes land, in every mode:
 *
 * - `keepsTablesApart`: once they land, a table stands right before the
 *   paragraph and another right after it: the last block put before it is a
 *   table, or, with none, the block before it; so is the first put after it,
 *   or, with none, the block after it, unless the batch deletes that one.
 *   The paragraph stays between them, its words removed directly and deleted
 *   when tracked, so accepting keeps its mark there and the tables apart, as
 *   the direct deletion does. Every block stays on its side.
 * - `placedAfter`: a table or another block, not a paragraph, follows it.
 *   Applied directly, from the first paragraph put before it, the blocks put
 *   before it go after it instead, ahead of the ones put after it, so they
 *   land in the requested order. The deletion then removes it by identity
 *   with a paragraph after it, and that paragraph keeps its own properties.
 *   Tracked, the paragraph stays until accepted, so every block lands on its
 *   own side.
 * - `inPlace`: a paragraph or nothing follows it; every block lands where it
 *   was put.
 */
type DeletedParagraphPlacement =
  | { type: "keepsTablesApart" }
  | { type: "placedAfter"; operationIds: ReadonlySet<string> }
  | { type: "inPlace" };

const deletedParagraphPlacement = ({
  beside,
  tableBefore,
  tableAfter,
  after,
}: DeletedParagraphPlacementOptions): DeletedParagraphPlacement => {
  const before = beside.filter(
    (operation) => placementSide(operation) === "before",
  );
  const last = before.at(-1);
  const firstAfter = beside.find(
    (operation) => placementSide(operation) === "after",
  );
  const precededByTable = last === undefined ? tableBefore : placesTable(last);
  const followedByTable =
    firstAfter === undefined ? tableAfter : placesTable(firstAfter);
  if (precededByTable && followedByTable) return { type: "keepsTablesApart" };
  if (after !== "block") return { type: "inPlace" };
  const firstParagraph = before.findIndex(
    (operation) => operation.type === "insertBeforeBlock",
  );
  return {
    type: "placedAfter",
    operationIds: new Set(
      firstParagraph < 0
        ? []
        : before.slice(firstParagraph).map(({ id }) => id),
    ),
  };
};

type ProjectedParagraph = NonNullable<
  ReturnType<CanonicalProjection["paragraph"]>
>;

type TableLandsBesideOptions = {
  /** The paragraph id the batch's blocks are put beside. */
  blockId: string;
  side: "before" | "after";
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
  besideDeletions: BesideDeletions;
};

/**
 * Whether the block the batch puts on `side` of a paragraph furthest from
 * it is a table that lands right beside it, not beside the pending join the
 * paragraph is part of (`besidePendingJoin`).
 */
const tableLandsBeside = ({
  blockId,
  side,
  projection,
  noteReferences,
  besideDeletions,
}: TableLandsBesideOptions): boolean => {
  const paragraph = projection.paragraph(blockId);
  return (
    paragraph !== undefined &&
    outermostPutIsTable(besideDeletions.placedBeside(blockId), side) &&
    besidePendingJoin({ side, paragraph, projection, noteReferences }).place
      .blockId === blockId
  );
};

type TableLandsAfterOptions = Omit<
  TableLandsBesideOptions,
  "blockId" | "side"
> & {
  paragraph: ProjectedParagraph;
};

/**
 * Whether, applied directly, a table lands right after `paragraph`: the
 * first block the batch puts after it is one, or, with none there, the first
 * it puts before the paragraph after it. No paragraph follows it then. One
 * whose own break is under review keeps the paragraph after it instead.
 */
const tableLandsAfter = ({
  paragraph,
  projection,
  noteReferences,
  besideDeletions,
}: TableLandsAfterOptions): boolean => {
  if (paragraph.node.attrs["pPrMark"] != null) return false;
  const first = besideDeletions
    .placedBeside(paragraph.blockId)
    .find((operation) => placementSide(operation) === "after");
  if (first !== undefined) return placesTable(first);
  const next = projection.doc.resolve(
    paragraph.start - 1 + paragraph.node.nodeSize,
  ).nodeAfter;
  const nextId: unknown =
    next?.type.name === "paragraph" ? next.attrs["paraId"] : undefined;
  return (
    typeof nextId === "string" &&
    tableLandsBeside({
      blockId: nextId,
      side: "before",
      projection,
      noteReferences,
      besideDeletions,
    })
  );
};

type MovedPastDeletionOptions = {
  operation: BlockInsertion | TableInsertion;
  /** The paragraph the insertion lands after. */
  place: ProjectedParagraph;
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
  besideDeletions: BesideDeletions;
  inputOrder: readonly FolioDocumentOperation[];
};

/**
 * The paragraph a direct insertion after `place` goes after instead: the
 * one right after it, which the batch deletes, when the blocks put before
 * that one go after it (`placedAfter`) from a paragraph put earlier in the
 * batch. Both sides of it are one gap once it goes, and its blocks keep the
 * batch's order there.
 */
const movedPastDeletion = ({
  operation,
  place,
  projection,
  noteReferences,
  besideDeletions,
  inputOrder,
}: MovedPastDeletionOptions): ProjectedParagraph | undefined => {
  const end = place.start - 1 + place.node.nodeSize;
  const next = projection.paragraphStartingAt(end);
  if (next === undefined || !besideDeletions.deleted.has(next.blockId))
    return undefined;
  const beside = besideDeletions.placedBeside(next.blockId);
  const order = (id: string) =>
    inputOrder.findIndex((other) => other.id === id);
  const first = beside.find(({ type }) => type === "insertBeforeBlock");
  if (first === undefined || order(first.id) > order(operation.id))
    return undefined;
  const after = projection.doc.resolve(end + next.node.nodeSize).nodeAfter;
  if (after === null) return undefined;
  const afterId: unknown =
    after.type.name === "paragraph" ? after.attrs["paraId"] : undefined;
  const options = { projection, noteReferences, besideDeletions };
  const followed =
    after.type.name === "paragraph" &&
    !tableLandsAfter({ paragraph: next, ...options });
  return deletedParagraphPlacement({
    beside,
    tableBefore: false,
    tableAfter:
      after.type.name === "table" ||
      (typeof afterId === "string" &&
        tableLandsBeside({ blockId: afterId, side: "before", ...options })),
    after: followed ? "paragraph" : "block",
  }).type === "placedAfter"
    ? next
    : undefined;
};

/**
 * The start of the block list holding the paragraph `blockId` names, at
 * `position`: where a block goes before a first block that holds no
 * paragraph to name it by, such as an empty content control.
 */
type ListStartPlace = { blockId: string; position: number };

type ResolveBlockInsertionOptions = {
  mintList: ReturnType<typeof createBatchListMinter>;
  operation: BlockInsertion;
  /** Where the insertion goes in place of beside `paragraph`. */
  listStart?: ListStartPlace;
  /** The paragraph the insertion goes beside. */
  paragraph: ProjectedParagraph;
  /** The paragraph the operation names, which a pending join may move from. */
  named: ProjectedParagraph;
  /** The paragraph whose stated properties the new paragraph copies. */
  formattingSource: ProjectedParagraph;
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
  story: OpStory;
  refusal: ResolutionRefusal;
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  /** `BesidePendingJoinOptions.landedEarlier`. */
  landedEarlier?: ReadonlySet<number>;
};

/**
 * Where the gap a block put before `boundary` lands in starts as the batch
 * runs (`insertionGap`): before the blocks right before it that accepting
 * removes whole, back to the last place among them an earlier operation of
 * the batch put a paragraph, which stays.
 */
const gapStartAsTheBatchRuns = (
  doc: PMNode,
  boundary: number,
  landed: ReadonlySet<number> | undefined,
): number => {
  let gap = insertionGap(doc, boundary, "before");
  for (const at of landed ?? []) if (at > gap && at <= boundary) gap = at;
  return gap;
};

/**
 * Where an insertion beside `paragraph` goes: beside the paragraph, or, for a
 * paragraph in a table cell, beside the outermost table, as the default
 * engine places it. The point names the paragraph on the far side of the
 * table; where none stands there (the table starts its container, or a table
 * or control stands beside it), it names the table itself.
 */
const blockInsertionPlace = (
  operation: BlockInsertion,
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>,
  projection: CanonicalProjection,
  noteReferences: NoteReferenceLabels,
  landedEarlier: ReadonlySet<number> | undefined,
): { at: TableInsertionPoint; position: number } | undefined => {
  const side = operation.type === "insertAfterBlock" ? "after" : "before";
  const table = findOutermostTableBoundary(projection.doc, paragraph.start - 1);
  // Put before a paragraph or a table, a block lands before the blocks right
  // before it that accepting removes whole, at the start of its gap, as the
  // legacy engine puts it: accepted, both sides of those blocks are one
  // place, and a block put after one of them follows this one there.
  const atGapStart = (boundary: number, position: number) => {
    const gap = gapStartAsTheBatchRuns(projection.doc, boundary, landedEarlier);
    if (gap >= boundary) return undefined;
    const $gap = projection.doc.resolve(gap);
    const first = $gap.nodeAfter;
    const previous = $gap.nodeBefore;
    const end =
      previous?.type.name === "paragraph"
        ? projection.paragraphStartingAt(gap - previous.nodeSize)
        : undefined;
    if (end !== undefined)
      return { at: { type: "after", blockId: end.blockId } as const, position };
    // Without a preceding paragraph, the pending-deleted table itself
    // supplies the address at the start of the gap.
    if (first?.type.name === "table") {
      const blockId = tableBlockAddress(projection, {
        table: first,
        tableStart: gap + 1,
      });
      if (blockId !== undefined)
        return { at: { type: "beforeTable", blockId } as const, position };
    }
    const start =
      first?.type.name === "paragraph"
        ? projection.paragraphStartingAt(gap)
        : undefined;
    return start === undefined
      ? undefined
      : { at: { type: "before", blockId: start.blockId } as const, position };
  };
  if (table === null && side === "before") {
    const place = atGapStart(paragraph.start - 1, paragraph.start);
    if (place !== undefined) return place;
  }
  if (table === null)
    return {
      at: { type: side, blockId: paragraph.blockId },
      position:
        side === "after"
          ? paragraph.start + paragraph.node.content.size
          : paragraph.start,
    };
  // Before a table a pending join runs words into, the block goes before
  // the join, which accepting makes one paragraph with the first cell.
  const joining =
    side === "before"
      ? firstParagraphJoiningTable({
          doc: projection.doc,
          paragraphPos: table.before,
          noteReferences,
        })
      : null;
  const joined =
    joining === null ? undefined : projection.paragraphStartingAt(joining);
  if (joined !== undefined)
    return {
      at: { type: "before", blockId: joined.blockId },
      position: joined.start,
    };
  const position = table[side];
  const walked = side === "before" ? atGapStart(position, position) : undefined;
  if (walked !== undefined) return walked;
  const resolved = projection.doc.resolve(position);
  const beside = side === "after" ? resolved.nodeAfter : resolved.nodeBefore;
  const neighbour =
    beside?.type.name === "paragraph"
      ? projection.paragraphStartingAt(
          side === "after" ? position : position - beside.nodeSize,
        )
      : undefined;
  // After the table is before the paragraph that follows it, and before it
  // is after the one it follows.
  if (neighbour !== undefined)
    return {
      at: {
        type: side === "after" ? "before" : "after",
        blockId: neighbour.blockId,
      },
      position,
    };
  const outerTable = projection.doc.nodeAt(table.before);
  const blockId =
    outerTable === null
      ? undefined
      : tableBlockAddress(projection, {
          table: outerTable,
          tableStart: table.before + 1,
        });
  if (blockId === undefined) return undefined;
  return {
    at: { type: side === "after" ? "afterTable" : "beforeTable", blockId },
    position,
  };
};

/** Where an insertion beside a paragraph goes, and whose properties it copies. */
type BesidePendingJoin = {
  place: ProjectedParagraph;
  formattingSource: ProjectedParagraph;
};

/** Beside the paragraph an insertion names as it stands, with its properties. */
const besideAsItStands = (
  paragraph: ProjectedParagraph,
): BesidePendingJoin => ({
  place: paragraph,
  formattingSource: paragraph,
});

type PlaceAfterOptions = Omit<BesidePendingJoinOptions, "side">;

/**
 * Where a paragraph put after `paragraph` goes (`insertionAfterPendingBreak`):
 * right after it, or after the paragraph its pending join leaves, in every
 * mode, so a suggestion lands where the same tracked insertion does. A
 * paragraph the batch already put between the two ends the join there, so
 * it goes right after `paragraph`, as one at a time (`joinAsTheBatchRuns`).
 */
const placeAfter = ({
  paragraph,
  projection,
  noteReferences,
  landedEarlier,
}: PlaceAfterOptions): BesidePendingJoin => {
  const rule = insertionAfterPendingBreak({
    doc: projection.doc,
    paragraphPos: paragraph.start - 1,
    noteReferences,
  });
  switch (rule.type) {
    case "afterParagraph":
      return besideAsItStands(paragraph);
    case "afterJoin": {
      const end = paragraph.start - 1 + paragraph.node.nodeSize;
      if (
        [...(landedEarlier ?? [])].some(
          (at) => at >= end && at <= rule.survivor,
        )
      )
        return besideAsItStands(paragraph);
      const survivor =
        projection.paragraphStartingAt(rule.survivor) ??
        panic("A pending join names a paragraph the projection lacks.");
      return besideAsItStands(survivor);
    }
    default: {
      const unreachable: never = rule;
      return unreachable;
    }
  }
};

type JoinAsTheBatchRunsOptions = {
  join: PendingJoinParagraphs;
  side: "before" | "after";
  /** `BesidePendingJoinOptions.landedEarlier`. */
  landed: ReadonlySet<number> | undefined;
  /** Where the paragraph the block is put beside starts. */
  paragraphPos: number;
  doc: PMNode;
};

/**
 * A pending join as the batch reads it once its earlier operations ran: a
 * paragraph one of them puts between two paragraphs of the join takes the
 * words before it, ending the join there, as it does applied one at a time.
 * A later block put before a paragraph past it reads a join starting past
 * it; one put after a paragraph before it, a join ending before it.
 */
const joinAsTheBatchRuns = ({
  join,
  side,
  landed,
  paragraphPos,
  doc,
}: JoinAsTheBatchRunsOptions): PendingJoinParagraphs => {
  if (landed === undefined || landed.size === 0) return join;
  const at = doc.resolve(join.first);
  let { first } = join;
  let position = join.first;
  for (let index = at.index(); index < at.parent.childCount; index++) {
    const next = position + at.parent.child(index).nodeSize;
    if (next > join.last || (side === "before" && position >= paragraphPos))
      break;
    if (landed.has(next) && side === "after" && position >= paragraphPos)
      return { first, last: position };
    if (landed.has(next) && side === "before") first = next;
    position = next;
  }
  return { first, last: join.last };
};

type BesidePendingJoinOptions = {
  side: "before" | "after";
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
  /**
   * The block boundaries, in the document the batch read, where its earlier
   * operations put paragraphs: one between two paragraphs of the join ends
   * the join there as the batch runs (`joinAsTheBatchRuns`).
   */
  landedEarlier?: ReadonlySet<number>;
};

/** Whether `paragraph` holds anything no pending deletion takes. */
const holdsStandingContent = (paragraph: PMNode): boolean => {
  let standing = false;
  paragraph.descendants((node) => {
    if (standing) return false;
    if (
      node.isInline &&
      !node.marks.some(({ type }) => type.name === "deletion")
    )
      standing = true;
    return !standing;
  });
  return standing;
};

/** `formatting` with no numbering of its own or from its style. */
const withoutNumbering = ({
  numPr: _numPr,
  numPrFromStyle: _numPrFromStyle,
  ...formatting
}: ParagraphFormatting): ParagraphFormatting => formatting;

/**
 * Whether accepting every pending change takes the paragraph whole: its break
 * goes (deleted, or inserted and deleted since), and every record it holds is
 * deleted or moved away.
 */
const isTakenWholeOnAcceptance = ({ pPrMark, content }: Paragraph) =>
  (pPrMark?.kind === "del" ||
    pPrMark?.kind === "moveFrom" ||
    pPrMark?.deletion !== undefined) &&
  content.every(({ type }) => type === "deletion" || type === "moveFrom");

type LentFormattingOptions = {
  formatting: ParagraphFormatting | undefined;
  source: Paragraph;
  /** The operation names `source` itself, not a paragraph a join reaches. */
  namesSource: boolean;
};

/**
 * The properties an inserted paragraph copies from `source`. A paragraph a
 * pending join reaches, which a pending deletion takes whole, lends no
 * numbering: once accepted, the list item it was is gone. A paragraph the
 * operation names lends all of them, as the default engine copies them.
 */
const lentFormatting = ({
  formatting,
  source,
  namesSource,
}: LentFormattingOptions): ParagraphFormatting | undefined =>
  formatting !== undefined && !namesSource && isTakenWholeOnAcceptance(source)
    ? withoutNumbering(formatting)
    : formatting;

/**
 * The ids of the paragraphs right before the one at `paragraphPos` whose
 * marks are pending deletion: a pending join runs them into it.
 */
const joinedBefore = (doc: PMNode, paragraphPos: number): string[] => {
  const at = doc.resolve(paragraphPos);
  const ids: string[] = [];
  for (let index = at.index() - 1; index >= 0; index--) {
    const previous = at.parent.child(index);
    const id: unknown = previous.attrs["paraId"];
    if (
      previous.type.name !== "paragraph" ||
      !isDeletedPPrMark(previous.attrs["pPrMark"]) ||
      typeof id !== "string"
    )
      break;
    ids.push(id);
  }
  return ids;
};

/** Whether a pending deletion already takes `paragraph` whole, mark and words. */
const goesWholeOnAcceptance = (paragraph: PMNode): boolean =>
  paragraph.type.name === "paragraph" &&
  isDeletedPPrMark(paragraph.attrs["pPrMark"]) &&
  !holdsStandingContent(paragraph);

/**
 * Whether the blocks after `position` in its container are paragraphs that
 * hold nothing once accepted, each before the last with its break pending
 * deletion, and no section break among them, as the docx-core block deletion
 * reads a container's end.
 */
const followersGoOnAcceptance = (doc: PMNode, position: number): boolean => {
  const at = doc.resolve(position);
  const { parent } = at;
  for (let index = at.index(); index < parent.childCount; index += 1) {
    const block = parent.child(index);
    if (
      block.type.name !== "paragraph" ||
      block.attrs["_sectionProperties"] != null ||
      holdsStandingContent(block) ||
      (index < parent.childCount - 1 && !goesWholeOnAcceptance(block))
    )
      return false;
  }
  return at.index() < parent.childCount;
};

/**
 * Where an insertion beside `paragraph` goes, read as the joined paragraph a
 * pending join makes of it in the document the batch read
 * (`pendingJoinParagraphs`, which the legacy engine places by too): after
 * the join's last paragraph, or before its first, copying the properties of
 * the paragraph the join leaves, which the joined paragraph has. The batch's
 * other operations do not move it: a split of the paragraph the join leaves
 * puts the block after both halves, as one put after that paragraph goes.
 * Where no join reads the paragraph, `pendingJoinParagraphs` names the
 * paragraph itself on that side, so the block goes beside it: one place
 * decides, so a paragraph going whole that earlier words run into is placed
 * before those words, as every other paragraph of the join is.
 */
const besidePendingJoin = ({
  side,
  paragraph,
  projection,
  noteReferences,
  landedEarlier,
}: BesidePendingJoinOptions): BesidePendingJoin => {
  const options = {
    doc: projection.doc,
    paragraphPos: paragraph.start - 1,
    noteReferences,
  };
  const join = joinAsTheBatchRuns({
    join: pendingJoinParagraphs(options),
    side,
    landed: landedEarlier,
    paragraphPos: options.paragraphPos,
    doc: projection.doc,
  });
  const after = side === "after";
  const at = (position: number) =>
    position === options.paragraphPos
      ? paragraph
      : projection.paragraphStartingAt(position);
  const place = at(after ? join.last : join.first);
  const formattingSource = at(join.last);
  if (place === undefined || formattingSource === undefined)
    return panic("A pending join names a paragraph the projection lacks.");
  return { place, formattingSource };
};

/**
 * The paragraph outside tables a block put beside the one `operation` names
 * lands beside, as the batch read it: after a paragraph, the one a pending
 * join leaves where the join reads the two as one (`placeAfter`,
 * `besidePendingJoin`); before one, the first the join takes in. A batch
 * reads what its blocks stand beside by it, as they land.
 */
const landingParagraph = (
  operation: BlockInsertion | TableInsertion,
  projection: CanonicalProjection,
  noteReferences: NoteReferenceLabels,
): ProjectedParagraph | undefined => {
  const paragraph = projection.paragraph(operation.blockId);
  if (
    paragraph === undefined ||
    findOutermostTableBoundary(projection.doc, paragraph.start - 1) !== null
  )
    return undefined;
  const options = { paragraph, projection, noteReferences };
  const side = placementSide(operation);
  return (
    operation.type === "insertAfterBlock"
      ? placeAfter(options)
      : besidePendingJoin({ ...options, side })
  ).place;
};

/**
 * One new paragraph beside the anchor, as the legacy engine writes it: the
 * anchor's stated paragraph properties unless `inheritFormatting` is false,
 * with the operation's own on top, holding the text as one plain run.
 */
const resolveBlockInsertion = ({
  operation,
  listStart,
  paragraph,
  named,
  formattingSource,
  projection,
  noteReferences,
  story,
  mintList,
  refusal,
  skipAlike,
  landedEarlier,
}: ResolveBlockInsertionOptions) => {
  const { text } = operation;
  const markdown =
    operation.textFormat === "markdown"
      ? compileMarkdownInsertion(text)
      : undefined;
  if (markdown?.type === "unsupported")
    return skipAlike("unsupportedBlock", markdown.message);
  const requested =
    markdown?.listKind === undefined
      ? operation.numbering
      : ({ kind: "newList", format: markdown.listKind } as const);
  const stated = statedNumbering(requested, mintList);
  const inline = operation.lineBreakMode === "inline";
  // A comparison writes the references its paragraph brings as their
  // markers, which it turns back into references once the batch lands. Any
  // other insertion has no note to give a marker it writes: the operations
  // create no notes, so the marker would read as a reference the document
  // does not hold.
  if (
    markdown === undefined &&
    !noteReferenceEditsAllowed() &&
    /\[\^(?:e)?\d+\]/u.test(text)
  )
    return refusal(
      "protectedReference",
      "An insertion writes no note marker of its own.",
    );
  // Each line its own paragraph, blank lines dropped, as the legacy engine
  // splits them; inline, line breaks stay in the one paragraph. A tab is a
  // tab element either way.
  const split =
    markdown === undefined && !inline && LINE_BREAK_PATTERN.test(text);
  const lines = split
    ? text.split(LINE_BREAK_PATTERN).filter((line) => line.trim().length > 0)
    : [text];
  const texts = lines.length === 0 ? [""] : lines;
  const ownPatch = Object.entries(
    statedParagraphPatch({
      ...(operation.styleId === undefined
        ? {}
        : { styleId: operation.styleId }),
      ...(operation.outlineLevel === undefined
        ? {}
        : { outlineLevel: operation.outlineLevel }),
      ...(operation.alignment === undefined
        ? {}
        : { alignment: operation.alignment }),
      ...(operation.spacing === undefined
        ? {}
        : { spacing: operation.spacing }),
      ...(operation.indentation === undefined
        ? {}
        : { indentation: operation.indentation }),
      ...(stated.numbering === undefined
        ? {}
        : { numbering: stated.numbering }),
    }),
  );
  const everyParagraph = operation.formattingScope === "allParagraphs";
  const markdownNumbering =
    stated.numbering?.kind === "reference" ? stated.numbering : undefined;
  let markdownParagraphs: Paragraph[] | undefined;
  if (markdown !== undefined) {
    markdownParagraphs = [];
    for (const {
      listRendering: _listRendering,
      ...item
    } of markdown.paragraphs) {
      if (markdownNumbering !== undefined)
        item.formatting = { ...item.formatting, numPr: markdownNumbering };
      markdownParagraphs.push(item);
    }
  }
  // The source's properties with the operation's set or cleared over them.
  const paragraphsFrom = (
    source: ParagraphFormatting | undefined,
  ): Paragraph[] => {
    const formatting: Record<string, unknown> =
      operation.inheritFormatting === false ? {} : { ...source };
    for (const [key, value] of ownPatch)
      if (value === null) Reflect.deleteProperty(formatting, key);
      else formatting[key] = value;
    return texts.map((line, index): Paragraph => {
      // The first paragraph takes the operation's formatting; later ones are
      // body paragraphs unless the scope formats every paragraph alike.
      const own: Record<string, unknown> =
        index === 0 || everyParagraph ? { ...formatting } : {};
      if (index === 0 && operation.pageBreakBefore === true)
        own["pageBreakBefore"] = true;
      // An authored page break stands in for the first paragraph's text.
      const pageBreak = index === 0 ? operation.hardPageBreak : undefined;
      let content: Paragraph["content"] = [];
      if (pageBreak !== undefined)
        content = [
          {
            type: "run",
            content: [
              {
                type: "break",
                breakType: "page",
                ...(pageBreak.clear === undefined
                  ? {}
                  : { clear: pageBreak.clear }),
              },
            ],
          },
        ];
      else if (line !== "") content = emphasisRuns(line);
      return {
        type: "paragraph",
        content,
        // The anchor's properties with the operation's set or cleared over them.
        ...(Object.keys(own).length === 0
          ? {}
          : { formatting: own as ParagraphFormatting }),
      };
    });
  };
  const sourceFormatting = paragraphFormattingWithAuthoredIndentation(
    formattingSource.source,
  );
  const namesSource = formattingSource.blockId === named.blockId;
  const paragraphs =
    markdownParagraphs ??
    paragraphsFrom(
      lentFormatting({
        formatting: sourceFormatting,
        source: formattingSource.source,
        namesSource,
      }),
    );
  const place =
    listStart === undefined
      ? blockInsertionPlace(
          operation,
          paragraph,
          projection,
          noteReferences,
          landedEarlier,
        )
      : {
          at: { type: "containerStart", blockId: listStart.blockId } as const,
          position: listStart.position,
        };
  if (place === undefined)
    return refusal(
      "unsupportedBlock",
      "Canonical insertion beside a table needs a paragraph on that side of it.",
      INVALID_REQUEST,
    );
  // The anchor keeps its identity; the new paragraph takes a fresh one.
  const intents: EditorIntent[] = [
    {
      type: "insertParagraphs",
      story,
      at: place.at,
      paragraphs,
      ...(stated.newList === undefined ? {} : { newList: stated.newList }),
    },
  ];
  const { position } = place;
  // A comment covers the inserted text, so an insertion with no text has
  // nothing for it to cover, as the legacy engine finds.
  if (
    operation.comment !== undefined &&
    texts.every((line) => line.length === 0)
  )
    return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
  const normalizations: FolioAIEditNormalization[] = split
    ? [
        {
          id: operation.id,
          code: "splitMultilineText",
          paragraphCount: texts.length,
        },
      ]
    : [];
  const resolved = {
    intents,
    normalizations,
    from: position,
    to: position,
    ...(markdownParagraphs === undefined &&
    operation.inheritFormatting !== false
      ? {
          inheritsFrom: {
            blockId: formattingSource.blockId,
            markDeleted: isDeletedPPrMark(
              formattingSource.node.attrs["pPrMark"],
            ),
            followsAnchor:
              operation.type === "insertAfterBlock" &&
              formattingSource.blockId === paragraph.blockId,
            namesSource,
            paragraphs: paragraphsFrom,
          },
        }
      : {}),
  };
  if (operation.comment === undefined)
    return Result.ok({ type: "plain" as const, operation, ...resolved });
  // Where the inserted paragraphs are is known once they are compiled; the
  // place they go in, at the edge of the paragraph they go beside, is the
  // point a comment marks where they hold no paragraph a range can name.
  const anchor = { story, blockId: place.at.blockId, offset: 0 };
  const beside =
    place.at.type === "after"
      ? projection.paragraph(place.at.blockId)
      : undefined;
  return Result.ok({
    type: "commented" as const,
    operation,
    comment: {
      text: operation.comment.text,
      from: anchor,
      to:
        beside === undefined
          ? anchor
          : { ...anchor, offset: paragraphLength(beside.source) },
      inserted: true as const,
    },
    ...resolved,
  });
};

/**
 * A tracked deletion of the paragraph that ends its block list, retiring the
 * break of the paragraph before it in `document`, the batch as it stands
 * when the deletion runs: an insertion the order ran first, right before the
 * deleted paragraph, is the paragraph whose break it retires, as deleting it
 * after that insertion does.
 */
const retiredBreakAsItRuns = (
  document: Document,
  story: OpStory,
  intents: readonly EditorIntent[],
): readonly EditorIntent[] => {
  return intents.map((intent) => {
    if (intent.type !== "replaceText") return intent;
    // A text box's paragraphs are named through the story drawing it.
    const here = locateParagraphThroughTextBoxes(
      document,
      story,
      intent.to.blockId,
    );
    const before = here?.blocks[here.index - 1];
    if (
      before?.type !== "paragraph" ||
      before.paraId === undefined ||
      idKey(before.paraId) === idKey(intent.from.blockId)
    )
      return intent;
    return {
      ...intent,
      from: {
        story: intent.from.story,
        blockId: before.paraId,
        offset: paragraphLength(before),
      },
    };
  });
};

/**
 * A direct deletion's join with the paragraph before it, bound to the
 * paragraph before it in `document`, the batch as it stands when the
 * deletion runs: an insertion the order ran first, right before the deleted
 * paragraph, is the paragraph it joins, as deleting it after that insertion
 * joins it.
 */
const deletionIntentsAsItRuns = (
  document: Document,
  story: OpStory,
  intents: readonly EditorIntent[],
): readonly EditorIntent[] => {
  return intents.map((intent) => {
    if (intent.type !== "joinParagraphs") return intent;
    // A text box's paragraphs are named through the story drawing it.
    const here = locateParagraphThroughTextBoxes(
      document,
      story,
      intent.nextBlockId,
    );
    const before = here?.blocks[here.index - 1];
    if (
      before?.type !== "paragraph" ||
      before.paraId === undefined ||
      idKey(before.paraId) === idKey(intent.blockId)
    )
      return intent;
    return { ...intent, blockId: before.paraId };
  });
};

/**
 * A direct paragraph deletion as it runs in `document`: a paragraph the
 * batch's earlier operations left last in its container is deleted as a
 * final one is. It loses its words and joins the paragraph before it, which
 * keeps its identity and properties; with none before it, it only loses its
 * words, as the container keeps a paragraph.
 */
const finalDeletionIntentsAsItRuns = (
  document: Document,
  story: OpStory,
  intents: readonly EditorIntent[],
): readonly EditorIntent[] => {
  const [deletion, ...rest] = intents;
  if (
    deletion?.type !== "deleteParagraphs" ||
    deletion.move === true ||
    rest.length > 0 ||
    deletion.blockIds.length !== 1
  )
    return intents;
  const blockId = deletion.blockIds.at(0);
  if (blockId === undefined) return intents;
  // A text box's paragraphs are named through the story drawing it.
  const here = locateParagraphThroughTextBoxes(document, story, blockId);
  if (here === undefined || here.index !== here.blocks.length - 1)
    return intents;
  const length = paragraphLength(here.paragraph);
  const cleared: EditorIntent[] =
    length === 0
      ? []
      : [
          {
            type: "replaceText",
            from: { story, blockId, offset: 0 },
            to: { story, blockId, offset: length },
            text: "",
          },
        ];
  const before = here.blocks[here.index - 1];
  if (before?.type !== "paragraph" || before.paraId === undefined)
    return cleared;
  return [
    ...cleared,
    {
      type: "joinParagraphs",
      story,
      blockId: before.paraId,
      nextBlockId: blockId,
      survivor: SPLIT_HALVES.FIRST,
    },
  ];
};

type FormattingIntentsAsItRunsOptions = {
  document: Document;
  story: OpStory;
  /** Each paragraph the batch joined away, to the one it now lives in. */
  survivorOf: ReadonlyMap<string, string>;
  /** Those of them whose words open the paragraph they now live in. */
  openingSurvivorOf: ReadonlyMap<string, string>;
  /**
   * Each paragraph a block deletion of the batch absorbed, such as the one
   * that takes the final mark of a deleted run ending its container, to the
   * one it now lives in. A join's absorbed half is not here: an insertion
   * beside it keeps its own place.
   */
  deletionSurvivorOf: ReadonlyMap<string, string>;
  intents: readonly EditorIntent[];
};

/**
 * Paragraph property edits, tables and paragraph insertions beside a
 * paragraph, bound to the paragraph as the batch stands when they run: one
 * an earlier operation joined away, such as the first half of a split whose
 * retraction joins it to the second, a paragraph a direct deletion of the
 * paragraphs after it joined into the last one, or a paragraph that took the
 * mark of the run a block deletion of the batch deleted after it, is edited,
 * or has its table or insertion beside it, in the paragraph that now holds
 * it; a join's absorbed half keeps its own place for an insertion, except
 * that a paragraph put before it goes before the paragraph its words open,
 * as a merge into the next paragraph leaves them. A table after it then
 * writes the empty paragraph the container ends with when that paragraph
 * ends its block list.
 */
const formattingIntentsAsItRuns = ({
  document,
  story,
  survivorOf,
  openingSurvivorOf,
  deletionSurvivorOf,
  intents,
}: FormattingIntentsAsItRunsOptions): readonly EditorIntent[] => {
  if (survivorOf.size === 0) return intents;
  // A text box's paragraphs are named through the story drawing it.
  const present = new Set(
    paragraphsThroughTextBoxes(document, story).map(({ paragraph }) =>
      idKey(paragraph.paraId ?? ""),
    ),
  );
  // The paragraph's start, in the paragraph that now holds it.
  const bound = (
    at: TextPosition,
    survivors: ReadonlyMap<string, string> = survivorOf,
  ): TextPosition => {
    const survivor = survivors.get(idKey(at.blockId));
    return survivor === undefined || present.has(idKey(at.blockId))
      ? at
      : { story: at.story, blockId: survivor, offset: 0 };
  };
  return intents.map((intent) => {
    switch (intent.type) {
      case "tableBlock": {
        const { operation } = intent;
        if (
          operation.type !== DOCUMENT_OP_TYPES.INSERT_TABLE ||
          operation.anchor !== "paragraph"
        )
          return intent;
        const { blockId } = bound({
          story: operation.story,
          blockId: operation.blockId,
          offset: 0,
        });
        if (blockId === operation.blockId) return intent;
        const here = locateParagraphThroughTextBoxes(document, story, blockId);
        const ends =
          here !== undefined && here.index === here.blocks.length - 1;
        return {
          ...intent,
          operation: {
            ...operation,
            blockId,
            ...(operation.side === "after" && ends
              ? { paragraphAfter: "empty" as const }
              : {}),
          },
        };
      }
      case "formatParagraph":
        return { ...intent, at: bound(intent.at) };
      case "setList":
        return {
          ...intent,
          items: intent.items.map((item) => ({ ...item, at: bound(item.at) })),
        };
      case "insertParagraphs": {
        const { at } = intent;
        if (at.type !== "before" && at.type !== "after") return intent;
        if (present.has(idKey(at.blockId))) return intent;
        const opening =
          at.type === "before"
            ? openingSurvivorOf.get(idKey(at.blockId))
            : undefined;
        if (opening !== undefined)
          return { ...intent, at: { type: "before", blockId: opening } };
        const { blockId } = bound(
          { story, blockId: at.blockId, offset: 0 },
          deletionSurvivorOf,
        );
        return blockId === at.blockId
          ? intent
          : { ...intent, at: { type: at.type, blockId } };
      }
      default:
        return intent;
    }
  });
};

type OpensSurvivorOptions = {
  /** The document before the operation that joined the paragraph away. */
  document: Document;
  story: OpStory;
  absorption: ParagraphAbsorption;
};

/**
 * Whether the absorbed paragraph's words open the survivor: it stood
 * before it in the same block list.
 */
const opensSurvivor = ({
  document,
  story,
  absorption: { absorbed, survivor },
}: OpensSurvivorOptions): boolean => {
  const gone = locateParagraphThroughTextBoxes(document, story, absorbed);
  const kept = locateParagraphThroughTextBoxes(document, story, survivor);
  return (
    gone !== undefined &&
    kept !== undefined &&
    gone.blocks === kept.blocks &&
    gone.index < kept.index
  );
};

type AddedMarkBeforeTableOptions = {
  document: Document;
  story: OpStory;
  blockId: string;
  author: string;
  /** Whether the batch puts a block right after the paragraph `id`. */
  gainsFollower: (id: string) => boolean;
  intents: readonly EditorIntent[];
};

/**
 * A tracked deletion of a body paragraph whose break is another owner's
 * pending insertion, right before a table, or `author`'s own right before a
 * table whose rows are all pending insertion, as the batch stands when it
 * runs: the block deletion records the deletion beside that insertion and
 * takes the words, so accepting removes the paragraph as the direct deletion
 * does. Read as it runs, it is the same whether that table, or that
 * insertion, came before the batch or from it.
 */
const addedMarkBeforeTableAsItRuns = ({
  document,
  story,
  blockId,
  author,
  gainsFollower,
  intents,
}: AddedMarkBeforeTableOptions): readonly EditorIntent[] => {
  const blocks = findStoryBody(document, story)?.content ?? [];
  const index = blocks.findIndex(
    (block) =>
      block.type === "paragraph" &&
      idKey(block.paraId ?? "") === idKey(blockId),
  );
  const paragraph = blocks[index];
  const following = blocks[index + 1];
  const previous = blocks[index - 1];
  // A paragraph before it whose mark goes would run into the table once both
  // are accepted: the deletion then takes the words only. A block the batch
  // puts after that paragraph stands between them instead.
  const previousJoins =
    previous?.type === "paragraph" &&
    !gainsFollower(previous.paraId ?? "") &&
    (previous.pPrMark?.kind === "del" ||
      previous.pPrMark?.kind === "moveFrom" ||
      previous.pPrMark?.deletion !== undefined);
  if (
    previousJoins ||
    paragraph?.type !== "paragraph" ||
    paragraph.paraId === undefined ||
    paragraph.pPrMark?.kind !== "ins" ||
    paragraph.pPrMark.deletion !== undefined ||
    following?.type !== "table" ||
    following.rows.length === 0 ||
    (paragraph.pPrMark.info.author === author &&
      !following.rows.every(
        (row) => row.structuralChange?.type === "tableRowInsertion",
      ))
  )
    return intents;
  return [{ type: "deleteParagraphs", story, blockIds: [paragraph.paraId] }];
};

type TakenOverBreakRetractionOptions = {
  document: Document;
  story: OpStory;
  blockId: string;
  author: string;
  suggestionId: string | undefined;
  intents: readonly EditorIntent[];
};

/**
 * A tracked deletion of the body's last paragraph, as the batch stands when
 * it runs, that takes back this owner's (`author`, and `suggestionId` in a
 * suggestion) insertion after the paragraph before it: that paragraph holds
 * the insertion's new break (`survivor: "first"`) and the last one only its
 * words, and the anchors of comments on them, under the former mark it took
 * over. The block deletion rejects that insertion, so the paragraph before
 * keeps its identity and gets its mark and properties back, and the comments
 * it leaves nothing of retire; a text range over the break would retire its
 * identity.
 */
const takenOverBreakRetractionAsItRuns = ({
  document,
  story,
  blockId,
  author,
  suggestionId,
  intents,
}: TakenOverBreakRetractionOptions): readonly EditorIntent[] => {
  const blocks = findStoryBody(document, story)?.content ?? [];
  const paragraph = blocks.at(-1);
  const previous = blocks.at(-2);
  const owns = (info: {
    author?: string | undefined;
    suggestionId?: string | undefined;
  }) => info.author === author && info.suggestionId === suggestionId;
  if (
    paragraph?.type !== "paragraph" ||
    paragraph.paraId === undefined ||
    idKey(paragraph.paraId) !== idKey(blockId) ||
    !paragraph.content.every(
      (item) =>
        (item.type === "insertion" && owns(item.info)) ||
        isCommentAnchorCarrierType(item.type),
    ) ||
    previous?.type !== "paragraph" ||
    previous.pPrMark?.kind !== "ins" ||
    previous.pPrMark.deletion !== undefined ||
    previous.pPrMark.survivor !== "first" ||
    !owns(previous.pPrMark.info)
  )
    return intents;
  return [{ type: "deleteParagraphs", story, blockIds: [paragraph.paraId] }];
};

/**
 * The comment threads `before` holds that `after` no longer does: a tracked
 * operation that takes back its own insertion retires the comments it leaves
 * nothing of outright, as a direct deletion does.
 */
const commentsRetiredBetween = (
  before: Document,
  after: Document,
): number[] => {
  const kept = new Set(
    (after.package.document.comments ?? []).map(({ id }) => id),
  );
  return (before.package.document.comments ?? []).flatMap(({ id, parentId }) =>
    parentId === undefined && !kept.has(id) ? [id] : [],
  );
};

/** The node after `$pos` past any block-level range boundaries, if any. */
const nodePastBoundaries = ($pos: ResolvedPos): PMNode | null => {
  for (let index = $pos.index(); index < $pos.parent.childCount; index += 1) {
    const node = $pos.parent.child(index);
    if (!isBlockBoundaryNode(node)) return node;
  }
  return null;
};

/**
 * Whether only paragraphs accepting removes whole stand in `blocks`, one of
 * them at least pending removal: a join before them, once accepted, has no
 * words after it to separate. Range boundaries among them hold nothing.
 */
const onlyRemovedParagraphsFollow = (
  blocks: readonly BlockContent[],
): boolean => {
  let removes = false;
  for (const block of blocks) {
    if (isBlockBoundary(block)) continue;
    if (block.type !== "paragraph" || paragraphHoldsContentOnceAccepted(block))
      return false;
    removes ||= paragraphHoldsContent(block) || block.pPrMark?.kind === "del";
  }
  return removes;
};

type MergeIntentsAsItRunsOptions = {
  document: Document;
  story: OpStory;
  intents: readonly EditorIntent[];
  mode: FolioDocumentOperationBatch["mode"];
};

/**
 * A merge's intents as the batch stands in `document` when the merge runs.
 * The separator goes in at the first paragraph's end as it now stands: an
 * edit of that paragraph the batch ran first, a comment's reference among
 * them, may have changed its length. Where only paragraphs accepting removes
 * whole follow it, deletions the batch made first among them, it goes in not
 * at all: accepting leaves no words after it to separate, as the legacy
 * engine reads the document as the merge runs. A deletion the batch refuses,
 * or makes after the merge, has not taken them. Direct, the join also binds
 * to the paragraph that now follows: a paragraph the batch deleted first is no
 * longer there, so the merge joins across it, as merging after that deletion
 * does; `undefined` when no paragraph follows any more. Tracked, a deleted
 * paragraph stays, and the join keeps it, unless an earlier operation joined
 * that paragraph away.
 */
const mergeIntentsAsItRuns = ({
  document,
  story,
  intents,
  mode,
}: MergeIntentsAsItRunsOptions): readonly EditorIntent[] | undefined => {
  const join = intents.find(
    (intent): intent is Extract<EditorIntent, { type: "joinParagraphs" }> =>
      intent.type === "joinParagraphs",
  );
  if (join === undefined || findStoryBody(document, story) === undefined)
    return intents;
  // A text box's paragraphs are named through the story drawing it.
  const locations = paragraphsThroughTextBoxes(document, story);
  const key = idKey(join.blockId);
  const at = locations.findIndex(
    ({ paragraph }) => idKey(paragraph.paraId ?? "") === key,
  );
  const here = locations[at];
  // Tracked, the join follows its paragraph too once an earlier operation
  // took the one it named, as retracting a split joins the split's halves
  // under the second's id.
  const was = idKey(join.nextBlockId);
  const bindsToNext =
    mode === "direct" ||
    !locations.some(({ paragraph }) => idKey(paragraph.paraId ?? "") === was);
  if (here === undefined) return bindsToNext ? undefined : intents;
  const end = paragraphLength(here.paragraph);
  const siblings = locateParagraphThroughTextBoxes(
    document,
    story,
    join.blockId,
  )?.blocks;
  const separates =
    siblings === undefined ||
    !onlyRemovedParagraphsFollow(siblings.slice(here.index + 1));
  const separatorAtEnd: EditorIntent[] = [];
  for (const intent of intents) {
    if (
      intent.type !== "replaceText" ||
      idKey(intent.from.blockId) !== key ||
      intent.from.offset !== intent.to.offset
    ) {
      separatorAtEnd.push(intent);
      continue;
    }
    if (separates)
      separatorAtEnd.push({
        ...intent,
        from: { ...intent.from, offset: end },
        to: { ...intent.to, offset: end },
      });
  }
  if (!bindsToNext) return separatorAtEnd;
  const next = locations[at + 1];
  // Range boundaries may stand between the two, where no section ends: the
  // join crosses them as resolving a mark does.
  const between = siblings?.slice(
    here.index + 1,
    next?.index ?? here.index + 1,
  );
  const follows =
    next !== undefined &&
    JSON.stringify(here.list) === JSON.stringify(next.list) &&
    (here.index + 1 === next.index ||
      (here.paragraph.sectionProperties === undefined &&
        between !== undefined &&
        between.length === next.index - here.index - 1 &&
        between.every(isBlockBoundary)));
  const nextId = follows ? next.paragraph.paraId : undefined;
  if (nextId === undefined) return undefined;
  const boundToNext = (intent: EditorIntent): EditorIntent => {
    switch (intent.type) {
      case "joinParagraphs":
        return idKey(intent.nextBlockId) === was
          ? { ...intent, nextBlockId: nextId }
          : intent;
      case "formatParagraph":
        return idKey(intent.at.blockId) === was
          ? { ...intent, at: { ...intent.at, blockId: nextId } }
          : intent;
      case "setList":
        return {
          ...intent,
          items: intent.items.map((item) =>
            idKey(item.at.blockId) === was
              ? { ...item, at: { ...item.at, blockId: nextId } }
              : item,
          ),
        };
      default:
        return intent;
    }
  };
  return separatorAtEnd.map(boundToNext);
};

type InsertionIntentsAsItRunsOptions = {
  document: Document;
  story: OpStory;
  /** Each paragraph the batch joined away, to the one it now lives in. */
  survivorOf: ReadonlyMap<string, string>;
  resolvedOperation: PreparedOperation;
};

/**
 * A block insertion's intents with the properties it copies read from
 * `document`, the batch as it stands when the insertion runs: an edit of
 * the copied paragraph the batch order runs first is in the copy, as the
 * legacy engine copies it, and a paragraph the batch joined away is copied
 * as the paragraph that now holds it.
 */
const insertionIntentsAsItRuns = ({
  document,
  story,
  survivorOf,
  resolvedOperation,
}: InsertionIntentsAsItRunsOptions): readonly EditorIntent[] => {
  const { intents, inheritsFrom } = resolvedOperation;
  if (inheritsFrom === undefined) return intents;
  // A text box's paragraphs are named through the story drawing it.
  const locate = (blockId: string) =>
    locateParagraphThroughTextBoxes(document, story, blockId);
  const survivor = survivorOf.get(idKey(inheritsFrom.blockId));
  const at =
    locate(inheritsFrom.blockId) ??
    (survivor === undefined ? undefined : locate(survivor));
  if (at === undefined) return intents;
  let source = at.paragraph;
  // A pending deletion takes its formatting from the surviving paragraph in
  // this block list. A table or wrapper ends the chain.
  if (!inheritsFrom.markDeleted && !inheritsFrom.followsAnchor) {
    let index = at.index;
    while (isDeletedPPrMark(source.pPrMark)) {
      const next = at.blocks[++index];
      if (next?.type !== "paragraph") break;
      source = next;
    }
  }
  const lent = lentFormatting({
    formatting: paragraphFormattingWithAuthoredIndentation(source),
    source,
    namesSource: inheritsFrom.namesSource && source === at.paragraph,
  });
  const copied: EditorIntent[] = [];
  for (const intent of intents)
    copied.push(
      intent.type === "insertParagraphs"
        ? Object.assign({}, intent, {
            paragraphs: inheritsFrom.paragraphs(lent),
          })
        : intent,
    );
  return copied;
};

type InsertionAnchor =
  | { type: "paragraph"; at: BlockInsertionPoint }
  | { type: "table"; at: BlockInsertionPoint }
  /**
   * The start of the block list holding the paragraph `at` names. Each
   * block put there goes ahead of those the batch put there before it, as
   * blocks put after one paragraph do, so it reads as `after` that start.
   */
  | { type: "containerStart"; at: { type: "after"; blockId: string } };

/**
 * The anchor a group of insertions after a paragraph replays against: the
 * paragraph after it, when one follows, so each block goes before it in
 * batch order. Placed after the anchor itself, a later block would find an
 * earlier one between the anchor and that paragraph, and a tracked paragraph
 * put after an anchor whose mark is under review, with a table and no
 * paragraph after it, has no mark it can record its insertion on.
 */
type GapAnchorBeforeFollowerOptions = {
  anchor: InsertionAnchor;
  projection: CanonicalProjection;
  /** Paragraphs the batch removes outright before its insertions replay. */
  removedBlockIds: ReadonlySet<string>;
};

/**
 * An insertion after a paragraph, as one before the paragraph that follows
 * it. With no paragraph after it, or one the batch removes, which is gone by
 * the time the insertion replays, the anchor is kept: the blocks the batch
 * puts before the next paragraph it keeps come after these.
 */
const gapAnchorBeforeFollower = ({
  anchor,
  projection,
  removedBlockIds,
}: GapAnchorBeforeFollowerOptions): InsertionAnchor => {
  if (anchor.type !== "paragraph" || anchor.at.type !== "after") return anchor;
  const paragraph = projection.paragraph(anchor.at.blockId);
  if (paragraph === undefined)
    return panic("An insertion anchor is missing from the projection.");
  const position = paragraph.start - 1 + paragraph.node.nodeSize;
  const next = projection.doc.resolve(position).nodeAfter;
  // Before a table, after an anchor whose break is pending deletion, the
  // blocks replay before the table: a paragraph after the anchor would take
  // over its mark once another block of the gap stands there.
  if (
    next?.type.name === "table" &&
    isDeletedPPrMark(paragraph.node.attrs["pPrMark"])
  ) {
    const blockId = tableBlockAddress(projection, {
      table: next,
      tableStart: position + 1,
    });
    return blockId === undefined
      ? anchor
      : { type: "table", at: { type: "before", blockId } };
  }
  if (next?.type.name !== "paragraph") return anchor;
  const follower =
    projection.paragraphStartingAt(position) ??
    panic("A following paragraph is missing from the projection.");
  return removedBlockIds.has(follower.blockId)
    ? anchor
    : { type: "paragraph", at: { type: "before", blockId: follower.blockId } };
};

type InsertionGapOptions = {
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
  projection: CanonicalProjection;
  /** `BesidePendingJoinOptions.landedEarlier`, for the gap of a block put before. */
  landedEarlier?: ReadonlySet<number>;
};

/** The physical anchor and accepted source gap, before this batch inserts. */
const canonicalInsertionGap = ({
  landedEarlier,
  operation,
  intents,
  projection,
}: InsertionGapOptions) => {
  const intent = intents.at(0);
  let anchor: InsertionAnchor;
  if (intent?.type === "insertParagraphs") {
    switch (intent.at.type) {
      case "before":
      case "after":
        anchor = { type: "paragraph", at: intent.at };
        break;
      case "beforeTable":
      case "afterTable":
        anchor = {
          type: "table",
          at: {
            type: intent.at.type === "beforeTable" ? "before" : "after",
            blockId: intent.at.blockId,
          },
        };
        break;
      case "containerStart":
        anchor = {
          type: "containerStart",
          at: { type: "after", blockId: intent.at.blockId },
        };
        break;
      default: {
        const unreachable: never = intent.at;
        return unreachable;
      }
    }
  } else if (
    intent?.type === "tableBlock" &&
    intent.operation.type === DOCUMENT_OP_TYPES.INSERT_TABLE
  ) {
    const { anchor: type, side, blockId } = intent.operation;
    if (type === "containerStart")
      anchor = { type, at: { type: "after", blockId } };
    else if (type === undefined || side === undefined)
      return panic(
        "A resolved table insertion needs an explicit anchor and side.",
      );
    else anchor = { type, at: { type: side, blockId } };
  } else return undefined;
  const paragraph = projection.paragraph(anchor.at.blockId);
  if (paragraph === undefined)
    panic("A resolved insertion lost its source anchor.");
  const outer =
    anchor.type === "table"
      ? findOutermostTableBoundary(projection.doc, paragraph.start - 1)
      : null;
  if (anchor.type === "table" && outer === null)
    panic("A resolved table insertion lost its source table.");
  let boundary: number;
  if (anchor.type === "containerStart")
    boundary = projection.doc.resolve(paragraph.start - 1).start();
  else
    boundary =
      outer === null
        ? paragraph.start -
          1 +
          (anchor.at.type === "after" ? paragraph.node.nodeSize : 0)
        : outer[anchor.at.type];
  if (!placesBlock(operation)) return undefined;
  const source = projection.paragraph(operation.blockId);
  if (source === undefined)
    panic("A resolved insertion lost its requested source block.");
  const sourceTable = findOutermostTableBoundary(
    projection.doc,
    source.start - 1,
  );
  const blockSide = operation.type === "insertBeforeBlock" ? "before" : "after";
  const gapOf = (at: number, side: "before" | "after") =>
    side === "before"
      ? gapStartAsTheBatchRuns(projection.doc, at, landedEarlier)
      : insertionGap(projection.doc, at, side);
  const requestedSide = isBlockInsertion(operation)
    ? blockSide
    : (operation.position ?? "after");
  return {
    anchor,
    boundary,
    // Lowering an insertion after a table to before its next paragraph
    // changes its physical anchor, not its accepted source gap. Before a
    // block, the gap stops where the placement does: at a paragraph an
    // earlier operation puts among the blocks accepting removes
    // (`gapStartAsTheBatchRuns`).
    gap: gapOf(
      sourceTable === null ? boundary : sourceTable[requestedSide],
      sourceTable === null ? anchor.at.type : requestedSide,
    ),
  };
};

type PlaceInsertionIntentsOptions = {
  intents: readonly EditorIntent[];
  anchor: InsertionAnchor;
  placement: "first" | "following";
};

/** Placement changes; authored formatting, lists and operation ownership stay put. */
const placeInsertionIntents = ({
  intents,
  anchor,
  placement,
}: PlaceInsertionIntentsOptions): readonly EditorIntent[] =>
  intents.map((intent) => {
    if (intent.type === "insertParagraphs") {
      switch (anchor.type) {
        case "paragraph":
          return { ...intent, at: anchor.at };
        case "table":
          return {
            ...intent,
            at: {
              type: anchor.at.type === "before" ? "beforeTable" : "afterTable",
              blockId: anchor.at.blockId,
            },
          };
        case "containerStart":
          return {
            ...intent,
            at: { type: "containerStart", blockId: anchor.at.blockId },
          };
        default: {
          const unreachable: never = anchor;
          return unreachable;
        }
      }
    }
    if (
      intent.type !== "tableBlock" ||
      intent.operation.type !== DOCUMENT_OP_TYPES.INSERT_TABLE
    )
      return intent;
    const { operation } = intent;
    if (anchor.type === "containerStart")
      return {
        type: "tableBlock",
        operation: {
          type: operation.type,
          story: operation.story,
          table: operation.table,
          blockId: anchor.at.blockId,
          anchor: "containerStart",
        },
      };
    if (anchor.type === "table")
      return {
        type: "tableBlock",
        operation: {
          type: operation.type,
          story: operation.story,
          table: operation.table,
          blockId: anchor.at.blockId,
          anchor: "table",
          side: anchor.at.type,
        },
      };
    return {
      type: "tableBlock",
      operation: {
        type: operation.type,
        story: operation.story,
        table: operation.table,
        blockId: anchor.at.blockId,
        anchor: "paragraph",
        side: anchor.at.type,
        ...(operation.anchor === "paragraph" &&
        operation.separated !== undefined
          ? { separated: operation.separated }
          : {}),
        paragraphAfter:
          placement === "first" && operation.anchor === "paragraph"
            ? (operation.paragraphAfter ?? "empty")
            : "none",
      },
    };
  });

type DroppedInsertionCarrierOptions = {
  admitted: readonly EditorIntent[];
  replayed: readonly EditorIntent[];
  ops: readonly DocumentOp[];
};

/** The table compiler emits its leading/trailing carriers before its table. */
const droppedInsertionCarrier = ({
  admitted,
  replayed,
  ops,
}: DroppedInsertionCarrierOptions) => {
  const before = admitted.at(0),
    after = replayed.at(0);
  if (
    before?.type !== "tableBlock" ||
    before.operation.type !== DOCUMENT_OP_TYPES.INSERT_TABLE ||
    before.operation.anchor !== "paragraph" ||
    before.operation.paragraphAfter === "none" ||
    after?.type !== "tableBlock" ||
    after.operation.type !== DOCUMENT_OP_TYPES.INSERT_TABLE ||
    after.operation.anchor !== "paragraph" ||
    after.operation.paragraphAfter !== "none"
  )
    return undefined;
  const insertion = ops.find(
    (op) => op.type === DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  );
  const carrier = insertion?.blocks.at(-1);
  if (carrier?.type !== "paragraph" || carrier.paraId === undefined)
    panic("A compiled table lost its trailing carrier allocation.");
  return carrier.paraId;
};

type MemberIdLeaseOptions = {
  lease: IdLease;
  siblings: readonly (readonly LeasedId[])[];
};

type ReplayOperation = {
  type: string;
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
};

const isPlainInsertion = ({
  type,
  operation,
  intents,
}: ReplayOperation): boolean => {
  const intent = intents.at(0);
  return (
    type === "plain" &&
    (operation.type === "insertBeforeBlock" ||
      operation.type === "insertAfterBlock") &&
    operation.moveId === undefined &&
    intents.length === 1 &&
    intent?.type === "insertParagraphs" &&
    intent.story === OP_STORIES.MAIN &&
    (intent.at.type === "before" || intent.at.type === "after") &&
    intent.move === undefined &&
    intent.newList === undefined &&
    intent.paragraphs.length === 1 &&
    intent.paragraphs.every(
      (paragraph) =>
        paragraph.sectionProperties === undefined &&
        paragraph.content.every(
          (node) =>
            node.type === "run" &&
            node.content.every(({ type: leaf }) => leaf === "text"),
        ),
    )
  );
};

/** A replay cannot spend identities admitted to another operation. */
const insertionMemberIdLease = ({
  lease,
  siblings,
}: MemberIdLeaseOptions): IdLease => {
  const blocks: IdBlock[] = [];
  for (const block of lease.blocks) {
    const excluded = siblings
      .flatMap((ids) =>
        ids
          .filter(
            ({ kind, id }) =>
              kind === block.kind && id >= block.start && id < block.end,
          )
          .map(({ id }) => id),
      )
      .toSorted((left, right) => left - right);
    let start = block.start;
    for (const id of excluded) {
      if (start < id) blocks.push({ kind: block.kind, start, end: id });
      start = id + 1;
    }
    if (start < block.end)
      blocks.push({ kind: block.kind, start, end: block.end });
  }
  return { blocks };
};

const LINE_BREAK_PATTERN = /\r\n|\r|\n/u;

type ResolveTableInsertionOptions = {
  operation: TableInsertion;
  /** Where the table goes in place of beside `paragraph`. */
  listStart?: ListStartPlace;
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  noteReferences: NoteReferenceLabels;
  refusal: ResolutionRefusal;
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  /** The table to copy, when the batch hands one for this operation. */
  template: PMNode | undefined;
  /** `BesidePendingJoinOptions.landedEarlier`. */
  landedEarlier: ReadonlySet<number>;
};

/**
 * A table beside the anchor, built as the legacy engine builds it (from the
 * cell texts, `buildTableNode`, or the parties, `buildSignatureTableNode`)
 * and placed where it places it. Beside a paragraph it goes as a paragraph
 * beside the anchor does (`besidePendingJoin`), or beside the paragraph as
 * it stands where no pending join reads that side: between the paragraph and
 * the block next to it, and after a container's last paragraph with the
 * empty paragraph the container ends with. From inside a table it goes
 * beside the outermost one, touching it.
 */
const resolveTableInsertion = ({
  operation,
  listStart,
  paragraph,
  projection,
  story,
  noteReferences,
  refusal,
  skipAlike,
  template,
  landedEarlier,
}: ResolveTableInsertionOptions) => {
  const side = operation.position ?? "after";
  const { schema } = projection.doc.type;
  let node: PMNode | null;
  if (operation.type === "insertTable")
    node =
      (template === undefined
        ? null
        : tableFromTemplate({ schema, template })) ??
      buildTableNode({ schema, rows: operation.rows });
  else {
    if (operation.parties.length === 0)
      return skipAlike("emptyOperation", "A signature table needs a party.");
    // A comment across the whole table would cross its cells, as the
    // legacy engine finds.
    if (operation.comment !== undefined)
      return skipAlike(
        "unsupportedBlock",
        "A signature table takes no comment.",
      );
    node = buildSignatureTableNode({ schema, parties: operation.parties });
  }
  if (node === null)
    return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
  const attrs = expectTableAttrs(node);
  const formatting = tableAttrsToFormatting(attrs);
  const table: Table = {
    type: "table",
    ...(formatting === undefined ? {} : { formatting }),
    ...(attrs.columnWidths === undefined
      ? {}
      : { columnWidths: [...attrs.columnWidths] }),
    rows: standaloneTableRowsFromProseMirror(node),
  };
  type InsertedOptions = {
    place: ProjectedParagraph;
    side: "before" | "after";
    paragraphAfter: "empty" | "none";
  };
  const inserted = ({
    place,
    side: placeSide,
    paragraphAfter,
  }: InsertedOptions) => {
    const position =
      placeSide === "after"
        ? place.start + place.node.content.size
        : place.start;
    const intent: EditorIntent = {
      type: "tableBlock",
      operation: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story,
        blockId: place.blockId,
        table,
        anchor: "paragraph",
        side: placeSide,
        paragraphAfter,
      },
    };
    return Result.ok({
      type: "plain" as const,
      operation,
      intents: [intent],
      normalizations: [],
      from: position,
      to: position,
    });
  };
  const besideTable = (blockId: string, position: number) =>
    Result.ok({
      type: "plain" as const,
      operation,
      intents: [
        {
          type: "tableBlock",
          operation: {
            type: DOCUMENT_OP_TYPES.INSERT_TABLE,
            story,
            blockId,
            table,
            anchor: "table",
            side,
          },
        } satisfies EditorIntent,
      ],
      normalizations: [],
      from: position,
      to: position,
    });
  if (listStart !== undefined)
    return Result.ok({
      type: "plain" as const,
      operation,
      intents: [
        {
          type: "tableBlock",
          operation: {
            type: DOCUMENT_OP_TYPES.INSERT_TABLE,
            story,
            blockId: listStart.blockId,
            table,
            anchor: "containerStart",
          },
        } satisfies EditorIntent,
      ],
      normalizations: [],
      from: listStart.position,
      to: listStart.position,
    });
  const stopsPendingJoin = (paragraphPos: number) =>
    joinStopsAtTableAfter({
      insertedBlock: "table",
      joinBeforeTable: "runsIntoFirstCell",
      doc: projection.doc,
      paragraphPos,
      noteReferences,
    });
  const pendingJoinRefusal = () =>
    refusal(
      PENDING_JOIN_INSERTION_REFUSAL.reason,
      PENDING_JOIN_INSERTION_REFUSAL.message,
    );
  // Whether accepting runs words into a table put at `position`: the
  // paragraph before it has a break pending deletion and keeps words, or goes
  // whole and hands its place to the one before it, whose break is pending
  // deletion too, and so on: the words of the whole join run in.
  const followsDeletedBreak = (position: number) => {
    const $position = projection.doc.resolve(position);
    for (let index = $position.index() - 1; index >= 0; index--) {
      const previous = $position.parent.child(index);
      if (
        previous.type.name !== "paragraph" ||
        !isDeletedPPrMark(previous.attrs["pPrMark"])
      )
        return false;
      if (!paragraphGoes(previous, "accept")) return true;
    }
    return false;
  };
  const pendingBreakRefusal = () =>
    refusal(
      PENDING_BREAK_TABLE_REFUSAL.reason,
      PENDING_BREAK_TABLE_REFUSAL.message,
    );
  const outer = findOutermostTableBoundary(projection.doc, paragraph.start - 1);
  if (outer !== null) {
    if (stopsPendingJoin(outer[side])) return pendingJoinRefusal();
    // Beside the outermost table, touching it, as the legacy engine puts
    // it: addressed by a paragraph of that table's own.
    const outerTable = projection.doc.nodeAt(outer.before);
    const blockId =
      outerTable === null
        ? undefined
        : tableBlockAddress(projection, {
            table: outerTable,
            tableStart: outer.before + 1,
          });
    if (blockId === undefined)
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    return besideTable(blockId, outer[side]);
  }
  const { place } = besidePendingJoin({
    side,
    paragraph,
    projection,
    noteReferences,
    landedEarlier,
  });
  // A paragraph the batch already put right after `place` stands between
  // it and the table, holding its words once accepted.
  const afterEarlier =
    side === "after" &&
    landedEarlier.has(place.start - 1 + place.node.nodeSize);
  const $place = projection.doc.resolve(place.start - 1);
  const neighbour = $place.parent.maybeChild(
    $place.index() + (side === "after" ? 1 : -1),
  );
  const gap =
    side === "after" ? place.start - 1 + place.node.nodeSize : place.start - 1;
  if (!afterEarlier && stopsPendingJoin(gap)) return pendingJoinRefusal();
  // Right after a paragraph whose break is pending deletion, accepting that
  // deletion would run the paragraph's words into the table's first cell;
  // one that goes whole leaves nothing of its own to run in, but passes on
  // the words of a paragraph joining into it.
  if (side === "after" && !afterEarlier && followsDeletedBreak(gap))
    return pendingBreakRefusal();
  // A pending-deleted table is at the start of this accepted gap. Use its
  // before edge so a later insertion explicitly after it can stay there.
  if (side === "before") {
    const start = gapStartAsTheBatchRuns(projection.doc, gap, landedEarlier);
    const pending = start < gap ? projection.doc.nodeAt(start) : null;
    if (pending?.type.name === "table") {
      const address = tableBlockAddress(projection, {
        table: pending,
        tableStart: start + 1,
      });
      if (address !== undefined) {
        if (stopsPendingJoin(start)) return pendingJoinRefusal();
        if (followsDeletedBreak(start)) return pendingBreakRefusal();
        return besideTable(address, start);
      }
    }
  }
  // After a container's last paragraph the container still ends with one.
  return inserted({
    place,
    side,
    paragraphAfter: side === "after" && neighbour === null ? "empty" : "none",
  });
};

/** A paragraph of the table's own that addresses a table edit, by its id. */
type KeptTableAddressOptions = {
  projection: CanonicalProjection;
  /** The table, row or cell the operation edits. */
  scope: PMNode;
  /** The position just before `scope`. */
  scopePosition: number;
  /** Paragraphs the batch deletes, which run before the table edit. */
  deleted: Pick<ReadonlySet<string>, "has">;
  /** Cells, by the position before them, the address may not stand in. */
  skipsCell?: (cellPosition: number) => boolean;
};

/**
 * The first paragraph standing directly in the cells of `scope` that the
 * batch keeps: a table, row or cell edit addressed by a paragraph the batch
 * deletes would find no paragraph once that deletion has run.
 */
const keptTableAddress = ({
  projection,
  scope,
  scopePosition,
  deleted,
  skipsCell = () => false,
}: KeptTableAddressOptions): string | undefined => {
  const cells: { cell: PMNode; position: number }[] = [];
  const collectCells = (row: PMNode, rowPosition: number) =>
    row.forEach((cell, offset) =>
      cells.push({ cell, position: rowPosition + 1 + offset }),
    );
  switch (scope.type.spec["tableRole"]) {
    case "cell":
      cells.push({ cell: scope, position: scopePosition });
      break;
    case "row":
      collectCells(scope, scopePosition);
      break;
    default:
      scope.forEach((row, offset) => {
        if (row.type.spec["tableRole"] === "row")
          collectCells(row, scopePosition + 1 + offset);
      });
  }
  // Each paragraph's first content position, in document order.
  const starts: number[] = [];
  for (const { cell, position } of cells) {
    if (skipsCell(position)) continue;
    cell.forEach((child, childOffset) => {
      if (child.type.name === "paragraph")
        starts.push(position + 1 + childOffset + 1);
    });
  }
  for (const start of starts) {
    const addressed = projection.addressAt(start);
    if (addressed.isOk() && !deleted.has(addressed.value.blockId))
      return addressed.value.blockId;
  }
  return undefined;
};

type TableIntentsAsItRunsOptions = {
  /** The batch as it stands when the operation runs. */
  document: Document;
  story: OpStory;
  /** The document as the batch read it. */
  projection: CanonicalProjection;
  plan: CanonicalBatchPlan;
  intents: readonly EditorIntent[];
};

/**
 * A table edit's intents, addressed by a paragraph that stands when it runs:
 * an earlier operation of the batch may have removed the one it was addressed
 * by, with the column or cell holding it. Any other paragraph of the table as
 * the batch read it names the same table; one a column deletion removes
 * itself only where no other stands.
 */
const tableIntentsAsItRuns = ({
  document,
  story,
  projection,
  plan: { address },
  intents,
}: TableIntentsAsItRunsOptions): readonly EditorIntent[] => {
  if (address.type !== "column" && address.type !== "cell") return intents;
  const stands = (blockId: string): boolean =>
    locateParagraphThroughTextBoxes(document, story, blockId) !== undefined;
  return intents.map((intent): EditorIntent => {
    if (
      intent.type !== "table" &&
      (intent.type !== "tableBlock" ||
        intent.operation.type !== DOCUMENT_OP_TYPES.DELETE_TABLE)
    )
      return intent;
    if (stands(intent.operation.blockId)) return intent;
    const table = projection.doc.nodeAt(address.tablePosition);
    if (table === null) return intent;
    const map = TableMap.get(table);
    const standing = (
      skipsCell?: KeptTableAddressOptions["skipsCell"],
    ): string | undefined =>
      keptTableAddress({
        projection,
        scope: table,
        scopePosition: address.tablePosition,
        deleted: { has: (blockId) => !stands(blockId) },
        ...(skipsCell === undefined ? {} : { skipsCell }),
      });
    if (intent.type === "tableBlock") {
      const blockId = standing();
      return blockId === undefined
        ? intent
        : { ...intent, operation: { ...intent.operation, blockId } };
    }
    const { operation } = intent;
    const blockId =
      operation.type === DOCUMENT_OP_TYPES.DELETE_COLUMN &&
      address.type === "column"
        ? (standing((cellPosition) => {
            const rect = map.findCell(cellPosition - address.tablePosition - 1);
            return (
              rect.left <= address.columnIndex &&
              address.columnIndex < rect.right
            );
          }) ?? standing())
        : standing();
    return blockId === undefined
      ? intent
      : { ...intent, operation: { ...operation, blockId } };
  });
};

const tableBlockAddress = (
  projection: CanonicalProjection,
  row: { table: PMNode; tableStart: number },
  deleted: ReadonlySet<string> = new Set(),
): string | undefined => {
  const at = tableAnchor({
    map: TableMap.get(row.table),
    table: row.table,
    tableStart: row.tableStart,
    left: 0,
    right: 0,
    top: 0,
    bottom: 0,
  });
  const addressed = at === undefined ? undefined : projection.addressAt(at);
  if (addressed === undefined || addressed.isErr()) return undefined;
  return deleted.has(addressed.value.blockId)
    ? (keptTableAddress({
        projection,
        scope: row.table,
        scopePosition: row.tableStart - 1,
        deleted,
      }) ?? addressed.value.blockId)
    : addressed.value.blockId;
};

/**
 * A new cell's paragraphs from its text: a line break starts a new
 * paragraph, as in a new row's cells.
 */
const cellParagraphs = (text: string): Paragraph[] =>
  splitCellParagraphTexts(text).map((line) => ({
    type: "paragraph",
    content: line === "" ? [] : [{ type: "run", content: inlineText(line) }],
  }));

/** The table offsets of the cells the rectangle covers, a merged one once per slot. */
const cellsIn = (map: TableMap, rectangle: TableRectangle): number[] => {
  const cells: number[] = [];
  for (let row = rectangle.top; row < rectangle.bottom; row += 1)
    for (let column = rectangle.left; column < rectangle.right; column += 1) {
      const offset = map.map[row * map.width + column];
      if (offset !== undefined) cells.push(offset);
    }
  return cells;
};

/** Whether a cell the rectangle covers is a slot its row omits. */
const holdsOmittedSlot = (
  table: PMNode,
  map: TableMap,
  rectangle: TableRectangle,
) => {
  for (let row = rectangle.top; row < rectangle.bottom; row += 1)
    for (let column = rectangle.left; column < rectangle.right; column += 1) {
      const offset = map.map[row * map.width + column];
      const cell = offset === undefined ? null : table.nodeAt(offset);
      if (cell === null || cell.attrs["_omittedGridSlot"] != null) return true;
    }
  return false;
};

/**
 * Whether a column inserted at the grid index runs through a cell spanning
 * both its sides, so that the cell would widen; a slot a row omits grows
 * instead, as the legacy engine grows it.
 */
const columnInsertionWidensCell = (table: PMNode, columnIndex: number) => {
  const map = TableMap.get(table);
  if (columnIndex <= 0 || columnIndex >= map.width) return false;
  for (let row = 0; row < map.height; row += 1) {
    const index = row * map.width + columnIndex;
    if (
      omittedSlotTakingColumn(map, table, row, columnIndex) === null &&
      map.map[index - 1] === map.map[index]
    )
      return true;
  }
  return false;
};

type TableStructureOperation = Extract<
  FolioDocumentOperation,
  {
    type:
      | "deleteTable"
      | "insertTableRow"
      | "deleteTableRow"
      | "insertTableColumn"
      | "deleteTableColumn"
      | "mergeTableCells"
      | "splitTableCell";
  }
>;

const isTableStructureOperation = (
  operation: FolioDocumentOperation,
): operation is TableStructureOperation =>
  operation.type === "deleteTable" ||
  operation.type === "insertTableRow" ||
  operation.type === "deleteTableRow" ||
  operation.type === "insertTableColumn" ||
  operation.type === "deleteTableColumn" ||
  operation.type === "mergeTableCells" ||
  operation.type === "splitTableCell";

type ResolveTableStructureOptions = {
  operation: TableStructureOperation;
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  refusal: ResolutionRefusal;
  /** Whether the batch records its changes as suggestions. */
  tracked: boolean;
  /** Resolve another paragraph the operation names, in the same story. */
  paragraphOf: (blockId: string) => ProjectedParagraph | undefined;
  /** The legacy engine skips an operation whose block is in no table alike. */
  notInTable: () => Result<never, CanonicalPublicOperationResolutionError>;
  /** A skip the legacy engine gives alike: no capability is missing. */
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  /** A skip the tracked vocabulary forces: the operation applies directly. */
  skipTracked: (
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  /** Paragraphs the batch deletes, which no table edit may be addressed by. */
  deleted: ReadonlySet<string>;
  /** The row a new row copies, when the batch hands one for this operation. */
  template: PMNode | undefined;
};

/**
 * A table, row or column change, addressed through the paragraph the
 * operation names, as the table edit the editor's own command lowers to. The
 * table it acts on is the innermost one holding that paragraph, as in the
 * legacy engine. A new row is the row the legacy engine builds
 * (`applyTableRowInsertion`, from the row beside it and the cell texts),
 * read off the table it leaves; the engine records it, tracked or not.
 */
const resolveTableStructure = ({
  operation,
  paragraph,
  projection,
  story,
  refusal,
  tracked,
  paragraphOf,
  notInTable,
  skipAlike,
  skipTracked,
  deleted,
  template,
}: ResolveTableStructureOptions) => {
  const { doc } = projection;
  const blockFrom = paragraph.start - 1;
  const row = findEnclosingTableRow(doc, blockFrom);
  if (row === null) return notInTable();
  const scopeOf = (scope: "cell" | "row" | "table") => {
    switch (scope) {
      case "table":
        return { scope: row.table, scopePosition: row.tablePosition };
      case "row":
        return {
          scope: row.table.child(row.rowIndex),
          scopePosition: row.rowPosition,
        };
      case "cell": {
        const cell = findEnclosingTableCell(doc, blockFrom);
        const node = cell === null ? null : doc.nodeAt(cell.cellPosition);
        return cell === null || node === null
          ? undefined
          : { scope: node, scopePosition: cell.cellPosition };
      }
      default: {
        const unreachable: never = scope;
        return unreachable;
      }
    }
  };
  /** The operation's own paragraph, unless the batch deletes it first. */
  const keptAddress = (scope: "cell" | "row" | "table") => {
    if (!deleted.has(paragraph.blockId)) return paragraph.blockId;
    const within = scopeOf(scope);
    return (
      (within && keptTableAddress({ projection, ...within, deleted })) ??
      paragraph.blockId
    );
  };
  const resolved = (intent: EditorIntent, from: number, to: number) =>
    Result.ok({
      type: "plain" as const,
      operation,
      intents: [intent],
      normalizations: [],
      from,
      to,
    });
  switch (operation.type) {
    case "deleteTable":
      return resolved(
        {
          type: "tableBlock",
          operation: {
            type: DOCUMENT_OP_TYPES.DELETE_TABLE,
            story,
            blockId: keptAddress("table"),
          },
        },
        row.tablePosition,
        row.tablePosition + row.table.nodeSize,
      );
    case "deleteTableRow":
      return resolved(
        {
          type: "tableRow",
          operation: {
            type: DOCUMENT_OP_TYPES.DELETE_ROW,
            story,
            blockId: keptAddress("row"),
          },
        },
        // The whole row: an edit inside it in the same batch overlaps.
        row.rowPosition,
        row.rowPosition + row.table.child(row.rowIndex).nodeSize,
      );
    case "deleteTableColumn": {
      const cell = findEnclosingTableCell(doc, blockFrom);
      if (cell === null) return notInTable();
      const map = TableMap.get(row.table);
      // The whole table: column indices of one table shift with each
      // deletion, so another edit of that table in the batch overlaps.
      const table = {
        from: row.tablePosition,
        to: row.tablePosition + row.table.nodeSize,
      };
      // Removing a table's only column removes the table.
      if (map.width === 1)
        return resolved(
          {
            type: "tableBlock",
            operation: {
              type: DOCUMENT_OP_TYPES.DELETE_TABLE,
              story,
              blockId: keptAddress("table"),
            },
          },
          table.from,
          table.to,
        );
      const column = {
        left: cell.leftColumnIndex,
        right: cell.leftColumnIndex + 1,
      };
      // The table is addressed by a paragraph the deletion keeps, as the
      // editor's column deletion addresses it.
      const at = tableAnchor(
        {
          map,
          table: row.table,
          tableStart: row.tableStart,
          ...column,
          top: 0,
          bottom: 0,
        },
        column,
      );
      const addressed = at === undefined ? undefined : projection.addressAt(at);
      if (addressed === undefined || addressed.isErr())
        return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
      const address = addressed.value;
      const blockId = deleted.has(address.blockId)
        ? (keptTableAddress({
            projection,
            scope: row.table,
            scopePosition: row.tablePosition,
            deleted,
            skipsCell: (cellPosition) => {
              const rect = map.findCell(cellPosition - row.tableStart);
              return rect.left >= column.left && rect.right <= column.right;
            },
          }) ?? address.blockId)
        : address.blockId;
      return resolved(
        {
          type: "table",
          operation: {
            type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
            story,
            blockId,
            column: cell.leftColumnIndex,
          },
        },
        table.from,
        table.to,
      );
    }
    case "insertTableColumn": {
      const insertion = findTableColumnInsertion({
        doc,
        blockFrom,
        position: operation.position ?? "after",
      });
      if (insertion === null) return notInTable();
      const columnCount = operation.columnCount ?? 1;
      const overflow = describeUnplacedCellTexts({
        cellTexts: operation.cellTexts,
        capacity:
          getInsertedPhysicalColumnRows(
            TableMap.get(row.table),
            row.table,
            insertion.columnIndex,
          ).length * columnCount,
        target: "column",
      });
      if (overflow !== null) return refusal("payloadDoesNotFit", overflow);
      // A suggestion records new cells; a column through a cell spanning
      // both its sides would widen that cell, which no suggestion records
      // in either engine.
      if (
        tracked &&
        columnInsertionWidensCell(row.table, insertion.columnIndex)
      )
        return skipTracked(
          "A suggested column cannot widen a cell that spans other columns.",
        );
      const blockId = tableBlockAddress(projection, row, deleted);
      if (blockId === undefined)
        return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
      const intent: EditorIntent = {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          story,
          blockId,
          column: insertion.columnIndex,
          // The neighbour's width, as the legacy engine gives it; a table
          // that states no grid keeps none, whatever the width.
          width: insertedColumnWidth(row.table, insertion.columnIndex) ?? 1,
          // The columns of one operation share the table's one grid change.
          ...(columnCount === 1 ? {} : { count: columnCount }),
        },
        ...(operation.cellTexts === undefined
          ? {}
          : { cellContent: operation.cellTexts.map(cellParagraphs) }),
      };
      return resolved(
        intent,
        row.tablePosition,
        row.tablePosition + row.table.nodeSize,
      );
    }
    case "mergeTableCells": {
      const start = findEnclosingTableCell(doc, blockFrom);
      if (start === null) return notInTable();
      const map = TableMap.get(row.table);
      let rectangle: TableRectangle;
      if (operation.rowCount === undefined) {
        const end = paragraphOf(operation.endBlockId);
        if (end === undefined) return refusal("missingBlock");
        const corner = findEnclosingTableCell(doc, end.start - 1);
        if (corner === null || corner.tablePosition !== start.tablePosition)
          return skipAlike(
            "unsupportedBlock",
            "The merge's corners are not in one table.",
          );
        rectangle = {
          left: Math.min(start.leftColumnIndex, corner.leftColumnIndex),
          top: Math.min(start.topRowIndex, corner.topRowIndex),
          right: Math.max(start.rightColumnIndex, corner.rightColumnIndex),
          bottom: Math.max(start.bottomRowIndex, corner.bottomRowIndex),
        };
      } else {
        rectangle = {
          left: start.leftColumnIndex,
          top: start.topRowIndex,
          right: start.rightColumnIndex,
          bottom: start.topRowIndex + operation.rowCount,
        };
        if (
          rectangle.right - rectangle.left !== 1 ||
          rectangle.bottom > map.height ||
          operation.rowCount < 2
        )
          return skipAlike(
            "unsupportedBlock",
            "A merge down takes one column and rows the table has.",
          );
      }
      if (
        rectangle.right - rectangle.left === 1 &&
        rectangle.bottom - rectangle.top === 1
      )
        return skipAlike(
          "noopOperation",
          "A merge of one cell changes nothing.",
        );
      if (
        tableRectangleCutsMergedCell(map, rectangle) ||
        holdsOmittedSlot(row.table, map, rectangle)
      )
        return skipAlike(
          "unsupportedBlock",
          "The merge would cut a merged cell or a slot its row omits.",
        );
      if (new Set(cellsIn(map, rectangle)).size < 2)
        return skipAlike(
          "unsupportedBlock",
          "The merge covers a single merged cell.",
        );
      const blockId = tableBlockAddress(projection, row, deleted);
      if (blockId === undefined)
        return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
      return resolved(
        {
          type: "table",
          operation: {
            type: DOCUMENT_OP_TYPES.MERGE_CELLS,
            story,
            blockId,
            ...rectangle,
          },
        },
        row.tablePosition,
        row.tablePosition + row.table.nodeSize,
      );
    }
    case "splitTableCell": {
      const cell = findEnclosingTableCell(doc, blockFrom);
      if (cell === null) return notInTable();
      if (
        cell.rightColumnIndex - cell.leftColumnIndex === 1 &&
        cell.bottomRowIndex - cell.topRowIndex === 1
      )
        return skipAlike(
          "noopOperation",
          "A one-row, one-column cell has nothing to split.",
        );
      // The cell is addressed by its own paragraph, the merge's first.
      return resolved(
        {
          type: "table",
          operation: {
            type: DOCUMENT_OP_TYPES.SPLIT_CELL,
            story,
            blockId: keptAddress("cell"),
          },
        },
        row.tablePosition,
        row.tablePosition + row.table.nodeSize,
      );
    }
    case "insertTableRow": {
      const side = operation.position ?? "after";
      const insertion = findTableRowInsertion({
        doc,
        blockFrom,
        position: side,
      });
      if (insertion === null) return notInTable();
      const overflow = describeUnplacedCellTexts({
        cellTexts: operation.cellTexts,
        capacity: insertion.cells.length,
        target: "row",
      });
      if (overflow !== null) return refusal("payloadDoesNotFit", overflow);
      const built = applyTableRowInsertion({
        tr: EditorState.create({ doc }).tr,
        insertion,
        cellTexts: operation.cellTexts,
        revision: null,
        ...(template === undefined ? {} : { template }),
      });
      if (built.type !== "applied")
        return refusal("unsupportedBlock", undefined, INTERNAL);
      const table = built.transaction.doc.nodeAt(
        built.transaction.mapping.map(row.tablePosition, -1),
      );
      if (table === null || table.type.name !== "table")
        return refusal("unsupportedBlock", undefined, INTERNAL);
      const index = side === "before" ? row.rowIndex : row.rowIndex + 1;
      const inserted = standaloneTableRowsFromProseMirror(table)[index];
      if (inserted === undefined)
        return refusal("unsupportedBlock", undefined, INTERNAL);
      const {
        structuralChange: _structuralChange,
        propertyChanges: _propertyChanges,
        ...plain
      } = inserted;
      return resolved(
        {
          type: "tableRow",
          operation: {
            type: DOCUMENT_OP_TYPES.INSERT_ROW,
            story,
            blockId: keptAddress("row"),
            at: index,
            row: plain,
          },
        },
        insertion.rowPosition,
        insertion.rowPosition,
      );
    }
    default: {
      const unreachable: never = operation;
      return unreachable;
    }
  }
};

/**
 * The operations a canonical suggestion can carry: the legacy engine's;
 * splits and merges, which the canonical session records as tracked paragraph
 * marks stamped with the suggestion; and comments, which it holds as proposed
 * by the suggestion until it is accepted or rejected.
 */
const CANONICAL_SUGGESTED_OPERATION_TYPES: ReadonlySet<
  FolioDocumentOperation["type"]
> = new Set<FolioDocumentOperation["type"]>([
  ...SUGGESTED_SUPPORTED_OPERATION_TYPES,
  "splitBlock",
  "mergeBlockWithNext",
  "commentOnRange",
  "commentOnBlock",
]);

/** The operations that only add marks over text: formatting and comments. */
const ANNOTATING_OPERATION_TYPES: ReadonlySet<FolioDocumentOperation["type"]> =
  new Set<FolioDocumentOperation["type"]>([
    "formatRange",
    "commentOnRange",
    "commentOnBlock",
  ]);

/** A block operation without the move it names: a plain insertion or deletion. */
const withoutMove = (
  operation: FolioDocumentOperation,
): FolioDocumentOperation => {
  if (
    operation.type !== "deleteBlock" &&
    operation.type !== "insertAfterBlock" &&
    operation.type !== "insertBeforeBlock"
  )
    return operation;
  const { moveId: _moveId, ...plain } = operation;
  return plain;
};

/** The move a block operation names, if it is a half of one. */
const blockMoveId = (operation: FolioDocumentOperation): string | undefined =>
  operation.type === "deleteBlock" ||
  operation.type === "insertAfterBlock" ||
  operation.type === "insertBeforeBlock"
    ? operation.moveId
    : undefined;

/**
 * A line's text with each tab and line break as the element it is; CRLF, CR
 * and LF are each one line break.
 */
/**
 * A line as runs: paired `**bold**` and `***bold italic***` markers state
 * that emphasis on their words, as the default engine reads them; the rest
 * states no run property, so the paragraph's style formats it.
 */
const emphasisRuns = (line: string): Run[] =>
  parseInlineEmphasisRuns(line).flatMap(({ text, bold, italic }): Run[] => {
    const content = inlineText(text);
    if (content.length === 0) return [];
    if (!bold && !italic) return [{ type: "run", content }];
    return [
      {
        type: "run",
        content,
        formatting: {
          ...(bold ? { bold } : {}),
          ...(italic ? { italic } : {}),
        },
      },
    ];
  });

/** Words of a replacement its emphasis markers mark, by offset in the words it writes. */
type EmphasizedSpan = { start: number; end: number; patch: RunPropsPatch };

/**
 * The spans of the words a replacement writes that its paired `**bold**` and
 * `***bold italic***` markers mark, and the emphasis each states.
 */
const emphasizedSpans = (
  emphasis: readonly InlineEmphasisRun[] | undefined,
): EmphasizedSpan[] => {
  const spans: EmphasizedSpan[] = [];
  let offset = 0;
  for (const { text, bold, italic } of emphasis ?? []) {
    const end = offset + text.length;
    if (end > offset && (bold || italic))
      spans.push({
        start: offset,
        end,
        patch: {
          ...(bold ? { bold: true } : {}),
          ...(italic ? { italic: true } : {}),
        },
      });
    offset = end;
  }
  return spans;
};

/** A plan that replaces the whole match, its words standing together. */
const rewriteWhole = (source: string, target: string) => [
  { start: 0, end: source.length, text: target },
];

/**
 * Paragraph properties kept by a replacement that drops the paragraph's
 * formatting: the review records of its numbering, which only resolving
 * them retires.
 */
const NUMBERING_REVIEW_RECORDS: ReadonlySet<string> = new Set([
  "numberingChangeXml",
  "numberingInsertionXml",
]);

/**
 * The patch a replacement that drops the paragraph's formatting writes:
 * every property the paragraph states goes, as the default engine writes a
 * fresh paragraph, save the review records of its numbering.
 */
const droppedParagraphFormatting = (
  formatting: ParagraphFormatting | undefined,
): ParagraphPropsPatch => {
  const patch: Record<string, null> = {};
  for (const key of Object.keys(formatting ?? {}))
    if (!NUMBERING_REVIEW_RECORDS.has(key)) patch[key] = null;
  // Every key is a ParagraphFormatting field set to null.
  return patch as ParagraphPropsPatch;
};

const inlineText = (line: string): Run["content"] =>
  line.split(/(\t|\r\n|\r|\n)/u).flatMap((part): Run["content"] => {
    if (part === "\t") return [{ type: "tab" }];
    if (LINE_BREAK_PATTERN.test(part)) return [{ type: "break" }];
    return part === "" ? [] : [{ type: "text", text: part }];
  });

type ResolveBlockDeletionOptions = {
  operation: Extract<FolioDocumentOperation, { type: "deleteBlock" }>;
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>;
  projection: CanonicalProjection;
  projectionStory: OpStory;
  doc: PMNode;
  tracked: boolean;
  /** A tracked deletion under the joined paragraph's identity. */
  deletesJoinedWords: boolean;
  /**
   * Whether the batch inserts a paragraph beside this one, which no
   * paragraph follows: the insertion applies first, and the deletion then
   * finds this paragraph by identity.
   */
  gainsFollower: boolean;
  /**
   * Whether this paragraph stands between two tables once the batch's blocks
   * land (`deletedParagraphPlacement`): the deletion takes its words only.
   */
  keepsTablesApart: boolean;
  /**
   * Whether the batch puts a table right before this paragraph, which then
   * follows the table rather than the paragraph before it.
   */
  followsInsertedTable: boolean;
  /**
   * Whether the batch puts a table right after this paragraph, which then
   * stands before the table when this deletion runs.
   */
  precedesInsertedTable: boolean;
  /**
   * Whether either physical neighbour is a table the batch removes.
   */
  bordersRemovedTable: boolean;
  /**
   * Whether the batch inserts a paragraph after the paragraph before this
   * one, which then stands between them when this deletion runs.
   */
  previousGainsFollower: boolean;
  /**
   * Whether, tracked, only paragraphs that pending deletions already take
   * follow this one, so that it ends its container once accepted.
   */
  endsOnceAccepted: boolean;
  /**
   * Whether the batch puts a paragraph after the table before this one,
   * which then stands between them when this deletion runs.
   */
  tableBeforeGainsFollower: boolean;
  /** The run of paragraphs right before this one the batch deletes. */
  deletedRun: DeletedRun;
  /** The paragraphs the batch deletes, by id. */
  deletedByBatch: ReadonlySet<string>;
  refusal: ResolutionRefusal;
};

type CanonicalPublicOperationsOptions = {
  session: CanonicalSession;
  getState: (story: OpStory) => EditorState | null;
  publish: (commit: CanonicalCommit) => boolean;
  /** How many batches stay undoable; older handles are `unknownHandle`. */
  undoLimit: number;
};

type PublicOperationStaging = ReturnType<
  CanonicalSession["beginVerifiedOperationStaging"]
>;
type VerifiedOperation = Parameters<PublicOperationStaging["append"]>[0];

type CompilePublicIntentsOptions = {
  document: Document;
  intents: readonly EditorIntent[];
  mode: "direct" | "tracked-changes";
  author: string;
  date: string;
  /** The pending suggestion the records belong to, in suggested mode. */
  suggestionId?: string | undefined;
  nextRevisionId: number;
  allocate: ReturnType<typeof createEditorIntentIdAllocator>;
  /** Paragraph identities the session has consumed, retired ones included. */
  reservedBlockIds: ReadonlySet<string>;
  /**
   * The ids the session may mint, when it holds a lease: revision ids then
   * come from the allocator's leased ones rather than one contiguous range.
   */
  idLease: IdLease | undefined;
  /** A window-owned paragraph allocation, independent of the session revision lease. */
  paragraphAllocationLease?: IdLease;
  allocations?: readonly ReturnType<
    ReturnType<typeof createEditorIntentIdAllocator>
  >[];
  verifyOperation?: PublicOperationStaging["apply"];
  /** The story the intents edit. */
  story: OpStory;
};

const PROPERTY_FOLDING_INTENTS: ReadonlySet<EditorIntent["type"]> = new Set([
  "joinParagraphs",
  "splitParagraph",
  "formatParagraph",
  "setList",
  "deleteParagraphs",
]);

/** Each paragraph property change of the story, by revision id. */
const propertyChangesById = (document: Document, story: OpStory) =>
  new Map(
    paragraphsThroughTextBoxes(document, story).flatMap(
      ({ paragraph }) =>
        paragraph.propertyChanges?.map(
          (change) => [change.info.id, change] as const,
        ) ?? [],
    ),
  );

// Bound the generic allocator's prospective paragraph lookup and keep
// the unrestricted revision/control ranges of an unleased editor.
/**
 * Room past the highest revision or annotation id below which a batch draws
 * its revision ids from the free ones instead: more than any one batch asks.
 */
const REVISION_ID_WRAP_HEADROOM = 1 << 16;

/**
 * A lease of every id, but for revisions only the ones no revision record or
 * annotation holds, lowest first.
 */
const wrappedRevisionLease = (taken: ReadonlySet<number>): IdLease => {
  const blocks: IdBlock[] = Object.values(ID_LEASE_KINDS)
    .filter((kind) => kind !== ID_LEASE_KINDS.REVISION)
    .map((kind) => ({ kind, start: 1, end: ID_KIND_LIMITS[kind] }));
  let start = 1;
  for (const id of [...taken].toSorted((left, right) => left - right)) {
    if (id < start) continue;
    if (id > start)
      blocks.push({ kind: ID_LEASE_KINDS.REVISION, start, end: id });
    start = id + 1;
  }
  const end = ID_KIND_LIMITS[ID_LEASE_KINDS.REVISION];
  if (start < end) blocks.push({ kind: ID_LEASE_KINDS.REVISION, start, end });
  return { blocks };
};

const allocatorLeaseForParagraphs = (paragraphLease: IdLease | undefined) =>
  paragraphLease === undefined
    ? undefined
    : {
        blocks: [
          ...paragraphLease.blocks,
          {
            kind: ID_LEASE_KINDS.REVISION,
            start: 1,
            end: ID_KIND_LIMITS.revision,
          },
          {
            kind: ID_LEASE_KINDS.CONTROL,
            start: 1,
            end: ID_KIND_LIMITS.control,
          },
        ],
      };

type PrimaryRevisionIdOptions = {
  operation: FolioDocumentOperation;
  /** The intents as the operation ran. */
  intents: readonly EditorIntent[];
  /** The document after the operation. */
  document: Document;
  revisionIds: readonly number[];
};

/**
 * The revision an applied operation names as its own: a tracked merge's is
 * the paragraph-mark deletion that is the merge, not the separator it types;
 * any other operation's is its first record.
 */
const primaryRevisionId = ({
  operation,
  intents,
  document,
  revisionIds,
}: PrimaryRevisionIdOptions): number | undefined => {
  const join =
    operation.type === "mergeBlockWithNext"
      ? intents.find(
          (
            intent,
          ): intent is Extract<EditorIntent, { type: "joinParagraphs" }> =>
            intent.type === "joinParagraphs",
        )
      : undefined;
  const mark =
    join === undefined
      ? undefined
      : locateParagraphThroughTextBoxes(document, join.story, join.blockId)
          ?.paragraph.pPrMark;
  const deletion =
    mark?.kind === "del" ? mark.info.id : mark?.deletion?.info.id;
  return deletion !== undefined && revisionIds.includes(deletion)
    ? deletion
    : revisionIds.at(0);
};

const compilePublicIntents = ({
  document,
  intents,
  mode,
  author,
  date,
  suggestionId,
  nextRevisionId,
  allocate,
  reservedBlockIds,
  idLease,
  paragraphAllocationLease,
  allocations: admittedAllocations,
  verifyOperation,
  story,
}: CompilePublicIntentsOptions) => {
  const ops: DocumentOp[] = [];
  const recorded: RecordedRevisionIds[] = [];
  let verifiedOperation: VerifiedOperation | undefined;
  const absorbed: ParagraphAbsorption[] = [];
  const source = document;
  // The changes pending joins of earlier operations of the batch own, under
  // this operation's stamp. The compiler refuses a fold into a join's change
  // from another stamp, as one at a time; stamped alike, it would take these
  // over, so an intent that changes or drops one is refused here. Only
  // intents that set, join, split or delete paragraphs reach those changes.
  const standing =
    mode === "direct" ||
    !intents.some((intent) => PROPERTY_FOLDING_INTENTS.has(intent.type))
      ? []
      : [...propertyChangesById(source, story)].flatMap(([id, change]) =>
          change.resolvesWith !== undefined &&
          change.info.author === author &&
          change.info.date === date &&
          change.info.suggestionId === suggestionId
            ? [[id, JSON.stringify(change)] as const]
            : [],
        );
  // The paragraph each tracked join of this operation deletes the mark of,
  // by the paragraph it keeps: properties the operation sets there belong to
  // that join.
  const joinedInto = new Map<string, string>();
  // A split's new half takes the id allocated here, not the one it resolved
  // with: a later intent naming that half names the allocated one.
  const allocated = new Map<string, string>();
  const allocations: ReturnType<
    ReturnType<typeof createEditorIntentIdAllocator>
  >[] = [];
  for (const [intentIndex, sourceIntent] of intents.entries()) {
    const paragraphLease =
      sourceIntent.type === "insertParagraphs"
        ? paragraphAllocationLease
        : undefined;
    const ids =
      admittedAllocations === undefined
        ? allocate(
            document,
            sourceIntent,
            idLease ?? allocatorLeaseForParagraphs(paragraphLease),
          )
        : admittedAllocations.at(intentIndex);
    if (ids === undefined) panic("An admitted intent lost its allocation.");
    allocations.push(ids);
    const compilerLease = idLease ?? paragraphLease;
    const leased =
      compilerLease === undefined ? {} : { idLease: compilerLease };
    let intent: EditorIntent = sourceIntent;
    if (sourceIntent.type === "splitParagraph") {
      allocated.set(sourceIntent.newBlockId, ids.newBlockId);
      intent = {
        type: "splitParagraph",
        at: sourceIntent.at,
        ...(sourceIntent.to === undefined ? {} : { to: sourceIntent.to }),
        ...(sourceIntent.newHalf === undefined
          ? {}
          : { newHalf: sourceIntent.newHalf }),
        newBlockId: ids.newBlockId,
      };
    } else if (sourceIntent.type === "formatParagraph") {
      const blockId = allocated.get(sourceIntent.at.blockId);
      if (blockId !== undefined)
        intent = { ...sourceIntent, at: { ...sourceIntent.at, blockId } };
    }
    const firstRevisionId = nextRevisionId;
    const revisionIds =
      idLease === undefined
        ? ids.newIds.revision.map((_, index) => firstRevisionId + index + 1)
        : ids.newIds.revision;
    const compiled = compileEditorIntent(document, {
      intent,
      ...(intent.type === "replaceText" &&
      intent.formattingSources !== undefined
        ? { replacementFormattingDocument: source }
        : {}),
      // An operation's text is body text: it never continues a note reference.
      typedFormatting: TYPED_TEXT_FORMATTING.BODY_TEXT,
      mode:
        mode === "direct"
          ? {
              type: "editing",
              newIds: ids.newIds,
              reservedBlockIds,
              ...leased,
            }
          : {
              type: "suggesting",
              reservedBlockIds,
              ...leased,
              revision: {
                id: idLease === undefined ? firstRevisionId : ids.revisionId,
                author,
                date,
                ...(suggestionId === undefined ? {} : { suggestionId }),
              },
              newIds: { control: ids.newIds.control, revision: revisionIds },
            },
    });
    if (compiled.isErr()) return Result.err(compiled.error);
    const [only] = compiled.value.ops;
    const verified =
      verifyOperation !== undefined &&
      intents.length === 1 &&
      compiled.value.ops.length === 1 &&
      ((sourceIntent.type === "insertParagraphs" &&
        only?.type === DOCUMENT_OP_TYPES.INSERT_BLOCKS) ||
        (sourceIntent.type === "replaceText" &&
          (only?.type === DOCUMENT_OP_TYPES.INSERT_CONTENT ||
            only?.type === DOCUMENT_OP_TYPES.INSERT_TEXT)))
        ? verifyOperation(document, only)
        : undefined;
    let applied: VerifiedOperation;
    if (verified === undefined) {
      const unverified = applyDocumentOps(document, compiled.value.ops);
      if (unverified.isErr()) return Result.err(unverified.error);
      absorbed.push(...(unverified.value.touched.absorbed ?? []));
      applied = unverified.value;
    } else {
      // A verified operation inserts blocks or text: it absorbs no paragraph.
      if (verified.isErr()) return Result.err(verified.error);
      verifiedOperation = verified.value;
      applied = verified.value;
    }
    // An intent the document already satisfies is a no-op.
    if (applied.document === document) continue;
    // The notes whose last live reference the intent's tracked deletion
    // took go with it, as operations of their own beside the intent's. A
    // comparison states each note's own changes as a story of its own, as
    // the legacy reviewer does.
    let after = readModel(applied.document);
    const noteDeletion =
      mode === "direct" || noteReferenceEditsAllowed()
        ? undefined
        : noteDeletionOps({
            before: document,
            after,
            stamp: {
              id: idLease === undefined ? firstRevisionId : ids.revisionId,
              author,
              date,
              ...(suggestionId === undefined ? {} : { suggestionId }),
            },
            revisionIds,
          });
    if (noteDeletion?.isErr() === true) return Result.err(noteDeletion.error);
    const notes = noteDeletion?.unwrap();
    const appliedRevisions = [...applied.revisions];
    if (notes !== undefined && notes.ops.length > 0) {
      const noted = applyDocumentOps(after, notes.ops);
      if (noted.isErr()) return Result.err(noted.error);
      // The staged operation no longer holds the whole edit.
      verifiedOperation = undefined;
      after = noted.value.document;
      appliedRevisions.push(...notes.revisions);
    }
    document = after;
    const now =
      standing.length === 0 ? undefined : propertyChangesById(after, story);
    if (
      now !== undefined &&
      standing.some(([id, change]) => JSON.stringify(now.get(id)) !== change)
    )
      return Result.err(
        new DocumentOpRefusal({
          reason: DOCUMENT_OP_REFUSAL_REASONS.PENDING_PARAGRAPH_PROPERTY_CHANGE,
          message:
            "Resolve the pending paragraph join before tracking a different change to its properties.",
          opType: undefined,
        }),
      );
    ops.push(...compiled.value.ops, ...(notes?.ops ?? []));
    recorded.push({ revisions: appliedRevisions, folded: applied.folded });
    for (const id of appliedRevisions)
      nextRevisionId = Math.max(nextRevisionId, id + 1);
    if (mode === "direct") continue;
    if (sourceIntent.type === "joinParagraphs")
      joinedInto.set(idKey(sourceIntent.nextBlockId), sourceIntent.blockId);
    if (intent.type !== "formatParagraph") continue;
    const joined = joinedInto.get(idKey(intent.at.blockId));
    if (joined === undefined) continue;
    const link = joinPropertyChangeLink({
      before: source,
      after: document,
      story: intent.at.story,
      blockId: joined,
      propertiesBlockId: intent.at.blockId,
    });
    if (link === undefined) continue;
    const linked = applyDocumentOps(document, [link]);
    if (linked.isErr()) return Result.err(linked.error);
    document = readModel(linked.value.document);
    ops.push(link);
  }
  return Result.ok({
    document,
    ops,
    ...combineRecordedRevisions(recorded),
    nextRevisionId,
    allocations,
    verifiedOperation,
    absorbed,
  });
};

const CanonicalPublicOperationResolutionErrorBase: TaggedErrorClass<"CanonicalPublicOperationResolutionError"> =
  TaggedError("CanonicalPublicOperationResolutionError");

class CanonicalPublicOperationResolutionError extends CanonicalPublicOperationResolutionErrorBase<{
  message: string;
  reason: FolioAIEditSkipReason;
  /**
   * What the skip tells the caller beyond its reason; `undefined` when the
   * reason says it all, as the legacy engine reports such a skip.
   */
  detail: string | undefined;
  /**
   * Why the session refused it, reported as the skip's `canonicalRefusal`;
   * absent for an operation the legacy engine skips alike, under the reason
   * alone, such as a row operation whose block is in no table.
   */
  ground?: CanonicalRefusalGround;
  /**
   * The modes the refusal holds in: `"trackedMode"` for an operation the
   * tracked vocabulary has no record for, which applies directly. Absent, it
   * holds in every mode.
   */
  refusedIn?: "everyMode" | "trackedMode";
}> {}

type MergeFormattingConflictOptions = {
  before: Document;
  after: Document;
  story: OpStory;
  ops: readonly DocumentOp[];
  author: string;
};

/** Both the tracked join and a subsequent property alignment may write the survivor. */
const overwritesForeignMergeFormatting = ({
  before,
  after,
  story,
  ops,
  author,
}: MergeFormattingConflictOptions): boolean => {
  const targets = new Set<string>();
  for (const op of ops) {
    if (op.type === DOCUMENT_OP_TYPES.JOIN_BLOCKS) targets.add(op.nextBlockId);
    if (op.type === DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS)
      targets.add(op.blockId);
  }
  const foreignChanges = (paragraph: Paragraph) =>
    paragraph.propertyChanges?.filter(({ info }) => info.author !== author) ??
    [];
  const pending = paragraphsThroughTextBoxes(before, story).flatMap(
    ({ paragraph }) => {
      if (paragraph.paraId === undefined || !targets.has(paragraph.paraId))
        return [];
      const changes = foreignChanges(paragraph);
      return changes.length === 0 ? [] : [[paragraph.paraId, changes] as const];
    },
  );
  if (pending.length === 0) return false;
  const survivors = new Map(
    paragraphsThroughTextBoxes(after, story).map(({ paragraph }) => [
      paragraph.paraId,
      paragraph,
    ]),
  );
  return pending.some(([id, revisions]) => {
    const survivor = survivors.get(id);
    return (
      survivor === undefined ||
      JSON.stringify(revisions) !== JSON.stringify(foreignChanges(survivor))
    );
  });
};

type BatchPropertyRefusalOptions = {
  document: Document;
  story: OpStory;
  operation: FolioDocumentOperation;
};

/** Recheck the native property boundary after earlier operations of this batch. */
const batchPropertyRefusal = ({
  document,
  story,
  operation,
}: BatchPropertyRefusalOptions) => {
  switch (operation.type) {
    case "setBlockParagraphProperties":
      break;
    case "replaceBlock":
      if (operation.styleId === undefined) return undefined;
      break;
    default:
      return undefined;
  }
  const paragraph = paragraphsThroughTextBoxes(document, story).find(
    (entry) => entry.paragraph.paraId === operation.blockId,
  )?.paragraph;
  if (
    operation.type === "replaceBlock" &&
    operation.styleId === (paragraph?.formatting?.styleId ?? null)
  )
    return undefined;
  return pendingParagraphPropertiesRefusal(paragraph?.pPrMark);
};

type MapColumnDeletionIntentsOptions = {
  intents: readonly EditorIntent[];
  columnIndex: number;
  /** Columns the batch inserted in the gap before the deleted one. */
  inserted: number;
  /** Columns the batch inserted in the gap after it. */
  insertedAfter: number;
};

/**
 * The native column applier uses the same successful same-gap insertion
 * count; a column inserted on either side keeps a one-column table standing.
 */
const mapColumnDeletionIntents = ({
  intents,
  columnIndex,
  inserted,
  insertedAfter,
}: MapColumnDeletionIntentsOptions): readonly EditorIntent[] => {
  if (inserted === 0 && insertedAfter === 0) return intents;
  return intents.map((intent): EditorIntent => {
    if (
      intent.type === "table" &&
      intent.operation.type === DOCUMENT_OP_TYPES.DELETE_COLUMN
    )
      return inserted === 0
        ? intent
        : {
            type: "table",
            operation: { ...intent.operation, column: columnIndex + inserted },
          };
    // A source table with one column no longer disappears after the batch
    // inserts beside that column: remove the original column, keeping the new.
    if (
      intent.type === "tableBlock" &&
      intent.operation.type === DOCUMENT_OP_TYPES.DELETE_TABLE
    )
      return {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
          story: intent.operation.story,
          blockId: intent.operation.blockId,
          column: columnIndex + inserted,
        },
      };
    return intent;
  });
};

type ResolveCanonicalPublicOperationOptions = {
  operation: FolioDocumentOperation;
  /**
   * Where an insertion goes in place of beside the block it names: the
   * start of a list, which no block names (`restatedBesideEmptyControl`).
   */
  listStart?: ListStartPlace;
  options: CanonicalPublicOperationOptions;
  state: EditorState;
  story: OpStory;
  diff: ReturnType<typeof wordDiffSessionFromOptions>["diff"];
  /** Whether the batch records its edits as revisions. */
  tracked: boolean;
  /** Whether the batch records its edits as suggestions. */
  suggested: boolean;
  /** A new list of one kind for this operation, minted after the batch's earlier ones. */
  mintList: ReturnType<typeof createBatchListMinter>;
  /** The blocks the batch deletes, and those it inserts beside. */
  besideDeletions: BesideDeletions;
  /** The comments the batch adds, which stand before its edits run. */
  batchComments: readonly BatchCommentOperation[];
};

type BatchCommentOperation = Extract<
  FolioDocumentOperation,
  { type: "commentOnBlock" | "commentOnRange" }
>;

type CleanTextSpan = { start: number; end: number };

/**
 * The words of `blockId` the batch's comments cover, in its clean text: a
 * batch's comments run before its edits, so an edit there meets them as it
 * meets the comments the paragraph already holds. A comment whose target no
 * longer matches covers nothing; its own resolution refuses it.
 */
const batchCommentSpans = (
  comments: readonly BatchCommentOperation[],
  blockId: string,
  text: string,
): CleanTextSpan[] =>
  comments.flatMap((comment): CleanTextSpan[] => {
    switch (comment.type) {
      case "commentOnRange": {
        const { range } = comment;
        return range.blockId === blockId &&
          hashFolioAIBlockText(
            text.slice(range.startOffset, range.endOffset),
          ) === range.selectedTextHash
          ? [{ start: range.startOffset, end: range.endOffset }]
          : [];
      }
      case "commentOnBlock": {
        if (comment.blockId !== blockId) return [];
        if (comment.quote === undefined)
          return [{ start: 0, end: text.length }];
        const start = text.indexOf(comment.quote);
        return start < 0 || comment.quote === ""
          ? []
          : [{ start, end: start + comment.quote.length }];
      }
      default: {
        const unreachable: never = comment;
        return unreachable;
      }
    }
  });

type BesideDeletions = {
  deleted: ReadonlySet<string>;
  insertedBeside: ReadonlySet<string>;
  /**
   * The blocks the batch puts a table after, those it puts one before, and
   * the paragraphs naming the tables it deletes.
   */
  tables: {
    after: ReadonlySet<string>;
    before: ReadonlySet<string>;
    deleted: ReadonlySet<string>;
  };
  /** The blocks the batch puts a paragraph after. */
  paragraphsAfter: ReadonlySet<string>;
  /**
   * The block boundaries where the operations before this one, in input
   * order, put a paragraph (`BesidePendingJoinOptions.landedEarlier`).
   */
  landedEarlier: (operationId: string) => ReadonlySet<number>;
  /** The insertions that put a block beside `blockId`, in input order. */
  placedBeside: (
    blockId: string,
  ) => readonly (BlockInsertion | TableInsertion)[];
};

type DeletedRunBeforeOptions = {
  doc: PMNode;
  /** Where the content of the paragraph whose run this is starts. */
  paragraphStart: number;
  besideDeletions: BesideDeletions;
  /**
   * Tracked, a paragraph a pending deletion already takes whole stands in
   * the run, as accepting leaves it gone.
   */
  tracked: boolean;
};

/**
 * The paragraphs right before a container's final paragraph that the batch
 * deletes, as `finalBreakCarrier` reads them, or, tracked, `bounded`: a
 * block the batch inserts stands right before them once it runs (after the
 * kept paragraph before them, or before the first of them). The deletion is
 * claimed from `claimFrom`, before that insertion and the run's deletions,
 * so it runs after them and retires the break before the run as it stands.
 */
type DeletedRun =
  | { type: FinalBreakCarrier<unknown>["type"] }
  | { type: "bounded"; claimFrom: number };

/** Whether the batch deletes none of the paragraphs right before. */
const noDeletedRunBefore = ({ type }: DeletedRun): boolean => {
  switch (type) {
    case "none":
      return true;
    case "kept":
    case "runStart":
    case "bounded":
      return false;
    default: {
      const unreachable: never = type;
      return unreachable;
    }
  }
};

/**
 * The paragraphs right before the one starting at `paragraphStart` that the
 * batch deletes, putting nothing beside them, as `finalBreakCarrier` reads
 * them: `kept` when a paragraph the batch keeps precedes them, `runStart`
 * when they reach back to the container's start or another block, `none`
 * when the batch keeps the paragraph right before it.
 */
const deletedRunBefore = ({
  doc,
  paragraphStart,
  besideDeletions,
  tracked,
}: DeletedRunBeforeOptions): DeletedRun => {
  const at = doc.resolve(paragraphStart - 1);
  let deleted = 0;
  let position = paragraphStart - 1;
  for (let index = at.index() - 1; index >= 0; index--) {
    const previous = at.parent.child(index);
    position -= previous.nodeSize;
    if (previous.type.name !== "paragraph") break;
    if (tracked && goesWholeOnAcceptance(previous)) continue;
    const id: unknown = previous.attrs["paraId"];
    if (typeof id !== "string")
      return { type: deleted === 0 ? "none" : "kept" };
    const blockAfter =
      besideDeletions.paragraphsAfter.has(id) ||
      besideDeletions.tables.after.has(id);
    // A block the batch puts after the kept paragraph stands between it and
    // the run instead: the break the deletion retires is that block's.
    if (!besideDeletions.deleted.has(id) && blockAfter)
      return tracked && deleted > 0
        ? { type: "bounded", claimFrom: position + 1 }
        : { type: "none" };
    // So does one the batch puts before the first paragraph of the run.
    const blockBefore =
      !blockAfter &&
      (besideDeletions.insertedBeside.has(id) ||
        besideDeletions.tables.before.has(id));
    if (tracked && besideDeletions.deleted.has(id) && blockBefore)
      return { type: "bounded", claimFrom: position };
    if (
      !besideDeletions.deleted.has(id) ||
      besideDeletions.insertedBeside.has(id) ||
      besideDeletions.tables.after.has(id) ||
      besideDeletions.tables.before.has(id)
    )
      return { type: deleted === 0 ? "none" : "kept" };
    deleted += 1;
  }
  return { type: deleted === 0 ? "none" : "runStart" };
};

/** The ids of the paragraphs `block` holds, at any depth. */
const paragraphIdsIn = (block: PMNode): string[] => {
  const ids: string[] = [];
  block.descendants((node) => {
    const id: unknown = node.attrs["paraId"];
    if (node.type.name === "paragraph" && typeof id === "string") ids.push(id);
    return node.type.name !== "paragraph";
  });
  return ids;
};

/**
 * The ids of the paragraphs a table's own cells hold, which name it: a
 * nested table's paragraphs name that one instead.
 */
const tableParagraphIds = (table: PMNode): string[] => {
  const ids: string[] = [];
  table.descendants((node) => {
    if (node.type.name === "table") return false;
    const id: unknown = node.attrs["paraId"];
    if (node.type.name === "paragraph" && typeof id === "string") ids.push(id);
    return node.type.name !== "paragraph";
  });
  return ids;
};

/** The blocks a batch's table insertions name, on `side` of them. */
const tableAnchors = (
  operations: readonly FolioDocumentOperation[],
  side: "after" | "before",
  anchorOf: (operation: TableInsertion) => string = ({ blockId }) => blockId,
): ReadonlySet<string> =>
  new Set(
    operations.flatMap((operation) =>
      (operation.type === "insertTable" ||
        operation.type === "insertSignatureTable") &&
      (operation.position ?? "after") === side
        ? [anchorOf(operation)]
        : [],
    ),
  );

const NO_BESIDE_DELETIONS: BesideDeletions = {
  deleted: new Set(),
  insertedBeside: new Set(),
  tables: { after: new Set(), before: new Set(), deleted: new Set() },
  paragraphsAfter: new Set(),
  landedEarlier: () => new Set(),
  placedBeside: () => [],
};

/**
 * The stated numbering a request names, minting its new list once for the
 * operation, as the legacy engine allocates one instance per operation.
 */
const statedNumbering = (
  numbering: FolioAIListNumbering | undefined,
  mint: ReturnType<typeof createBatchListMinter>,
): { numbering?: FolioContentStatedNumbering; newList?: NewListInstance } => {
  if (numbering === undefined) return {};
  if (numbering.kind !== "newList") return { numbering };
  const minted = mint(numbering);
  return {
    numbering: paragraphNumberingReference({
      numId: minted.numId,
      ilvl: numbering.level ?? 0,
    }),
    newList: minted.newList,
  };
};

/**
 * Whether the story's editor document holds the paragraph `blockId`. One the
 * session's projection does not address still exists: a text box's
 * paragraphs belong to its shape, not to the story, so no story operation
 * reaches them.
 */
const holdsParagraph = (doc: PMNode, blockId: string): boolean => {
  let found = false;
  doc.descendants((node) => {
    if (found) return false;
    if (node.type.name === "paragraph" && node.attrs["paraId"] === blockId)
      found = true;
    return !found && !node.isTextblock;
  });
  return found;
};

/** Where a block the source leaves empty stands in the story's document. */
type EmptyControlSpan = { before: number; after: number };

/**
 * The content control or custom XML element the source leaves empty whose
 * filler paragraph `blockId` names, or `undefined`. The filler has no source
 * and no id of its own: a snapshot names it by the positional id the legacy
 * engine mints for it (`ensureParaIdsInDoc`), and positions in that copy are
 * the document's.
 */
const emptyControlOfFiller = (
  doc: PMNode,
  blockId: string,
): EmptyControlSpan | undefined => {
  const minted = ensureParaIdsInDoc(doc);
  if (minted === doc) return undefined;
  let filler: number | undefined;
  minted.descendants((node, position) => {
    if (filler !== undefined) return false;
    if (node.type.name !== "paragraph") return true;
    if (node.attrs["paraId"] === blockId) filler = position;
    return false;
  });
  if (filler === undefined || doc.nodeAt(filler)?.attrs["paraId"] != null)
    return undefined;
  const $filler = doc.resolve(filler);
  const control = $filler.parent;
  if (
    (control.type.name !== "blockSdt" &&
      control.type.name !== "blockCustomXml") ||
    control.attrs["_originallyEmpty"] !== true
  )
    return undefined;
  return { before: $filler.before(), after: $filler.after() };
};

type BesideEmptyControlOptions = {
  operation: FolioDocumentOperation;
  control: EmptyControlSpan;
  projection: CanonicalProjection;
};

type BesideEmptyControl =
  | { type: "beside"; operation: FolioDocumentOperation }
  | {
      type: "listStart";
      operation: FolioDocumentOperation;
      listStart: ListStartPlace;
    }
  | { type: "refused"; message: string };

/** What stands on one side of a block, as an insertion there names it. */
type ControlSide =
  | { type: "paragraph"; blockId: string }
  | { type: "table"; blockId: string }
  | { type: "listStart"; listStart: ListStartPlace }
  | { type: "none" };

/**
 * An operation on the filler of a block the source leaves empty, restated
 * beside the block: the paragraph or table it inserts goes in the gap before
 * or after the block, named by the paragraph or table on that side of it,
 * or, before a block that starts its list, by the start of that list; the
 * block stays as the source wrote it. Inside a table the insertion goes
 * beside the outermost table, as beside any paragraph of a cell. The
 * filler's properties are none, so an inserted paragraph inherits none.
 * Every other operation would edit the filler, which holds nothing to edit.
 */
const besideEmptyControl = ({
  operation,
  control,
  projection,
}: BesideEmptyControlOptions): BesideEmptyControl => {
  const { doc } = projection;
  const tableAddress = (before: number) => {
    const table = doc.nodeAt(before);
    return table === null
      ? undefined
      : tableBlockAddress(projection, { table, tableStart: before + 1 });
  };
  const outer = findOutermostTableBoundary(doc, control.before);
  const sideOf = (side: "before" | "after"): ControlSide => {
    const $gap = doc.resolve(control[side]);
    const beside = side === "after" ? $gap.nodeAfter : $gap.nodeBefore;
    const at = side === "after" ? control.after : control.before;
    const besideAt = side === "after" ? at : at - (beside?.nodeSize ?? 0);
    if (beside?.type.name === "paragraph") {
      const paragraph = projection.paragraphStartingAt(besideAt);
      return paragraph === undefined
        ? { type: "none" }
        : { type: "paragraph", blockId: paragraph.blockId };
    }
    if (beside?.type.name === "table") {
      const blockId = tableAddress(besideAt);
      return blockId === undefined
        ? { type: "none" }
        : { type: "table", blockId };
    }
    if (beside !== null) return { type: "none" };
    // The block starts its list: the list is named by a paragraph it holds,
    // and every list ends with one.
    let named: string | undefined;
    $gap.parent.forEach((child, offset) => {
      if (named !== undefined || child.type.name !== "paragraph") return;
      named = projection.paragraphStartingAt($gap.start() + offset)?.blockId;
    });
    return named === undefined
      ? { type: "none" }
      : { type: "listStart", listStart: { blockId: named, position: at } };
  };
  const noNeighbour = {
    type: "refused",
    message:
      "Canonical insertion beside an empty content control needs a paragraph or table on that side of it.",
  } as const;
  type Placement = {
    side: "before" | "after";
    beside: (
      blockId: string,
      side: "before" | "after",
    ) => FolioDocumentOperation;
  };
  // The insertion named by what stands on `side` of the control, put on the
  // side of it that faces the control.
  const placed = ({ side, beside }: Placement): BesideEmptyControl => {
    const facing = side === "after" ? "before" : "after";
    if (outer !== null) {
      const blockId = tableAddress(outer.before);
      return blockId === undefined
        ? noNeighbour
        : { type: "beside", operation: beside(blockId, side) };
    }
    const standing = sideOf(side);
    switch (standing.type) {
      case "paragraph":
      case "table":
        return {
          type: "beside",
          operation: beside(standing.blockId, facing),
        };
      case "listStart":
        return {
          type: "listStart",
          operation: beside(standing.listStart.blockId, side),
          listStart: standing.listStart,
        };
      case "none":
        return noNeighbour;
      default: {
        const unreachable: never = standing;
        return unreachable;
      }
    }
  };
  switch (operation.type) {
    case "insertAfterBlock":
    case "insertBeforeBlock": {
      const { precondition: _precondition, ...rest } = operation;
      return placed({
        side: operation.type === "insertAfterBlock" ? "after" : "before",
        beside: (blockId, side) => ({
          ...rest,
          type: side === "after" ? "insertAfterBlock" : "insertBeforeBlock",
          blockId,
          inheritFormatting: false,
        }),
      });
    }
    case "insertTable":
    case "insertSignatureTable": {
      const { precondition: _precondition, ...rest } = operation;
      return placed({
        side: operation.position ?? "after",
        beside: (blockId, side) => ({ ...rest, blockId, position: side }),
      });
    }
    case "replaceInBlock":
    case "replaceRange":
    case "replaceBlock":
    case "splitBlock":
    case "formatRange":
    case "mergeBlockWithNext":
    case "commentOnRange":
    case "commentOnBlock":
    case "deleteBlock":
    case "setBlockParagraphProperties":
    case "deleteTable":
    case "insertTableRow":
    case "deleteTableRow":
    case "insertTableColumn":
    case "deleteTableColumn":
    case "mergeTableCells":
    case "splitTableCell":
      return {
        type: "refused",
        message:
          "The paragraph shown in an empty content control is not part of the document; insert a block before or after the control instead.",
      };
    default: {
      const unreachable: never = operation;
      return unreachable;
    }
  }
};

const canonicalPublicStory = (story: FolioDocumentOperationStory): OpStory => {
  if (story === "main") return story;
  switch (story.type) {
    case "header":
    case "footer":
      return { kind: story.type, rId: story.relationshipId };
    case "footnote":
    case "endnote":
      return { kind: story.type, id: story.noteId };
    default: {
      const unreachable: never = story;
      return panic(`Unknown public operation story ${unreachable}`);
    }
  }
};

const batchStatus = (
  batch: FolioDocumentOperationBatch,
  skippedCount: number,
) => {
  if (batch.dryRun) return "previewed" as const;
  if (batch.atomic && skippedCount > 0) return "rejected" as const;
  return "committed" as const;
};

/** Handles address isolated journal groups; no second document or PM history is retained. */
export class CanonicalPublicOperations {
  private readonly options: CanonicalPublicOperationsOptions;
  private readonly handles: {
    handle: FolioDocumentOperationUndoHandle;
    /** The session version the batch's undo expects. */
    version: number;
    /** The session version the batch started from. */
    startVersion: number;
    story: OpStory;
  }[] = [];
  private nextHandle = 1;
  private readonly handleNamespace = crypto.randomUUID();

  constructor(options: CanonicalPublicOperationsOptions) {
    this.options = options;
  }

  apply(
    options: CanonicalPublicOperationOptions,
  ): FolioDocumentOperationResult {
    const { session } = this.options;
    const parsed = parseFolioDocumentOperationBatch(options.batch);
    if (options.restagesSuggestions === true && parsed.mode !== "suggested")
      panic("Only a suggested batch restages pending suggestions.");
    const lockAdmission =
      options.restagesSuggestions === true
        ? CANONICAL_ADMISSION.restage
        : CANONICAL_ADMISSION.edit;
    // A suggested operation never folds into a suggestion already pending,
    // here or in the host's other stories.
    const batch =
      parsed.mode !== "suggested" ||
      options.restagesSuggestions === true ||
      session.isComposing
        ? parsed
        : withDistinctSuggestionIds(
            parsed,
            new Set([
              ...pendingSuggestionIdsIn(readModel(session.document)),
              ...(options.pendingSuggestionIds ?? []),
            ]),
          );
    const publicStory = options.story ?? "main";
    const story = canonicalPublicStory(publicStory);
    const state = this.options.getState(story);
    const projected = session.projectStory(story);
    const skipped: FolioAIEditSkippedOperation[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    const refusals = new Map<string, CanonicalPublicOperationRefusal>();
    const resolved: (PreparedOperation & { plan: CanonicalBatchPlan })[] = [];
    // Read once, at the first operation that reaches the check: a batch the
    // session refuses as a whole reads no document.
    let definedReferences:
      | {
          styles: ReturnType<typeof styleDefinitionLookup>;
          definitions: StyleDefinitions | undefined;
          numberingInstanceIds: ReadonlySet<number>;
        }
      | undefined;
    const references = () => {
      if (definedReferences !== undefined) return definedReferences;
      const model = readModel(session.document);
      definedReferences = {
        definitions: model.package.styles,
        styles: styleDefinitionLookup(
          createStyleResolver(model.package.styles),
        ),
        numberingInstanceIds: new Set(
          model.package.numbering?.nums.map(({ numId }) => numId) ?? [],
        ),
      };
      return definedReferences;
    };
    const headingStyles = createOperationHeadingStyles();
    const headingLookups = new WeakMap<
      StyleDefinitions,
      ReturnType<typeof styleDefinitionLookup>
    >();
    const headingLookup = (definitions: StyleDefinitions) => {
      const existing = headingLookups.get(definitions);
      if (existing !== undefined) return existing;
      const lookup = styleDefinitionLookup(createStyleResolver(definitions));
      headingLookups.set(definitions, lookup);
      return lookup;
    };
    const mintList = createBatchListMinter(() => readModel(session.document));
    const { diff } = wordDiffSessionFromOptions(options.wordDiff);
    let nextRevisionId = options.revisionStamp?.idSeed ?? 1;
    let batchLease = session.idLease;
    if (!session.isComposing && options.revisionStamp === undefined) {
      // Past every revision id the package holds, records the projection
      // does not show (a table's grid change, say) included.
      const taken = new Set(
        leasedIdCensus(readModel(session.document)).revision,
      );
      for (const currentStory of documentStories(readModel(session.document))) {
        const currentProjection = session.projectStory(currentStory);
        if (currentProjection.isOk())
          forEachAnnotationIdInDoc(currentProjection.value.doc, (id) =>
            taken.add(id),
          );
      }
      for (const id of taken) nextRevisionId = Math.max(nextRevisionId, id + 1);
      // With little room left past the highest id, the batch draws its
      // revision ids from the lowest ones no record or annotation holds, as
      // a lease does, so a package holding the largest id still takes edits.
      if (
        batchLease === undefined &&
        MAX_REVISION_ID + 1 - nextRevisionId < REVISION_ID_WRAP_HEADROOM
      )
        batchLease = wrappedRevisionLease(taken);
    }
    const skip = (
      id: string,
      reason: FolioAIEditSkipReason,
      message?: string,
    ) => {
      skipped.push({
        id,
        reason,
        ...(message === undefined ? {} : { message }),
      });
    };
    // A refused resolution reports the ground it names, if any.
    const refuseResolution = (
      id: string,
      error: CanonicalPublicOperationResolutionError,
    ) => {
      if (error.ground !== undefined) refusals.set(id, error.ground);
      skip(id, error.reason, error.detail);
    };
    const result = (
      applied: FolioDocumentOperationResult["applied"],
      undoHandle: FolioDocumentOperationUndoHandle | null,
      disposition: "normal" | "refused" = "normal",
    ): FolioDocumentOperationResult => {
      const failures = new Map(skipped.map((item) => [item.id, item]));
      const ordered = batch.operations.flatMap(({ id }) => {
        const failure = failures.get(id);
        if (!failure) return [];
        const canonicalRefusal = refusals.get(id);
        return [canonicalRefusal ? { ...failure, canonicalRefusal } : failure];
      });
      return {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        status:
          disposition === "refused" && !batch.dryRun
            ? "rejected"
            : batchStatus(batch, ordered.length),
        applied,
        skipped: ordered,
        issues: getFolioDocumentOperationIssues(batch.operations, ordered),
        receipts: getFolioDocumentOperationReceiptsForStory({
          operations: batch.operations,
          applied,
          story: publicStory,
        }),
        undoHandle,
        nextRevisionId,
        ...(normalizations.length === 0 ? {} : { normalizations }),
      };
    };
    const batchComments = batch.operations.filter(
      (operation): operation is BatchCommentOperation =>
        operation.type === "commentOnBlock" ||
        operation.type === "commentOnRange",
    );
    let batchNoteReferences: NoteReferenceLabels | undefined;
    const placeOf = (operation: BlockInsertion | TableInsertion) =>
      state && projected.isOk()
        ? landingParagraph(
            operation,
            projected.value,
            (batchNoteReferences ??= collectNoteReferenceLabels(state.doc)),
          )
        : undefined;
    const landsBeside = (operation: BlockInsertion | TableInsertion) =>
      placeOf(operation)?.blockId ?? operation.blockId;
    /**
     * The block boundary a paragraph the operation puts lands at: put
     * before a paragraph, at the start of its gap (`insertionGap`).
     */
    const landingOf = (operation: FolioDocumentOperation): number[] => {
      const place = isBlockInsertion(operation)
        ? placeOf(operation)
        : undefined;
      if (place === undefined || !projected.isOk()) return [];
      return [
        operation.type === "insertAfterBlock"
          ? place.start - 1 + place.node.nodeSize
          : insertionGap(projected.value.doc, place.start - 1, "before"),
      ];
    };
    const besideDeletions: BesideDeletions = {
      deleted: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "deleteBlock" ? [operation.blockId] : [],
        ),
      ),
      insertedBeside: new Set(
        batch.operations.flatMap((operation) =>
          isBlockInsertion(operation) ? [landsBeside(operation)] : [],
        ),
      ),
      tables: {
        after: tableAnchors(batch.operations, "after", landsBeside),
        before: tableAnchors(batch.operations, "before", landsBeside),
        deleted: new Set(
          batch.operations.flatMap((operation) =>
            operation.type === "deleteTable" ? [operation.blockId] : [],
          ),
        ),
      },
      paragraphsAfter: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "insertAfterBlock" ? [landsBeside(operation)] : [],
        ),
      ),
      landedEarlier: (operationId) => {
        const index = batch.operations.findIndex(
          ({ id }) => id === operationId,
        );
        return new Set(
          batch.operations.slice(0, Math.max(index, 0)).flatMap(landingOf),
        );
      },
      placedBeside: (blockId) =>
        batch.operations.filter(
          (operation): operation is BlockInsertion | TableInsertion =>
            placesBlock(operation) && landsBeside(operation) === blockId,
        ),
    };
    const inputIndices = new Map(
      batch.operations.map(({ id }, index) => [id, index]),
    );
    const planned = (item: PreparedOperation) => {
      if (projected.isErr())
        return panic("A resolved batch lost its read projection");
      const inputIndex =
        inputIndices.get(item.operation.id) ??
        panic("A resolved operation lost its input position");
      return {
        ...item,
        plan: canonicalBatchPlan({
          ...item,
          projection: projected.value,
          tracked: batch.mode !== "direct",
          inputIndex,
        }),
      };
    };
    let claims = new BatchClaims();
    const conflictsWithRevision =
      batch.mode !== "direct" &&
      batch.existingRevisions === "refuse-cross-author" &&
      projected.isOk()
        ? createExistingRevisionGuard({
            doc: projected.value.doc,
            author: options.author ?? "AI",
          })
        : undefined;
    const refusesExistingRevision = ({
      operation,
      plan,
    }: ReturnType<typeof planned>) => {
      if (
        operation.type === "commentOnBlock" ||
        operation.type === "commentOnRange" ||
        !conflictsWithRevision?.(plan.claim)
      ) {
        return false;
      }
      skip(
        operation.id,
        "existingRevisionConflict",
        EXISTING_REVISION_CONFLICT_MESSAGE,
      );
      return true;
    };
    /** Skip a candidate an earlier operation of the batch already claims. */
    const skipsOverlap = (candidate: ReturnType<typeof planned>): boolean => {
      const claimedBy = claims.conflictOf(candidate.plan.claim);
      if (claimedBy === null) return false;
      skip(
        candidate.operation.id,
        "overlappingOperation",
        `operation ${JSON.stringify(claimedBy)}, earlier in this batch, already claims its target.`,
      );
      return true;
    };
    const removedLines = new Set<string>();
    /** Skip a row or column deletion an earlier one of the batch already makes. */
    const skipsRepeatedRemoval = (
      candidate: ReturnType<typeof planned>,
    ): boolean => {
      const { removes } = candidate.plan;
      if (removes === null || !removedLines.has(tableLineRemovalKey(removes)))
        return false;
      skip(candidate.operation.id, "noopOperation");
      return true;
    };
    const admit = (candidate: ReturnType<typeof planned>) => {
      claims.add(candidate.operation.id, candidate.plan.claim);
      const { removes } = candidate.plan;
      if (removes !== null) removedLines.add(tableLineRemovalKey(removes));
      resolved.push(candidate);
    };
    /** Refused, but claiming the target their direct form would. */
    const claimedUnapplied: ReturnType<typeof planned>[] = [];
    for (const operation of batch.operations) {
      if (!state || projected.isErr() || session.isComposing) {
        skip(operation.id, "documentNotEditable");
        continue;
      }
      if (
        batch.mode === "suggested" &&
        !CANONICAL_SUGGESTED_OPERATION_TYPES.has(operation.type)
      ) {
        skip(operation.id, "unsupportedMode");
        continue;
      }
      // The same reference check the legacy engine runs, in the same
      // place: before the operation compiles.
      const referenceSet = references();
      // Kept references stand as they are spelled: a heading the document
      // does not define stays undefined rather than gaining a definition.
      const headings =
        options.undefinedReferences === "keep"
          ? referenceSet.definitions
          : headingStyles(referenceSet.definitions, operation);
      const undefinedReference = findUndefinedOperationReference({
        operation,
        numberingInstanceIds: referenceSet.numberingInstanceIds,
        styles:
          headings === undefined || headings === referenceSet.definitions
            ? referenceSet.styles
            : headingLookup(headings),
        undefinedReferences: options.undefinedReferences,
      });
      if (undefinedReference !== undefined) {
        skip(
          operation.id,
          undefinedReference.reason,
          undefinedReference.message,
        );
        continue;
      }
      const compiled = this.resolve({
        operation,
        options,
        state,
        story,
        diff,
        tracked: batch.mode !== "direct",
        suggested: batch.mode === "suggested",
        mintList,
        besideDeletions,
        batchComments,
      });
      if (compiled.isErr()) {
        // An operation only the tracked vocabulary refuses still claims what
        // it would edit directly, so whether a later operation of the batch
        // overlaps it does not depend on the mode.
        const direct =
          compiled.error.refusedIn === "trackedMode"
            ? this.resolve({
                operation,
                options,
                state,
                story,
                diff,
                tracked: false,
                suggested: false,
                mintList,
                besideDeletions,
                batchComments,
              })
            : undefined;
        if (direct?.isOk()) {
          const candidate = planned(direct.value);
          if (skipsRepeatedRemoval(candidate)) continue;
          if (refusesExistingRevision(candidate)) continue;
          if (skipsOverlap(candidate)) continue;
          claims.add(operation.id, candidate.plan.claim);
          claimedUnapplied.push(candidate);
        }
        refuseResolution(operation.id, compiled.error);
        continue;
      }
      const candidate = planned(compiled.value);
      if (skipsRepeatedRemoval(candidate)) continue;
      if (refusesExistingRevision(candidate)) continue;
      if (skipsOverlap(candidate)) continue;
      admit(candidate);
    }
    // An operation resolved beside a deleted paragraph that ends its run
    // counts on its other half; where that half did not resolve, it resolves
    // again on its own, against the operations that did.
    for (const item of [...resolved]) {
      if (item.besideDeletion !== true || !state) continue;
      // As requested: resolved beside the deletion, an insertion before the
      // paragraph may have been placed after it instead.
      const operation =
        batch.operations.find(({ id }) => id === item.operation.id) ??
        panic("A resolved operation lost its request.");
      if (
        operation.type !== "deleteBlock" &&
        operation.type !== "insertAfterBlock" &&
        operation.type !== "insertBeforeBlock" &&
        operation.type !== "insertTable" &&
        operation.type !== "insertSignatureTable"
      )
        continue;
      const partnered = resolved.some(({ operation: other }) =>
        operation.type === "deleteBlock"
          ? (other.type === "insertAfterBlock" ||
              other.type === "insertBeforeBlock") &&
            other.blockId === operation.blockId
          : // One moved past the deletion names the paragraph it deletes.
            other.type === "deleteBlock" &&
            "blockId" in item.operation &&
            other.blockId === item.operation.blockId,
      );
      if (partnered) continue;
      resolved.splice(resolved.indexOf(item), 1);
      // Conditional resolution may retire an accepted claim; rebuild only at
      // this boundary, so its old target cannot reject the independent retry.
      claims = new BatchClaims();
      for (const earlier of [...resolved, ...claimedUnapplied])
        claims.add(earlier.operation.id, earlier.plan.claim);
      const alone = this.resolve({
        operation,
        options,
        state,
        story,
        diff,
        tracked: batch.mode !== "direct",
        suggested: batch.mode === "suggested",
        mintList,
        besideDeletions: NO_BESIDE_DELETIONS,
        batchComments,
      });
      if (alone.isErr()) {
        refuseResolution(operation.id, alone.error);
        continue;
      }
      const candidate = planned(alone.value);
      if (refusesExistingRevision(candidate)) continue;
      if (skipsOverlap(candidate)) continue;
      admit(candidate);
    }
    // A move id names one relocation, a move only when it reaches exactly one
    // source and one destination, as the legacy engine pairs them; an
    // unpaired id records a plain insertion or deletion and is reported.
    const moveSides = new Map<string, { from: number; to: number }>();
    for (const { operation } of resolved) {
      const moveId = blockMoveId(operation);
      if (moveId === undefined) continue;
      const sides = moveSides.get(moveId) ?? { from: 0, to: 0 };
      if (operation.type === "deleteBlock") sides.from += 1;
      else sides.to += 1;
      moveSides.set(moveId, sides);
    }
    // A suggested destination before its own source would cut the move.
    const unpairable = new Set(
      batch.mode === "suggested"
        ? resolved.flatMap(({ operation }) =>
            operation.type === "insertBeforeBlock" &&
            operation.moveId !== undefined &&
            resolved.some(
              (candidate) =>
                candidate.operation.type === "deleteBlock" &&
                candidate.operation.moveId === operation.moveId &&
                candidate.operation.blockId === operation.blockId,
            )
              ? [operation.moveId]
              : [],
          )
        : [],
    );
    const paired = (moveId: string) =>
      !unpairable.has(moveId) &&
      moveSides.get(moveId)?.from === 1 &&
      moveSides.get(moveId)?.to === 1;
    const pairedMoves = new Set<string>();
    for (const [index, item] of resolved.entries()) {
      const moveId = blockMoveId(item.operation);
      if (moveId === undefined) continue;
      if (!paired(moveId)) {
        normalizations.push({
          id: item.operation.id,
          code: "unpairedMove",
          moveId,
        });
        continue;
      }
      pairedMoves.add(moveId);
      const intents: EditorIntent[] = [];
      for (const intent of item.intents)
        intents.push(
          intent.type === "insertParagraphs" ||
            intent.type === "deleteParagraphs"
            ? Object.assign({}, intent, { move: true as const })
            : intent,
        );
      resolved[index] = { ...item, intents };
    }
    // What an operation only the tracked vocabulary refuses would edit directly
    // counts here too, so whether another is refused does not depend on the
    // mode.
    const shaping = [...resolved, ...claimedUnapplied].toSorted(
      (left, right) =>
        (inputIndices.get(left.operation.id) ?? 0) -
        (inputIndices.get(right.operation.id) ?? 0),
    );
    const refuseShapes = (
      items: readonly ReturnType<typeof planned>[],
      message: string,
    ): void => {
      for (const item of items) {
        shaping.splice(shaping.indexOf(item), 1);
        const index = resolved.indexOf(item);
        if (index < 0) continue;
        resolved.splice(index, 1);
        skip(item.operation.id, "unsupportedBlock", message);
      }
    };
    refuseShapes(rowShiftedCellShapes(shaping), ROW_SHIFTED_CELL_SHAPE);
    refuseShapes(columnEditedCellShapes(shaping), COLUMN_EDITED_CELL_SHAPE);
    if (projected.isOk())
      refuseShapes(
        rowClosingCellShapes({ items: shaping, doc: projected.value.doc }),
        ROW_CLOSING_CELL_SHAPE,
      );
    refuseShapes(overlappingCellShapes(shaping), OVERLAPPING_CELL_SHAPE);
    if (batch.atomic && skipped.length > 0) {
      for (const { operation } of resolved)
        skip(operation.id, "atomicBatchRejected");
      return result([], null);
    }
    if (!state || projected.isErr() || session.isComposing)
      return result([], null);
    const executionOrder = orderCanonicalBatch({
      items: resolved,
      doc: projected.value.doc,
      mode: batch.mode ?? "tracked-changes",
    });
    let document = readModel(session.document);
    // The batch's comments run before its edits, each adding a reference
    // unit; the document once they stand maps the edits' pre-batch positions.
    const beforeBatch = document;
    let annotated: Document | undefined;
    const ops: DocumentOp[] = [];
    const operationStaging = session.beginVerifiedOperationStaging(
      canonicalOperationStagingAccess,
      lockAdmission,
    );
    let reuseStaging = true;
    const appliedById = new Map<
      string,
      FolioDocumentOperationResult["applied"][number]
    >();
    const allocate = createEditorIntentIdAllocator({
      reservedBlockIds: session.reservedBlockIds,
    });
    const date = options.revisionStamp?.date ?? new Date().toISOString();
    type AdmittedOperation = (typeof resolved)[number];
    const insertedColumnCounts = new Map<string, number>();
    const recordColumnInsertion = (
      member: AdmittedOperation,
      counts: Map<string, number>,
    ) => {
      const column = member.plan.address;
      if (column.type !== "column" || !column.insertion) return;
      const key = getTableColumnCoordinateKey(column);
      const added =
        member.operation.type === "insertTableColumn"
          ? (member.operation.columnCount ?? 1)
          : 1;
      counts.set(key, (counts.get(key) ?? 0) + added);
    };
    type PrepareOperationOptions = {
      document: Document;
      resolvedOperation: AdmittedOperation;
      columnInsertions: ReadonlyMap<string, number>;
      nextRevisionId: number;
      idLease?: IdLease;
      paragraphAllocationLease?: IdLease;
      allocations?: CompilePublicIntentsOptions["allocations"];
      reservedBlockIds?: ReadonlySet<string>;
      verifyOperation?: PublicOperationStaging["apply"];
      commentAllocation?: {
        id: number;
        content: ReturnType<typeof canonicalCommentBody>;
      };
    };
    type DeclinedOperationOptions = {
      reason: FolioAIEditSkipReason;
      message?: string | undefined;
      refusal?: CanonicalPublicOperationRefusal | undefined;
    };
    const declined = ({
      reason,
      message,
      refusal,
    }: DeclinedOperationOptions) => ({
      type: "declined" as const,
      reason,
      message,
      refusal,
    });
    /**
     * The paragraph each paragraph an earlier operation of the batch joined
     * away now lives in: an edit of its properties edits that one, as it
     * does one at a time after that operation.
     */
    const survivorOf = new Map<string, string>();
    /** Those whose words open the paragraph they now live in. */
    const openingSurvivorOf = new Map<string, string>();
    const deletionSurvivorOf = new Map<string, string>();
    // Preparation owns no session state: even comments and lock refusals are
    // facts until the batch chooses this attempt.
    const prepareOperation = ({
      document: sourceDocument,
      resolvedOperation,
      columnInsertions,
      nextRevisionId: firstRevisionId,
      idLease = batchLease,
      paragraphAllocationLease,
      allocations,
      reservedBlockIds = session.reservedBlockIds,
      verifyOperation,
      commentAllocation: admittedComment,
    }: PrepareOperationOptions) => {
      const { operation } = resolvedOperation;
      // An author no save accepts is refused before it is recorded: on a
      // tracked change, or on the comment the operation makes.
      const authorIssue =
        (batch.mode === "direct"
          ? undefined
          : trackedChangeAuthorIssue(options.author ?? "AI")) ??
        (resolvedOperation.type === "commented"
          ? commentAuthorIssue(options.author ?? "AI")
          : undefined);
      if (authorIssue !== undefined)
        return declined({ reason: "invalidResult", message: authorIssue });
      const inserting = formattingIntentsAsItRuns({
        document: sourceDocument,
        story,
        survivorOf,
        openingSurvivorOf,
        deletionSurvivorOf,
        intents: insertionIntentsAsItRuns({
          document: sourceDocument,
          story,
          survivorOf,
          resolvedOperation,
        }),
      });
      const shift = batchCommentReferenceShift({
        before: beforeBatch,
        commented:
          resolvedOperation.plan.address.type === "annotation"
            ? sourceDocument
            : (annotated ?? sourceDocument),
      });
      // Pre-batch positions first take the batch's comment references; the
      // as-it-runs bindings below read the document as it stands, past them.
      const shifted = inserting.map((intent) =>
        mapIntentPositions(intent, shift),
      );
      const rebound =
        resolvedOperation.joinsParagraphBefore === true
          ? deletionIntentsAsItRuns(sourceDocument, story, shifted)
          : shifted;
      const retired =
        resolvedOperation.retiresBreakBefore === true
          ? retiredBreakAsItRuns(sourceDocument, story, rebound)
          : rebound;
      const deleting =
        resolvedOperation.mayEndContainer === true
          ? finalDeletionIntentsAsItRuns(sourceDocument, story, retired)
          : retired;
      // A suggested operation's records are one suggestion: its own id, or the
      // one it names.
      const suggestionId =
        batch.mode === "suggested"
          ? (("suggestionId" in operation
              ? operation.suggestionId
              : undefined) ?? operation.id)
          : undefined;
      const retracting =
        operation.type === "deleteBlock" &&
        batch.mode !== "direct" &&
        operation.moveId === undefined &&
        resolvedOperation.keepsTablesApart !== true
          ? takenOverBreakRetractionAsItRuns({
              document: sourceDocument,
              story,
              blockId: operation.blockId,
              author: options.author ?? "AI",
              suggestionId,
              intents: addedMarkBeforeTableAsItRuns({
                document: sourceDocument,
                story,
                blockId: operation.blockId,
                author: options.author ?? "AI",
                gainsFollower: (id) =>
                  besideDeletions.tables.after.has(id) ||
                  besideDeletions.paragraphsAfter.has(id),
                intents: deleting,
              }),
            })
          : deleting;
      const joining =
        operation.type === "mergeBlockWithNext"
          ? mergeIntentsAsItRuns({
              document: sourceDocument,
              story,
              intents: retracting,
              mode: batch.mode,
            })
          : retracting;
      if (joining === undefined)
        return declined({
          reason: "missingBlock",
          message:
            "No paragraph follows the paragraph to join once the batch's deletions are made.",
        });
      const intents = joining;
      const livePropertyRefusal = batchPropertyRefusal({
        document: sourceDocument,
        story,
        operation,
      });
      if (livePropertyRefusal !== undefined)
        return declined({
          reason: livePropertyRefusal.reason,
          message: livePropertyRefusal.message,
          refusal: livePropertyRefusal.canonicalRefusal,
        });
      const column = resolvedOperation.plan.address;
      const columnIntents =
        column.type === "column" && operation.type === "deleteTableColumn"
          ? mapColumnDeletionIntents({
              intents,
              columnIndex: column.columnIndex,
              inserted:
                columnInsertions.get(getTableColumnCoordinateKey(column)) ?? 0,
              insertedAfter:
                columnInsertions.get(
                  getTableColumnCoordinateKey({
                    ...column,
                    columnIndex: column.columnIndex + 1,
                  }),
                ) ?? 0,
            })
          : intents;
      const mappedIntents = tableIntentsAsItRuns({
        document: sourceDocument,
        story,
        projection: projected.value,
        plan: resolvedOperation.plan,
        intents: columnIntents,
      });
      const headings =
        options.undefinedReferences === "keep"
          ? sourceDocument.package.styles
          : headingStyles(sourceDocument.package.styles, operation);
      const withHeadings =
        headings === undefined || headings === sourceDocument.package.styles
          ? sourceDocument
          : {
              ...sourceDocument,
              package: { ...sourceDocument.package, styles: headings },
            };
      const compiled = compilePublicIntents({
        document: withHeadings,
        story,
        intents: mappedIntents,
        allocate,
        reservedBlockIds,
        idLease: idLease,
        ...(paragraphAllocationLease === undefined
          ? {}
          : { paragraphAllocationLease }),
        nextRevisionId: firstRevisionId,
        mode: batch.mode === "direct" ? "direct" : "tracked-changes",
        suggestionId,
        author: options.author ?? "AI",
        date,
        ...(allocations === undefined ? {} : { allocations }),
        ...(withHeadings === sourceDocument &&
        resolvedOperation.type === "plain" &&
        verifyOperation !== undefined
          ? { verifyOperation }
          : {}),
      });
      if (compiled.isErr())
        return declined(compilerRefusalSkip(compiled.error));
      let nextDocument = compiled.value.document;
      const operationOps = [...compiled.value.ops];
      if (
        operation.type === "mergeBlockWithNext" &&
        conflictsWithRevision !== undefined
      ) {
        if (
          overwritesForeignMergeFormatting({
            before: sourceDocument,
            after: nextDocument,
            story,
            ops: operationOps,
            author: options.author ?? "AI",
          })
        ) {
          return declined({
            reason: "existingRevisionConflict",
            message: EXISTING_REVISION_CONFLICT_MESSAGE,
          });
        }
      }
      let commentId: number | undefined;
      let commentAllocation: PrepareOperationOptions["commentAllocation"];
      // A lease the batch takes only for its revision ids leaves comments to
      // their own allocator.
      const commentLease = session.idLease === undefined ? undefined : idLease;
      switch (resolvedOperation.type) {
        case "plain":
          break;
        case "commented": {
          const { comment } = resolvedOperation;
          const id =
            admittedComment === undefined
              ? freshCommentId(nextDocument, commentLease)
              : Result.ok(admittedComment.id);
          // A suggested operation's comment is proposed with it: pending,
          // never saved, accepted and rejected with the suggestion.
          const proposedBy = suggestionId === undefined ? {} : { suggestionId };
          if (comment.inserted === true) {
            const { firstAdded, lastAdded } = insertedCommentRange({
              before: sourceDocument,
              after: nextDocument,
              story,
              ops: compiled.value.ops,
            });
            if (id.isErr())
              return declined({
                reason: "unsupportedBlock",
                message: id.error.message,
                refusal: groundForDocumentOpRefusal(id.error.reason),
              });
            const commentContent =
              admittedComment?.content ??
              canonicalCommentBody(nextDocument, comment.text, commentLease);
            // Where the insertion added no paragraph a range can name, the
            // comment marks the place it went in instead.
            const commented = compileCanonicalComments({
              document: nextDocument,
              idLease: idLease,
              command: {
                type: "create",
                comment: {
                  id: id.value,
                  author: options.author ?? "AI",
                  date,
                  done: false,
                  content: commentContent,
                  ...proposedBy,
                },
                anchor:
                  firstAdded?.paraId === undefined ||
                  lastAdded?.paraId === undefined
                    ? { kind: "point", at: comment.to }
                    : {
                        kind: "range",
                        from: {
                          story,
                          blockId: firstAdded.paraId,
                          offset: 0,
                        },
                        to: {
                          story,
                          blockId: lastAdded.paraId,
                          offset: paragraphLength(lastAdded),
                        },
                      },
              },
            });
            if (commented.isErr())
              return declined({
                reason: "unsupportedBlock",
                message: commented.error.message,
                refusal: commented.error.ground,
              });
            nextDocument = commented.value.document;
            operationOps.push(...commented.value.ops);
            commentId = id.value;
            commentAllocation = { id: id.value, content: commentContent };
            break;
          }
          // The anchor paragraph at any depth: in a table cell or a content
          // control as much as in the story's own block list.
          const anchorIn = (current: Document) =>
            findCommentAnchorParagraph({
              document: current,
              story,
              blockId: comment.from.blockId,
            });
          const before = anchorIn(sourceDocument);
          const after = anchorIn(nextDocument);
          if (id.isErr())
            return declined({
              reason: "unsupportedBlock",
              message: id.error.message,
              refusal: groundForDocumentOpRefusal(id.error.reason),
            });
          // An earlier change of the batch took the paragraph the comment
          // was read against.
          if (before === undefined || after === undefined)
            return declined({
              reason: "missingBlock",
              message: "The comment anchor no longer exists.",
              refusal: STALE,
            });
          // A tracked edit's comment covers what the edit changed, from its
          // first revision to its last, as the legacy engine anchors it.
          const changed =
            batch.mode === "direct"
              ? undefined
              : revisionSpan(after.content, new Set(compiled.value.revisions));
          const from =
            changed === undefined
              ? shift(comment.from)
              : {
                  story: comment.from.story,
                  blockId: comment.from.blockId,
                  offset: changed.from,
                };
          const to =
            changed === undefined
              ? {
                  ...comment.to,
                  offset:
                    shift(comment.to).offset +
                    paragraphLength(after) -
                    paragraphLength(before),
                }
              : {
                  story: comment.from.story,
                  blockId: comment.from.blockId,
                  offset: changed.to,
                };
          const commentContent =
            admittedComment?.content ??
            canonicalCommentBody(nextDocument, comment.text, commentLease);
          const commented = compileCanonicalComments({
            document: nextDocument,
            idLease: idLease,
            command: {
              type: "create",
              comment: {
                id: id.value,
                author: options.author ?? "AI",
                date,
                done: false,
                content: commentContent,
                ...proposedBy,
              },
              anchor:
                to.offset === from.offset
                  ? { kind: "point", at: from }
                  : { kind: "range", from, to },
            },
          });
          if (commented.isErr())
            return declined({
              reason: "unsupportedBlock",
              message: commented.error.message,
              refusal: commented.error.ground,
            });
          nextDocument = commented.value.document;
          operationOps.push(...commented.value.ops);
          commentId = id.value;
          commentAllocation = { id: id.value, content: commentContent };
          break;
        }
        default: {
          const unreachable: never = resolvedOperation;
          panic(`Unknown resolved operation ${unreachable}`);
        }
      }
      if (operationOps.length === 0) {
        return declined({ reason: "noopOperation" });
      }
      // An operation a content-control lock refuses is skipped on its own,
      // as the document's rule; the rest of the batch still applies.
      const locked =
        lockAdmission === CANONICAL_ADMISSION.restage
          ? null
          : lockViolationBetween(sourceDocument, nextDocument);
      if (locked !== null) {
        return declined({
          reason: CONTENT_CONTROL_LOCK,
          message: `Refused: ${lockedControlMessage(locked)}`,
        });
      }
      return {
        type: "prepared" as const,
        retiredCommentIds:
          batch.mode === "direct"
            ? []
            : commentsRetiredBetween(sourceDocument, nextDocument),
        // The intents as they ran, which name what this operation joined.
        intents: mappedIntents,
        document: nextDocument,
        ops:
          withHeadings === sourceDocument
            ? operationOps
            : [
                packageResourcesOpOf({
                  before: sourceDocument,
                  after: withHeadings,
                }),
                ...operationOps,
              ],
        nextRevisionId: Math.max(
          compiled.value.nextRevisionId,
          (commentId ?? -1) + 1,
        ),
        // A fold into a pending property change shares that change's id.
        revisionIds: [...compiled.value.revisions, ...compiled.value.folded],
        absorbed: compiled.value.absorbed,
        commentId,
        suggestionId,
        allocations: compiled.value.allocations,
        commentAllocation,
        verifiedOperation: compiled.value.verifiedOperation,
      };
    };
    type OperationAttempt = ReturnType<typeof prepareOperation>;
    /** The operations that resolved but were declined as they ran. */
    const declinedAsRun = new Set<string>();
    const commitOperation = (
      resolvedOperation: AdmittedOperation,
      prepared: OperationAttempt,
    ) => {
      const { operation, normalizations: adjustments } = resolvedOperation;
      if (prepared.type === "declined") {
        if (prepared.refusal !== undefined)
          refusals.set(operation.id, prepared.refusal);
        skip(operation.id, prepared.reason, prepared.message);
        declinedAsRun.add(operation.id);
        return;
      }
      const survivorMaps =
        operation.type === "deleteBlock"
          ? [survivorOf, deletionSurvivorOf]
          : [survivorOf];
      for (const absorption of prepared.absorbed) {
        const { absorbed, survivor } = absorption;
        const opens = opensSurvivor({ document, story, absorption });
        for (const survivors of survivorMaps) {
          for (const [gone, kept] of survivors)
            if (kept === idKey(absorbed)) survivors.set(gone, idKey(survivor));
          survivors.set(idKey(absorbed), idKey(survivor));
        }
        for (const [gone, kept] of openingSurvivorOf) {
          if (kept !== idKey(absorbed)) continue;
          if (opens) openingSurvivorOf.set(gone, idKey(survivor));
          else openingSurvivorOf.delete(gone);
        }
        if (opens) openingSurvivorOf.set(idKey(absorbed), idKey(survivor));
      }
      const { commentId, suggestionId, revisionIds, retiredCommentIds } =
        prepared;
      normalizations.push(...adjustments);
      recordColumnInsertion(resolvedOperation, insertedColumnCounts);
      document = prepared.document;
      ops.push(...prepared.ops);
      if (reuseStaging)
        reuseStaging =
          prepared.verifiedOperation !== undefined &&
          operationStaging.append(prepared.verifiedOperation);
      nextRevisionId = Math.max(nextRevisionId, prepared.nextRevisionId);
      const revisionId = primaryRevisionId({
        operation,
        intents: prepared.intents,
        document: prepared.document,
        revisionIds,
      });
      const joined =
        batch.mode === "direct"
          ? joinedParagraphsOf(operation.type, prepared.intents)
          : undefined;
      appliedById.set(operation.id, {
        id: operation.id,
        ...(commentId === undefined ? {} : { commentId }),
        ...(joined === undefined ? {} : { paragraphIdentity: joined }),
        ...(revisionId === undefined ? {} : { revisionId, revisionIds }),
        ...(suggestionId === undefined ? {} : { suggestionId }),
        ...(retiredCommentIds.length === 0 ? {} : { retiredCommentIds }),
      });
    };
    const inputOrder = new Map(
      batch.operations.map(({ id }, index) => [id, index]),
    );
    const batchRemovedBlockIds = new Set(
      batch.mode === "direct"
        ? resolved.flatMap(({ operation }) =>
            operation.type === "deleteBlock" ? [operation.blockId] : [],
          )
        : [],
    );
    const orderOf = ({ operation }: AdmittedOperation) => {
      const order = inputOrder.get(operation.id);
      if (order === undefined)
        panic("A resolved insertion lost its batch order.");
      return order;
    };
    /**
     * Where an insertion stands among those at its gap: in batch order, but a
     * block put before one the direct batch deletes, placed after it instead
     * (`besideDeletion`), still goes ahead of the blocks put after it.
     */
    const placementOrderOf = (member: AdmittedOperation) =>
      member.besideDeletion === true &&
      placesBlock(member.operation) &&
      placementSide(member.operation) === "after"
        ? orderOf(member) - batch.operations.length
        : orderOf(member);
    const gaps = new Map<number, AdmittedOperation[]>();
    for (const member of executionOrder) {
      if (!placesBlock(member.operation)) continue;
      const place = canonicalInsertionGap({
        operation: member.operation,
        intents: member.intents,
        projection: projected.value,
        landedEarlier: besideDeletions.landedEarlier(member.operation.id),
      });
      if (place === undefined)
        panic("A resolved block insertion lost its placement.");
      const members = gaps.get(place.gap) ?? [];
      members.push(member);
      gaps.set(place.gap, members);
    }
    const indices = new Map(
      executionOrder.map((member, index) => [member, index] as const),
    );
    const intervals: { first: number; last: number }[] = [];
    const groupOf = new Map<AdmittedOperation, readonly AdmittedOperation[]>();
    for (const members of gaps.values()) {
      if (members.length < 2) continue;
      const first = members.at(0),
        last = members.at(-1);
      const from = first === undefined ? undefined : indices.get(first);
      const to = last === undefined ? undefined : indices.get(last);
      if (from === undefined || to === undefined)
        panic("An insertion group lost its execution interval.");
      intervals.push({ first: from, last: to });
      for (const member of members) groupOf.set(member, members);
    }
    const merged: typeof intervals = [];
    for (const interval of intervals.toSorted(
      (left, right) => left.first - right.first,
    )) {
      const previous = merged.at(-1);
      if (previous !== undefined && interval.first <= previous.last)
        previous.last = Math.max(previous.last, interval.last);
      else merged.push({ ...interval });
    }
    const alreadyInReplayOrder = (window: readonly AdmittedOperation[]) => {
      const first = window.at(0);
      const initial = first?.intents.at(0);
      if (
        first === undefined ||
        initial?.type !== "insertParagraphs" ||
        (initial.at.type !== "before" && initial.at.type !== "after")
      )
        return false;
      const group = groupOf.get(first);
      if (group?.length !== window.length) return false;
      const { type: side, blockId } = initial.at;
      const { at: replayAt } = gapAnchorBeforeFollower({
        anchor: { type: "paragraph", at: initial.at },
        projection: projected.value,
        removedBlockIds: batchRemovedBlockIds,
      });
      if (replayAt.type !== side || replayAt.blockId !== blockId) return false;
      let previousOrder: number | undefined;
      for (const member of window) {
        const { operation, intents } = member;
        const intent = intents.at(0);
        if (
          groupOf.get(member) !== group ||
          member.type !== "plain" ||
          !isBlockInsertion(operation) ||
          operation.moveId !== undefined ||
          intents.length !== 1 ||
          intent?.type !== "insertParagraphs" ||
          intent.move !== undefined ||
          intent.newList !== undefined ||
          intent.at.type !== side ||
          intent.at.blockId !== blockId
        )
          return false;
        const order = placementOrderOf(member);
        if (
          previousOrder !== undefined &&
          (side === "before" ? order <= previousOrder : order >= previousOrder)
        )
          return false;
        previousOrder = order;
      }
      // These insertions only add paragraph IDs. Later siblings cannot take
      // any of the IDs an earlier member already selected from the free space.
      return true;
    };
    /** The anchor a gap's insertions replay against, read from its first. */
    const replayAnchorOf = (first: AdmittedOperation) => {
      const place =
        canonicalInsertionGap({
          operation: first.operation,
          intents: first.intents,
          projection: projected.value,
        }) ?? panic("An admitted insertion lost its source placement.");
      return gapAnchorBeforeFollower({
        anchor: place.anchor,
        projection: projected.value,
        removedBlockIds: batchRemovedBlockIds,
      });
    };
    const windows = new Map<AdmittedOperation, readonly AdmittedOperation[]>();
    for (const interval of merged) {
      const first = executionOrder.at(interval.first);
      if (first === undefined)
        panic("An insertion admission window lost its first operation.");
      const window = executionOrder.slice(interval.first, interval.last + 1);
      if (
        batchLease === undefined &&
        batch.mode !== "suggested" &&
        alreadyInReplayOrder(window)
      )
        continue;
      windows.set(first, window);
    }
    const completed = new Set<string>();
    for (const resolvedOperation of executionOrder) {
      if (completed.has(resolvedOperation.operation.id)) continue;
      if (resolvedOperation.plan.address.type !== "annotation")
        annotated ??= document;
      const window = windows.get(resolvedOperation);
      // The ordinary one-operation path performs no admission replay or census.
      if (window === undefined) {
        commitOperation(
          resolvedOperation,
          prepareOperation({
            document,
            resolvedOperation,
            nextRevisionId,
            columnInsertions: insertedColumnCounts,
            ...(reuseStaging
              ? { verifyOperation: operationStaging.apply }
              : {}),
          }),
        );
        continue;
      }
      type AdmittedMember = {
        resolvedOperation: AdmittedOperation;
        firstRevisionId: number;
        ids: LeasedId[];
        ops: readonly DocumentOp[];
        allocations: NonNullable<CompilePublicIntentsOptions["allocations"]>;
        commentAllocation: PrepareOperationOptions["commentAllocation"];
      };
      const admitted: AdmittedMember[] = [];
      const declinedMembers: {
        resolvedOperation: AdmittedOperation;
        prepared: OperationAttempt;
      }[] = [];
      const admissionColumns = new Map(insertedColumnCounts);
      let admissionDocument = document;
      let admissionRevisionId = nextRevisionId;
      let admissionCensus:
        | { type: "paragraphs"; ids: Set<number> | undefined }
        | { type: "package" } =
        batchLease === undefined
          ? { type: "paragraphs", ids: undefined }
          : { type: "package" };
      const rejectParagraphWindow = () => {
        if (batchLease !== undefined) return "lease";
        if (batch.mode === "suggested") return "suggestion";
        if (!window.every(isPlainInsertion)) return "windowShape";
        return undefined;
      };
      const paragraphWindowRejection = rejectParagraphWindow();
      const cursor = startParagraphAdmissionCursor(paragraphWindowRejection);
      for (const member of window) {
        completed.add(member.operation.id);
        const admissionOptions = {
          document: admissionDocument,
          columnInsertions: admissionColumns,
          resolvedOperation: member,
          nextRevisionId: admissionRevisionId,
        };
        const nextLease = cursor.nextLease();
        let prepared = prepareOperation({
          ...admissionOptions,
          ...(nextLease.type === "bounded"
            ? { paragraphAllocationLease: nextLease.lease }
            : {}),
        });
        // An allocation or validator refusal is still determined by the
        // ordinary compiler, with its original exhaustion diagnostics.
        if (nextLease.type === "bounded" && prepared.type === "declined")
          prepared = prepareOperation(admissionOptions);
        // Refused the mark of a paragraph under review beside a block of its
        // gap admitted before it (a table now after the paragraph it names,
        // say), it is admitted where the gap replays, as it then runs.
        const group = groupOf.get(member);
        const lead =
          group === undefined
            ? undefined
            : admitted.find(
                (entry) => groupOf.get(entry.resolvedOperation) === group,
              )?.resolvedOperation;
        if (
          prepared.type === "declined" &&
          prepared.reason === "pendingParagraphMarkDeletion" &&
          lead !== undefined
        ) {
          const replayed = prepareOperation({
            ...admissionOptions,
            resolvedOperation: {
              ...member,
              intents: placeInsertionIntents({
                intents: member.intents,
                anchor: replayAnchorOf(lead),
                placement: "following",
              }),
            },
          });
          if (replayed.type !== "declined") prepared = replayed;
        }
        if (prepared.type === "declined") {
          declinedMembers.push({ resolvedOperation: member, prepared });
          continue;
        }
        recordColumnInsertion(member, admissionColumns);
        const paragraphs =
          admissionCensus.type === "paragraphs"
            ? plainInsertedParagraphIds(prepared.ops)
            : undefined;
        cursor.commit(paragraphs);
        let ids: LeasedId[];
        if (admissionCensus.type === "paragraphs" && paragraphs !== undefined) {
          // Without a lease replay reserves only sibling paragraph IDs. Seed
          // the numeric set from the full census to preserve legacy spellings.
          admissionCensus.ids ??= new Set(
            leasedIdCensus(admissionDocument).paragraph,
          );
          ids = [];
          for (const id of paragraphs) {
            if (admissionCensus.ids.has(id)) continue;
            admissionCensus.ids.add(id);
            ids.push({ kind: ID_LEASE_KINDS.PARAGRAPH, id });
          }
        } else {
          ids = introducedDocumentIds({
            before: admissionDocument,
            after: prepared.document,
          });
          admissionCensus = { type: "package" };
        }
        admitted.push({
          resolvedOperation: member,
          firstRevisionId: admissionRevisionId,
          allocations: prepared.allocations,
          ops: prepared.ops,
          commentAllocation: prepared.commentAllocation,
          ids,
        });
        admissionDocument = prepared.document;
        admissionRevisionId = prepared.nextRevisionId;
      }
      const admissionOf = new Map(
        admitted.map((member) => [member.resolvedOperation, member] as const),
      );
      const ordered: {
        member: AdmittedMember;
        operation: AdmittedOperation;
      }[] = [];
      const scheduled = new Set<AdmittedOperation>();
      for (const member of admitted) {
        if (scheduled.has(member.resolvedOperation)) continue;
        const group = groupOf.get(member.resolvedOperation);
        const partners =
          group === undefined
            ? []
            : group.flatMap((sourceMember) => {
                const admission = admissionOf.get(sourceMember);
                return admission === undefined ? [] : [admission];
              });
        const first = partners.at(0);
        if (first === undefined || partners.length === 1) {
          scheduled.add(member.resolvedOperation);
          ordered.push({ member, operation: member.resolvedOperation });
          continue;
        }
        const anchor = replayAnchorOf(first.resolvedOperation);
        const placed = partners.toSorted(
          (left, right) =>
            (placementOrderOf(left.resolvedOperation) -
              placementOrderOf(right.resolvedOperation)) *
            (anchor.at.type === "before" ? 1 : -1),
        );
        for (const [index, partner] of placed.entries()) {
          scheduled.add(partner.resolvedOperation);
          ordered.push({
            member: partner,
            operation: {
              ...partner.resolvedOperation,
              intents: placeInsertionIntents({
                intents: partner.resolvedOperation.intents,
                anchor,
                placement: index === 0 ? "first" : "following",
              }),
            },
          });
        }
      }
      const paragraphRejection =
        paragraphWindowRejection ??
        (ordered.every(({ operation }) => isPlainInsertion(operation))
          ? undefined
          : "windowShape");
      const paragraphReplay = planReplayParagraphIds(
        paragraphRejection,
        paragraphRejection === undefined
          ? ordered.map(({ member, operation }) => ({
              id: operation.operation.id,
              ops: member.ops,
              ids: member.ids,
            }))
          : [],
      );
      for (const declinedMember of declinedMembers)
        commitOperation(
          declinedMember.resolvedOperation,
          declinedMember.prepared,
        );
      for (const { member, operation: replayOperation } of ordered) {
        const idLease =
          batchLease === undefined
            ? undefined
            : insertionMemberIdLease({
                lease: batchLease,
                siblings: admitted
                  .filter((other) => other !== member)
                  .map(({ ids }) => ids),
              });
        let reservedBlockIds = session.reservedBlockIds;
        let paragraphAllocationLease: IdLease | undefined;
        if (paragraphReplay.type === "admitted") {
          paragraphAllocationLease = paragraphReplay.leases.get(
            replayOperation.operation.id,
          );
          if (paragraphAllocationLease === undefined)
            panic("An admitted insertion lost its paragraph replay lease.");
        } else {
          const reserved = new Set(session.reservedBlockIds);
          for (const other of admitted) {
            if (other === member) continue;
            for (const { kind, id } of other.ids) {
              if (kind === "paragraph") reserved.add(formatParaId(id));
            }
          }
          const droppedCarrier = droppedInsertionCarrier({
            admitted: member.resolvedOperation.intents,
            replayed: replayOperation.intents,
            ops: member.ops,
          });
          if (droppedCarrier !== undefined) reserved.add(droppedCarrier);
          reservedBlockIds = reserved;
        }
        const replayOptions = {
          document,
          columnInsertions: insertedColumnCounts,
          resolvedOperation: replayOperation,
          nextRevisionId: member.firstRevisionId,
          reservedBlockIds,
          ...(paragraphAllocationLease === undefined
            ? {}
            : { paragraphAllocationLease }),
          allocations: member.allocations,
          ...(reuseStaging ? { verifyOperation: operationStaging.apply } : {}),
          ...(member.commentAllocation === undefined
            ? {}
            : { commentAllocation: member.commentAllocation }),
          ...(idLease === undefined ? {} : { idLease }),
        };
        let prepared = prepareOperation(replayOptions);
        // Replayed where it lands beside the others of its gap, a member may
        // need more revision ids than it was admitted with: it takes them
        // past every admitted member's.
        if (prepared.type === "declined")
          prepared = prepareOperation({
            ...replayOptions,
            nextRevisionId: Math.max(nextRevisionId, admissionRevisionId),
          });
        if (prepared.type === "declined")
          panic(`Admitted insertion failed when ordered: ${prepared.message}`);
        commitOperation(replayOperation, prepared);
      }
      nextRevisionId = Math.max(nextRevisionId, admissionRevisionId);
    }
    // A move half that did not apply leaves the other a plain insertion or
    // deletion, reported unpaired: the batch applies again without the move.
    const applyingHalves = new Map<string, number>();
    for (const { operation } of resolved) {
      const moveId = blockMoveId(operation);
      if (moveId !== undefined && appliedById.has(operation.id))
        applyingHalves.set(moveId, (applyingHalves.get(moveId) ?? 0) + 1);
    }
    const orphaned = new Set(
      [...pairedMoves].filter((moveId) => applyingHalves.get(moveId) === 1),
    );
    if (orphaned.size > 0) {
      const orphanedMove = (operation: FolioDocumentOperation) => {
        const moveId = blockMoveId(operation);
        return moveId !== undefined && orphaned.has(moveId);
      };
      const plain = this.apply({
        ...options,
        batch: {
          ...batch,
          operations: batch.operations.map((operation) =>
            orphanedMove(operation) ? withoutMove(operation) : operation,
          ),
        },
      });
      const appliedIds = new Set(plain.applied.map(({ id }) => id));
      return {
        ...plain,
        normalizations: [
          ...(plain.normalizations ?? []),
          ...batch.operations.flatMap((operation) => {
            const moveId = blockMoveId(operation);
            return moveId !== undefined &&
              orphaned.has(moveId) &&
              appliedIds.has(operation.id)
              ? [{ id: operation.id, code: "unpairedMove" as const, moveId }]
              : [];
          }),
        ],
      };
    }
    // The batch's deletions and merges were resolved beside the blocks it
    // places or removes there. Replan beside the blocks that actually land
    // or stay, and preserve the original refusal for each excluded operation.
    const declinedLayoutChanges = new Set(
      batch.atomic
        ? []
        : batch.operations.flatMap((operation) =>
            (placesBlock(operation) && declinedAsRun.has(operation.id)) ||
            (operation.type === "deleteTable" &&
              !appliedById.has(operation.id) &&
              besideDeletions.deleted.size > 0)
              ? [operation.id]
              : [],
          ),
    );
    if (
      declinedLayoutChanges.size > 0 &&
      batch.operations.some(
        ({ type }) => type === "deleteBlock" || type === "mergeBlockWithNext",
      )
    ) {
      const rest = this.apply({
        ...options,
        batch: {
          ...batch,
          operations: batch.operations.filter(
            ({ id }) => !declinedLayoutChanges.has(id),
          ),
        },
      });
      const declinedSkips = skipped.flatMap((failure) => {
        if (!declinedLayoutChanges.has(failure.id)) return [];
        const canonicalRefusal = refusals.get(failure.id);
        return [canonicalRefusal ? { ...failure, canonicalRefusal } : failure];
      });
      const allSkipped = [...rest.skipped, ...declinedSkips].toSorted(
        (left, right) =>
          (inputOrder.get(left.id) ?? 0) - (inputOrder.get(right.id) ?? 0),
      );
      return {
        ...rest,
        skipped: allSkipped,
        issues: getFolioDocumentOperationIssues(batch.operations, allSkipped),
        receipts: getFolioDocumentOperationReceiptsForStory({
          operations: batch.operations,
          applied: rest.applied,
          story: publicStory,
        }),
      };
    }
    if (batch.atomic && skipped.length > 0) {
      for (const { id } of appliedById.values())
        skip(id, "atomicBatchRejected");
      return result([], null);
    }
    if (ops.length === 0) return result([], null);
    const prepared = projected.value
      .selectionAt(state)
      .andThen((selection) =>
        reuseStaging
          ? operationStaging.prepare(state, selection)
          : operationStaging.prepareOps(state, ops, selection),
      );
    if (prepared.isErr()) {
      // A batch that changes nothing is a no-op, not a refusal; any other
      // refusal reports the session's own ground.
      const { reason, ground, message } = prepared.error;
      const skipReason = stagingRefusalReason(reason, ground);
      for (const { id } of appliedById.values()) {
        if (reason !== "noChange") refusals.set(id, ground);
        skip(id, skipReason, message);
      }
      return result([], null);
    }
    addJoinedParagraphReceipts({
      absorptions: prepared.value.touched.absorbed,
      operations: batch.operations,
      applied: appliedById,
    });
    if (batch.mode === "direct") {
      const finalBody = findStoryBody(prepared.value.document, story);
      if (finalBody === undefined)
        panic("A committed public operation lost its story.");
      const finalIds = new Set(
        paragraphsThroughTextBoxes(prepared.value.document, story).map(
          ({ paragraph }) => paragraph.paraId,
        ),
      );
      for (const operation of batch.operations) {
        const entry = appliedById.get(operation.id);
        if (entry === undefined) continue;
        const identity = entry.paragraphIdentity;
        if (identity?.type === "joined") {
          if (
            finalIds.has(identity.retiredBlockId) &&
            finalIds.has(identity.blockId)
          )
            entry.paragraphIdentity = {
              type: "retained",
              blockId: identity.blockId,
            };
          continue;
        }
        if (operation.type === "deleteBlock" && finalIds.has(operation.blockId))
          entry.paragraphIdentity = {
            type: "retained",
            blockId: operation.blockId,
          };
      }
    } else if (batch.mode !== "suggested") {
      const trackedOperations = [...appliedById.values()];
      const originalDocument = readModel(session.document);
      const review = createCanonicalBatchReview({
        story,
        before: { document: originalDocument, projection: state.doc },
        after: {
          document: prepared.value.document,
          projection: prepared.value.projection.doc,
        },
      });
      addRetiredCommentReceipts({
        before: state,
        after: prepared.value.projection.doc,
        applied: trackedOperations,
        operations: batch.operations.filter(
          (operation) =>
            operation.type !== "deleteBlock" ||
            operation.moveId === undefined ||
            !pairedMoves.has(operation.moveId),
        ),
        sourceRanges: new Map(
          resolved.flatMap(({ operation, deletionSpans }) =>
            deletionSpans === undefined
              ? []
              : [[operation.id, deletionSpans] as const],
          ),
        ),
        acceptedCommentIds: review.acceptedCommentIds,
      });
      const message = canonicalTrackedBatchFidelityError({
        review,
        createdCommentIds: new Set(
          trackedOperations.flatMap(({ commentId }) =>
            commentId === undefined ? [] : [commentId],
          ),
        ),
        retiredCommentIds: trackedOperations.flatMap(
          ({ retiredCommentIds }) => retiredCommentIds ?? [],
        ),
      });
      if (message !== null) {
        for (const { id } of trackedOperations)
          skip(id, "invalidResult", message);
        return result([], null, "refused");
      }
    }
    keepStandingRevisionReceipts({
      document: prepared.value.document,
      applied: appliedById,
    });
    const applied = batch.operations.flatMap(({ id }) => {
      const entry = appliedById.get(id);
      if (entry === undefined) return [];
      return [
        batch.dryRun
          ? {
              id,
              ...(entry.paragraphIdentity === undefined
                ? {}
                : { paragraphIdentity: entry.paragraphIdentity }),
              ...(entry.retiredCommentIds === undefined
                ? {}
                : { retiredCommentIds: entry.retiredCommentIds }),
            }
          : entry,
      ];
    });
    if (batch.dryRun) return result(applied, null);
    const startVersion = session.version;
    if (!this.options.publish(prepared.value)) {
      for (const { id } of applied) skip(id, "documentNotEditable");
      return result([], null);
    }
    const handle = {
      type: "documentOperationUndo",
      id: `canonical-${this.handleNamespace}-${String(this.nextHandle++)}`,
    } as const;
    this.handles.push({
      handle,
      version: session.version,
      startVersion,
      story,
    });
    const excess = this.handles.length - this.options.undoLimit;
    if (excess > 0) this.handles.splice(0, excess);
    return result(applied, handle);
  }

  undo(
    handle: FolioDocumentOperationUndoHandle,
  ): FolioDocumentOperationUndoResult {
    const index = this.handles.findIndex(
      ({ handle: existing }) =>
        existing.id === handle.id && existing.type === handle.type,
    );
    if (index < 0)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "unknownHandle",
      };
    if (index !== this.handles.length - 1)
      return { status: "rejected", undoHandle: handle, reason: "notLatest" };
    const entry = this.handles.at(-1);
    const { session } = this.options;
    if (!entry || entry.version !== session.version)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "documentChanged",
      };
    const state = this.options.getState(entry.story);
    if (!state)
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "storyUnavailable",
      };
    const prepared = session.prepareUndo(state, entry.story);
    if (prepared.isErr())
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "inverseRefused",
        message: prepared.error.message,
      };
    if (!this.options.publish(prepared.value))
      return {
        status: "rejected",
        undoHandle: handle,
        reason: "documentNotEditable",
      };
    this.handles.pop();
    // Undo advances the version while exposing the preceding unchanged
    // journal group: the batch before stays undoable when nothing came between
    // it and the batch just undone.
    const previous = this.handles.at(-1);
    if (previous && previous.version === entry.startVersion)
      previous.version = session.version;
    return { status: "undone", undoHandle: handle };
  }

  /**
   * A whole paragraph removed as the editor removes a selection of it with
   * its paragraph mark: from its start to the start of the paragraph after
   * it, or, for the paragraph that ends its block list, from the end of the
   * paragraph before it, whose mark then ends the list. Tracked, the text and
   * the mark are marked deleted.
   */
  private resolveBlockDeletion({
    operation,
    paragraph,
    projection,
    projectionStory,
    doc,
    tracked,
    deletesJoinedWords,
    gainsFollower,
    keepsTablesApart,
    followsInsertedTable,
    precedesInsertedTable,
    bordersRemovedTable,
    previousGainsFollower,
    tableBeforeGainsFollower,
    endsOnceAccepted,
    deletedRun,
    deletedByBatch,
    refusal,
  }: ResolveBlockDeletionOptions) {
    // A direct deletion leaves no text for a comment to cover, as the default
    // engine finds; a tracked one covers the deleted words.
    if (operation.comment !== undefined && !tracked)
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    // A pending deletion already takes it whole, mark and words, as one of
    // the joined paragraph a deletion of the paragraph it joins took.
    if (tracked && goesWholeOnAcceptance(paragraph.node))
      return refusal("noopOperation");
    const end = paragraph.start + paragraph.node.content.size;
    const after = doc.resolve(end + 1).nodeAfter;
    const before = doc.resolve(paragraph.start - 1).nodeBefore;
    const neighbour = (node: PMNode | null) => {
      const id: unknown = node?.attrs["paraId"];
      return node?.type.name === "paragraph" && typeof id === "string"
        ? projection.paragraph(id)
        : undefined;
    };
    // Block-level range boundaries hold nothing: the paragraph past them is
    // the one after it, whose mark the block deletion joins it to.
    const next = neighbour(nodePastBoundaries(doc.resolve(end + 1)));
    // A table before it, standing or put there by the batch, leaves it no
    // paragraph before it to end the list in its place.
    const previous = followsInsertedTable ? undefined : neighbour(before);
    // A tracked deletion's comment covers the deleted paragraph's words.
    const commentFrom = projection.addressAt(paragraph.start);
    const commentTo = projection.addressAt(end);
    if (commentFrom.isErr() || commentTo.isErr())
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    const tracking = (
      resolved: Omit<
        Extract<PreparedOperation, { type: "plain" }>,
        "type" | "operation"
      >,
    ): Result<PreparedOperation, CanonicalPublicOperationResolutionError> =>
      Result.ok(
        operation.comment === undefined
          ? { type: "plain", operation, ...resolved }
          : {
              type: "commented",
              operation,
              comment: {
                text: operation.comment.text,
                from: commentFrom.value,
                to: commentTo.value,
              },
              ...resolved,
            },
      );
    // Before a table, tracked, the paragraph's own mark is the one deleted,
    // which a text range cannot say: the block deletion records it, as it
    // records a move's source, which the batch pairs before compiling.
    // So does a deletion the paragraph after it survives, by identity: an
    // insertion before that paragraph in the same batch leaves it in place.
    // With no paragraph after it, one the batch inserts beside it applies
    // first at one position: after it, that paragraph then follows it;
    // tracked, before it, the deletion finds it by identity there too, and
    // the inserted paragraph's own pending mark joins it. A move's halves do
    // the same, so no mark is left on the container's final paragraph.
    // Before a table, a paragraph whose own mark is a pending insertion and
    // that holds words of the original keeps that mark: tracked, the deletion
    // takes its words only, and rejecting the insertion later runs them into
    // the table as it would have. A paragraph the insertion added whole goes
    // with it, as the block deletion takes it. Where the block deletion can
    // record the deletion beside that insertion, it does, as the batch runs
    // (`addedMarkBeforeTableAsItRuns`).
    // Between two tables (`keepsTablesApart`), the deletion takes the words
    // and leaves the mark in every mode: directly the paragraph stays, empty,
    // and accepting the tracked deletion keeps it, so the tables stay apart.
    // With survivor ids option C the joined paragraph a pending join makes of
    // it carries its id: tracked, the words joined into it go too, as deleting
    // it directly after the join leaves nothing. A paragraph of that join the
    // batch deletes itself goes by its own deletion.
    const joined = deletesJoinedWords
      ? joinedBefore(doc, paragraph.start - 1)
      : [];
    const takesJoinedWords =
      deletesJoinedWords && !joined.some((id) => deletedByBatch.has(id));
    const opening = takesJoinedWords ? joined.at(-1) : undefined;
    const joinedWordsStart =
      opening === undefined ? undefined : projection.paragraph(opening)?.start;
    let holdsOriginalWords = false;
    paragraph.node.descendants((node) => {
      if (
        node.isText &&
        !node.marks.some(
          ({ type }) => type.name === "insertion" || type.name === "deletion",
        )
      )
        holdsOriginalWords = true;
    });
    if (
      operation.moveId === undefined &&
      !gainsFollower &&
      (keepsTablesApart ||
        (tracked &&
          // A table after it, standing or put there by the batch first.
          (after?.type.name === "table" || precedesInsertedTable) &&
          isInsertedPPrMark(paragraph.node.attrs["pPrMark"]) &&
          holdsOriginalWords &&
          // Words a pending join runs into it go by the block deletion,
          // which reads the joined paragraph whole and records itself on an
          // added mark before a table its author's insertion put there.
          joinedWordsStart === undefined))
    ) {
      if (paragraph.start === end) return refusal("noopOperation");
      // Words a pending join runs into it go with its own, as below. Between
      // two tables nothing runs into it: a table the batch puts before it
      // stands between it and the words of the join.
      const start = keepsTablesApart
        ? paragraph.start
        : (joinedWordsStart ?? paragraph.start);
      const from = projection.addressAt(start);
      const to = projection.addressAt(end);
      if (from.isErr() || to.isErr())
        return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
      return tracking({
        intents: [
          { type: "replaceText", from: from.value, to: to.value, text: "" },
        ],
        normalizations: [],
        from: start,
        to: end,
        ...(keepsTablesApart ? { keepsTablesApart: true as const } : {}),
      });
    }
    // Ending a section before a table or a content control, with a paragraph
    // before it, the paragraph keeps its mark and with it its section break,
    // as the direct deletion's join leaves it: tracked, the deletion takes
    // its words and the mark of the paragraph before, which accepting joins
    // into it, as it does for a container's final paragraph.
    const keepsSectionBreakBeforeBlock =
      tracked &&
      after !== null &&
      after.type.name !== "paragraph" &&
      paragraph.node.attrs["_sectionProperties"] != null &&
      previous !== undefined &&
      previous.node.attrs["_sectionProperties"] == null &&
      operation.moveId === undefined &&
      !gainsFollower &&
      !previousGainsFollower &&
      noDeletedRunBefore(deletedRun);
    // At a final join destination, the range fallback starts after the
    // joined words. Use the block deletion to take those words while the
    // destination still supplies the container's last mark.
    // The intent collects deleted marks. Move-source marks belong to
    // their paired move and their words stay under that move's review.
    const deletesFinalJoin =
      after === null &&
      takesJoinedWords &&
      previous?.source.pPrMark?.kind === "del" &&
      !isDeletedPPrMark(paragraph.node.attrs["pPrMark"]) &&
      !previousGainsFollower &&
      noDeletedRunBefore(deletedRun);
    if (
      !keepsSectionBreakBeforeBlock &&
      (next !== undefined ||
        gainsFollower ||
        deletesFinalJoin ||
        (tracked && (after !== null || operation.moveId !== undefined)))
    ) {
      const intents: EditorIntent[] = [
        {
          type: "deleteParagraphs",
          story: projectionStory,
          blockIds: [paragraph.blockId],
          ...(takesJoinedWords ? { joinedWords: "deleted" as const } : {}),
        },
      ];
      // A move's source ending its block list after paragraphs the batch
      // deletes retires the break before them, as a plain deletion there
      // does, and runs in the same order beside them. So does a deletion with
      // only paragraphs after it that pending deletions already take.
      const retiresBefore =
        tracked &&
        ((operation.moveId !== undefined &&
          next === undefined &&
          after === null) ||
          endsOnceAccepted) &&
        !gainsFollower &&
        previous !== undefined;
      const retiresRun = retiresBefore && !noDeletedRunBefore(deletedRun);
      let from = paragraph.start;
      if (retiresRun)
        from =
          deletedRun.type === "bounded" ? deletedRun.claimFrom : previous.start;
      // It retires the break of a paragraph the batch puts right before it.
      else if (retiresBefore && endsOnceAccepted && previousGainsFollower)
        from = previous.start;
      const runCarrier =
        retiresRun &&
        (deletedRun.type === "kept" || deletedRun.type === "runStart")
          ? deletedRun.type
          : undefined;
      return tracking({
        intents,
        normalizations: [],
        from,
        to: end,
        ...(runCarrier === undefined ? {} : { retiresRunBreak: runCarrier }),
        ...(after === null && !gainsFollower
          ? {}
          : { removesParagraphBeforeBlock: true }),
        // A paragraph follows it now; the batch may delete that one first.
        ...(!tracked && next !== undefined && !gainsFollower
          ? { mayEndContainer: true as const }
          : {}),
      });
    }
    let range: { from: number; to: number };
    // Removed directly, a paragraph no paragraph follows loses its words and
    // then joins the paragraph before it, whose properties it keeps; the
    // joined paragraph is the one whose mark remains, this one, as accepting
    // the tracked deletion leaves it. With none before it, it only loses its
    // words.
    if (!tracked) {
      // Before a table or another block, first in its container or after a
      // paragraph, it goes whole, as accepting its tracked deletion removes
      // it: the paragraph before it keeps its own mark, and with it any
      // section break. A section break of its own is resolved by the join
      // below. This also applies beside a removed table when the paragraph
      // no longer keeps surviving tables apart.
      if (
        after !== null &&
        !followsInsertedTable &&
        (before === null ||
          ((before.type.name === "paragraph" || bordersRemovedTable) &&
            paragraph.node.attrs["_sectionProperties"] == null))
      )
        return tracking({
          intents: [
            {
              type: "deleteParagraphs",
              story: projectionStory,
              blockIds: [paragraph.blockId],
            },
          ],
          normalizations: [],
          from: paragraph.start,
          to: end,
          removesParagraphBeforeBlock: true,
        });
      const own = (
        from: number,
        to: number,
      ): Result<EditorIntent[], CanonicalPublicOperationResolutionError> => {
        const start = projection.addressAt(from);
        const stop = projection.addressAt(to);
        return start.isErr() || stop.isErr()
          ? refusal("unsupportedBlock", undefined, INVALID_REQUEST)
          : Result.ok(
              from === to
                ? []
                : [
                    {
                      type: "replaceText",
                      from: start.value,
                      to: stop.value,
                      text: "",
                    },
                  ],
            );
      };
      // A paragraph the batch puts after the table before it stands before
      // it when the deletion runs: it joins that one then, as deleting it
      // after that insertion does.
      if (previous === undefined && after === null && tableBeforeGainsFollower)
        return tracking({
          intents: [
            {
              type: "deleteParagraphs",
              story: projectionStory,
              blockIds: [paragraph.blockId],
            },
          ],
          normalizations: [],
          from: paragraph.start,
          to: end,
          mayEndContainer: true,
          followsTableInsertion: true,
        });
      const cleared = own(paragraph.start, end);
      if (cleared.isErr()) return cleared;
      const intents: EditorIntent[] = [...cleared.value];
      // The paragraph before it takes over: before a table or a content
      // control the deleted paragraph's mark goes with it, and the one before
      // it keeps its identity (`directJoinSurvivor`), unless that mark ends a
      // section: it stays with its break, as accepting the tracked deletion
      // leaves it. Collapsing a run the batch deletes back to its start, it
      // joins the run's paragraph before it, which keeps its identity and
      // properties until its own deletion.
      const collapsesRunStart =
        previous !== undefined && deletedRun.type === "runStart";
      if (previous !== undefined)
        intents.push({
          type: "joinParagraphs",
          story: projectionStory,
          blockId: previous.blockId,
          nextBlockId: paragraph.blockId,
          survivor: collapsesRunStart
            ? SPLIT_HALVES.FIRST
            : directJoinSurvivor({
                laterMarkRemains:
                  after === null ||
                  paragraph.node.attrs["_sectionProperties"] != null,
              }),
        });
      if (intents.length === 0) return refusal("noopOperation");
      let joinOrder = {};
      if (collapsesRunStart) joinOrder = { collapsesRunStart: true as const };
      else if (previous !== undefined && previousGainsFollower)
        joinOrder = { joinsParagraphBefore: true as const };
      return Result.ok({
        type: "plain" as const,
        operation,
        intents,
        normalizations: [],
        from: previous === undefined ? paragraph.start : previous.start,
        to: end,
        ...joinOrder,
      });
    } else if (previous !== undefined && deletedRun.type === "bounded") {
      // The batch deletes the paragraphs right before it and inserts a block
      // that then stands before them. Claimed from before that insertion and
      // the run's deletions, the block deletion runs after them and retires
      // the break before the run as it stands then: the inserted paragraph's,
      // whose properties this paragraph takes, or none after a table, as one
      // at a time.
      return tracking({
        intents: [
          {
            type: "deleteParagraphs",
            story: projectionStory,
            blockIds: [paragraph.blockId],
          },
        ],
        normalizations: [],
        from: deletedRun.claimFrom,
        to: end,
      });
    } else if (
      previous !== undefined &&
      (deletedRun.type === "kept" || deletedRun.type === "runStart")
    ) {
      // The batch deletes the paragraphs right before it. With a kept one
      // before them, it runs after their deletions and retires the break
      // before them; back to the container's start or a table, it runs
      // before the first one's deletion and retires that one's break, as it
      // does one at a time once the others are pending.
      return tracking({
        intents: [
          {
            type: "deleteParagraphs",
            story: projectionStory,
            blockIds: [paragraph.blockId],
          },
        ],
        normalizations: [],
        from: previous.start,
        to: end,
        retiresRunBreak: deletedRun.type,
      });
    } else if (previous !== undefined) {
      // Paragraphs a pending deletion already takes whole, words and mark,
      // stand between it and the break the deletion retires: that is the
      // break before them, as accepting them would leave it.
      let retiring = previous;
      for (;;) {
        if (
          !isDeletedPPrMark(retiring.node.attrs["pPrMark"]) ||
          holdsStandingContent(retiring.node)
        )
          break;
        const earlier = neighbour(doc.resolve(retiring.start - 1).nodeBefore);
        if (earlier === undefined) break;
        retiring = earlier;
      }
      // The block deletion retires that break and gives this paragraph the
      // words' properties, which a text range over the joined marks cannot.
      if (retiring !== previous)
        return tracking({
          intents: [
            {
              type: "deleteParagraphs",
              story: projectionStory,
              blockIds: [paragraph.blockId],
            },
          ],
          normalizations: [],
          from: retiring.start,
          to: end,
        });
      range = {
        from: previous.start + previous.node.content.size,
        to: end,
      };
      // A paragraph the batch puts right after the one before it is the one
      // whose break this deletion retires: claimed from the start of the
      // paragraph before, the deletion runs after that insertion, as the
      // direct deletion's join does. A run the batch deletes before it is
      // `retiresRunBreak`'s.
      if (previousGainsFollower && noDeletedRunBefore(deletedRun)) {
        const from = projection.addressAt(range.from);
        const to = projection.addressAt(end);
        if (from.isErr() || to.isErr())
          return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
        return tracking({
          intents: [
            { type: "replaceText", from: from.value, to: to.value, text: "" },
          ],
          normalizations: [],
          from: previous.start,
          to: end,
          retiresBreakBefore: true,
        });
      }
    }
    // After a table, or as the only paragraph of its container, the
    // paragraph ends its block list for good: it keeps its mark and loses its
    // words, as a tracked deletion of the whole paragraph leaves it and as
    // deleting it directly does.
    else {
      // Unless the batch puts a paragraph after that table: standing before
      // it when the deletion runs, that one's break is the one it retires.
      if (after === null && tableBeforeGainsFollower) {
        const from = projection.addressAt(paragraph.start);
        const to = projection.addressAt(end);
        if (from.isErr() || to.isErr())
          return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
        return tracking({
          intents: [
            { type: "replaceText", from: from.value, to: to.value, text: "" },
          ],
          normalizations: [],
          from: paragraph.start,
          to: end,
          retiresBreakBefore: true,
          followsTableInsertion: true,
        });
      }
      if (paragraph.node.content.size === 0) return refusal("noopOperation");
      range = { from: paragraph.start, to: end };
    }
    const from = projection.addressAt(range.from);
    const to = projection.addressAt(range.to);
    if (from.isErr() || to.isErr())
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    const intents: EditorIntent[] = [
      { type: "replaceText", from: from.value, to: to.value, text: "" },
    ];
    return tracking({
      intents,
      normalizations: [],
      from: Math.min(range.from, paragraph.start),
      to: Math.max(range.to, end),
    });
  }

  /**
   * An operation resolved against its story. One naming the filler of a
   * block the source leaves empty is first restated beside that block.
   */
  private resolve(options: ResolveCanonicalPublicOperationOptions) {
    const restated = this.restatedBesideEmptyControl(options);
    if (restated.isErr()) return Result.err(restated.error);
    const { operation, listStart } = restated.value;
    return this.resolveOperation({
      ...options,
      operation,
      ...(listStart === undefined ? {} : { listStart }),
    });
  }

  /**
   * `operation`, or, where it names the filler of a block the source leaves
   * empty (`emptyControlOfFiller`), the same insertion beside that block,
   * with the start of its list where that is where it goes; any other
   * operation on the filler is refused, as is one whose precondition the
   * filler, which holds no words, does not meet.
   */
  private restatedBesideEmptyControl({
    operation,
    options,
    state,
    story,
  }: ResolveCanonicalPublicOperationOptions): Result<
    Pick<ResolveCanonicalPublicOperationOptions, "operation" | "listStart">,
    CanonicalPublicOperationResolutionError
  > {
    const blockId =
      operation.type === "replaceRange" ||
      operation.type === "formatRange" ||
      operation.type === "commentOnRange"
        ? operation.range.blockId
        : operation.blockId;
    const anchor = Object.hasOwn(options.snapshot.anchors, blockId)
      ? options.snapshot.anchors[blockId]
      : undefined;
    const projection = this.options.session.projectStory(story);
    if (
      anchor === undefined ||
      projection.isErr() ||
      projection.value.paragraph(blockId) !== undefined
    )
      return Result.ok({ operation });
    const control = emptyControlOfFiller(state.doc, blockId);
    if (control === undefined) return Result.ok({ operation });
    const refusal: ResolutionRefusal = (reason, message, ground) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message: message ?? `The operation was refused: ${reason}.`,
          detail: message,
          ...(ground === undefined ? {} : { ground }),
        }),
      );
    const fillerHash = hashFolioAIBlockContent("");
    if (
      operation.precondition &&
      fillerHash !== operation.precondition.blockTextHash
    )
      return refusal("preconditionFailed");
    if (fillerHash !== anchor.textHash) return refusal("changedBlock");
    const beside = besideEmptyControl({
      operation,
      control,
      projection: projection.value,
    });
    switch (beside.type) {
      case "refused":
        return refusal("unsupportedBlock", beside.message, INVALID_REQUEST);
      case "beside":
        return Result.ok({ operation: beside.operation });
      case "listStart":
        return Result.ok({
          operation: beside.operation,
          listStart: beside.listStart,
        });
      default: {
        const unreachable: never = beside;
        return unreachable;
      }
    }
  }

  private resolveOperation({
    operation,
    listStart,
    options,
    state,
    story,
    diff,
    tracked,
    suggested,
    mintList,
    besideDeletions,
    batchComments,
  }: ResolveCanonicalPublicOperationOptions) {
    const refusal: ResolutionRefusal = (reason, message, ground) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message: message ?? `The operation was refused: ${reason}.`,
          detail: message,
          ...(ground === undefined ? {} : { ground }),
        }),
      );
    const skipAlike = (reason: FolioAIEditSkipReason, message: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message,
          detail: message,
        }),
      );
    const skipTracked = (message: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason: "unsupportedBlock",
          message,
          detail: message,
          refusedIn: "trackedMode",
        }),
      );
    const notInTable = () =>
      skipAlike("unsupportedBlock", "The block is not in a table.");
    switch (operation.type) {
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
      case "splitBlock":
      case "formatRange":
      case "mergeBlockWithNext":
      case "commentOnRange":
      case "commentOnBlock":
      case "deleteBlock":
      case "setBlockParagraphProperties":
      case "insertAfterBlock":
      case "insertBeforeBlock":
      case "insertTable":
      case "insertSignatureTable":
      case "deleteTable":
      case "insertTableRow":
      case "deleteTableRow":
      case "insertTableColumn":
      case "deleteTableColumn":
      case "mergeTableCells":
      case "splitTableCell":
        break;
      default: {
        const unreachable: never = operation;
        return unreachable;
      }
    }
    const blockId =
      operation.type === "replaceRange" ||
      operation.type === "formatRange" ||
      operation.type === "commentOnRange"
        ? operation.range.blockId
        : operation.blockId;
    const anchor = Object.hasOwn(options.snapshot.anchors, blockId)
      ? options.snapshot.anchors[blockId]
      : undefined;
    const projection = this.options.session.projectStory(story);
    if (!anchor || projection.isErr()) return refusal("missingBlock");
    const paragraph = projection.value.paragraph(blockId);
    if (!paragraph)
      return holdsParagraph(state.doc, blockId)
        ? refusal(
            "unsupportedBlock",
            "The canonical session holds no address in this paragraph.",
            INVALID_REQUEST,
          )
        : refusal("missingBlock");
    if (
      (operation.type === "insertTable" ||
        operation.type === "insertSignatureTable") &&
      !supportsTableInsertionMode(
        operation.type,
        options.batch.mode ?? "tracked-changes",
      )
    )
      return refusal("unsupportedMode");
    const noteReferences = collectNoteReferenceLabels(state.doc);
    const clean = buildCleanBlockText(paragraph.node, paragraph.start - 1, {
      fieldResults: "text",
      noteReferences,
    });
    // A replacement that names a style the paragraph does not have restyles
    // it, as the legacy engine does, even when its words stay the same.
    const restyles =
      operation.type === "replaceBlock" &&
      operation.styleId !== undefined &&
      operation.styleId !== (paragraph.source.formatting?.styleId ?? null);
    // A paragraph whose break is pending removal takes the edits the default
    // engine takes there, and refuses what it refuses, for the same reason
    // (`pendingMarkEditRefusal`): a restyling replacement edits the
    // paragraph's properties as well, after its words. A block beside it is
    // placed by the join it is part of (`besidePendingJoin`).
    const pendingRefusal = [
      pendingMarkEditOf(operation.type),
      ...(restyles ? (["paragraphProperties"] as const) : []),
    ]
      .map((edit) =>
        pendingMarkEditRefusal({
          edit,
          doc: state.doc,
          paragraphPos: paragraph.start - 1,
          paragraph: paragraph.node,
          cleanText: clean.text,
        }),
      )
      .find((refused) => refused !== undefined);
    if (pendingRefusal !== undefined)
      return Result.err(
        new CanonicalPublicOperationResolutionError({
          reason: pendingRefusal.reason,
          message:
            pendingRefusal.message ??
            `The operation was refused: ${pendingRefusal.reason}.`,
          detail: pendingRefusal.message,
          ...(pendingRefusal.canonicalRefusal === undefined
            ? {}
            : { ground: pendingRefusal.canonicalRefusal }),
        }),
      );
    const hash = hashFolioAIBlockContent(clean.text);
    if (operation.precondition && hash !== operation.precondition.blockTextHash)
      return refusal("preconditionFailed");
    if (hash !== anchor.textHash) return refusal("changedBlock");
    // A paragraph no paragraph follows, deleted directly by the same batch:
    // before it and after it are one place once it is gone, and after it,
    // the deletion keeps a paragraph to remove it by identity. A tracked
    // deletion leaves the paragraph in place, so its sides stay apart, and it
    // removes the paragraph by identity once the insertion beside it applies.
    const blockAfter = state.doc.resolve(
      paragraph.start + paragraph.node.content.size + 1,
    ).nodeAfter;
    // Applied directly, a table the batch puts right after it leaves no
    // paragraph following it (`tableLandsAfter`).
    const followed =
      blockAfter?.type.name === "paragraph" &&
      (tracked ||
        !tableLandsAfter({
          paragraph,
          projection: projection.value,
          noteReferences,
          besideDeletions,
        }));
    const blockBefore = state.doc.resolve(paragraph.start - 1).nodeBefore;
    const previousId: unknown =
      blockBefore?.type.name === "paragraph"
        ? blockBefore.attrs["paraId"]
        : undefined;
    // Find the tables that will actually border the deleted paragraph.
    // A run of removed tables can expose another table behind it.
    const batchRemovesTable = (block: PMNode): boolean =>
      besideDeletions.tables.deleted.size > 0 &&
      tableParagraphIds(block).some((id) =>
        besideDeletions.tables.deleted.has(id),
      );
    const tableWillBeRemoved = (block: PMNode): boolean =>
      block.type.name === "table" &&
      ((tracked && tableGoes(block, "accept")) || batchRemovesTable(block));
    const resolveDeletedPlacement = () => {
      let beforePos = paragraph.start - 1;
      let afterPos = beforePos + paragraph.node.nodeSize;
      let before = blockBefore;
      let after = blockAfter;
      while (before !== null && tableWillBeRemoved(before)) {
        beforePos -= before.nodeSize;
        before = state.doc.resolve(beforePos).nodeBefore;
      }
      while (after !== null && tableWillBeRemoved(after)) {
        const beyond = state.doc.resolve(afterPos + after.nodeSize).nodeAfter;
        // Past a table only a pending deletion takes, a paragraph the batch
        // deletes too leaves this one closing what remains: it still stands
        // between the tables as the batch runs.
        if (
          !batchRemovesTable(after) &&
          beyond?.type.name === "paragraph" &&
          typeof beyond.attrs["paraId"] === "string" &&
          besideDeletions.deleted.has(beyond.attrs["paraId"])
        )
          break;
        afterPos += after.nodeSize;
        after = beyond;
      }
      const idOf = (block: PMNode | null): unknown =>
        block?.type.name === "paragraph" ? block.attrs["paraId"] : undefined;
      // A table the batch puts after the paragraph before, or before the
      // one after, borders it as one standing there does.
      const tableBeside = (block: PMNode | null, side: "before" | "after") => {
        const id = idOf(block);
        return (
          block?.type.name === "table" ||
          (typeof id === "string" &&
            tableLandsBeside({
              blockId: id,
              side,
              projection: projection.value,
              noteReferences,
              besideDeletions,
            }))
        );
      };
      const follower = followed ? "paragraph" : "block";
      return deletedParagraphPlacement({
        beside: besideDeletions.placedBeside(blockId),
        tableBefore: tableBeside(before, "after"),
        tableAfter: tableBeside(after, "before"),
        after: blockAfter === null ? "none" : follower,
      });
    };
    const besideDeleted = besideDeletions.deleted.has(blockId)
      ? resolveDeletedPlacement()
      : undefined;
    const keepsTablesApart = besideDeleted?.type === "keepsTablesApart";
    if (isTableStructureOperation(operation))
      return resolveTableStructure({
        operation,
        paragraph,
        projection: projection.value,
        story,
        refusal,
        tracked,
        paragraphOf: (id) =>
          Object.hasOwn(options.snapshot.anchors, id)
            ? projection.value.paragraph(id)
            : undefined,
        notInTable,
        skipAlike,
        skipTracked,
        deleted: besideDeletions.deleted,
        template: options.tableTemplates?.get(operation.id),
      });
    if (
      operation.type === "insertTable" ||
      operation.type === "insertSignatureTable"
    ) {
      const landedEarlier = besideDeletions.landedEarlier(operation.id);
      // Put after the paragraph before one the direct batch deletes, it may
      // go after that one too (`movedPastDeletion`).
      const movedPast =
        tracked ||
        listStart !== undefined ||
        (operation.position ?? "after") !== "after" ||
        findOutermostTableBoundary(
          projection.value.doc,
          paragraph.start - 1,
        ) !== null
          ? undefined
          : movedPastDeletion({
              operation,
              place: besidePendingJoin({
                side: "after",
                paragraph,
                projection: projection.value,
                noteReferences,
                landedEarlier,
              }).place,
              projection: projection.value,
              noteReferences,
              besideDeletions,
              inputOrder: options.batch.operations,
            });
      // Put before a paragraph the direct batch deletes behind a paragraph
      // put there, it goes after it with that paragraph.
      const besideDeletion =
        movedPast !== undefined ||
        (!tracked &&
          listStart === undefined &&
          besideDeleted?.type === "placedAfter" &&
          besideDeleted.operationIds.has(operation.id));
      let placedOperation: TableInsertion = operation;
      if (movedPast !== undefined)
        placedOperation = { ...operation, blockId: movedPast.blockId };
      else if (besideDeletion)
        placedOperation = { ...operation, position: "after" };
      const inserted = resolveTableInsertion({
        operation: placedOperation,
        ...(listStart === undefined ? {} : { listStart }),
        paragraph: movedPast ?? paragraph,
        projection: projection.value,
        story,
        noteReferences,
        refusal,
        skipAlike,
        template: options.tableTemplates?.get(operation.id),
        landedEarlier,
      });
      return besideDeletion && inserted.isOk()
        ? Result.ok({ ...inserted.value, besideDeletion: true as const })
        : inserted;
    }
    // Nothing stands before a list's start: no pending join reads the place,
    // and no paragraph the batch deletes is beside it.
    if (isBlockInsertion(operation) && listStart !== undefined)
      return resolveBlockInsertion({
        operation,
        listStart,
        paragraph,
        named: paragraph,
        formattingSource: paragraph,
        projection: projection.value,
        noteReferences,
        story,
        mintList,
        refusal,
        skipAlike,
      });
    if (isBlockInsertion(operation)) {
      const landedEarlier = besideDeletions.landedEarlier(operation.id);
      const after =
        operation.type === "insertAfterBlock"
          ? placeAfter({
              paragraph,
              projection: projection.value,
              noteReferences,
              landedEarlier,
            })
          : undefined;
      const movedPast =
        after === undefined || tracked
          ? undefined
          : movedPastDeletion({
              operation,
              place: after.place,
              projection: projection.value,
              noteReferences,
              besideDeletions,
              inputOrder: options.batch.operations,
            });
      const besideDeletion =
        movedPast !== undefined ||
        (!tracked &&
          operation.type === "insertBeforeBlock" &&
          !followed &&
          besideDeletions.deleted.has(blockId) &&
          !keepsTablesApart);
      // A container's final mark remains on a paragraph inserted before it.
      // Before a table or control, its own mark goes: insert after it first
      // so the block deletion has a following paragraph to receive its text
      // (`deletedParagraphPlacement`).
      let insertion: BlockInsertion = operation;
      if (movedPast !== undefined)
        insertion = { ...operation, blockId: movedPast.blockId };
      else if (besideDeletion && blockAfter !== null)
        insertion = { ...operation, type: "insertAfterBlock" };
      // A paragraph put after one whose break a pending join removes goes
      // where `insertionAfterPendingBreak` puts it, and one put before a
      // paragraph goes before the join it is part of, in every mode: a
      // suggestion lands where the same tracked insertion does, so
      // accepting it leaves what the direct edit does.
      const placing = {
        paragraph,
        projection: projection.value,
        noteReferences,
        landedEarlier,
      };
      let placed: BesidePendingJoin;
      if (movedPast !== undefined) placed = besideAsItStands(movedPast);
      else if (insertion.type === "insertAfterBlock")
        placed = after ?? placeAfter(placing);
      else placed = besidePendingJoin({ ...placing, side: "before" });
      // A block the join would run into the words before it, or into a
      // table after it, waits for the join to be resolved.
      const gap =
        insertion.type === "insertAfterBlock"
          ? placed.place.start - 1 + placed.place.node.nodeSize
          : placed.place.start - 1;
      if (
        joinStopsAtTableAfter({
          insertedBlock: "paragraph",
          doc: projection.value.doc,
          paragraphPos: gap,
          noteReferences,
        })
      )
        return refusal(
          PENDING_JOIN_INSERTION_REFUSAL.reason,
          PENDING_JOIN_INSERTION_REFUSAL.message,
        );
      const inserted = resolveBlockInsertion({
        operation: insertion,
        paragraph: placed.place,
        named: paragraph,
        formattingSource: placed.formattingSource,
        projection: projection.value,
        noteReferences,
        story,
        mintList,
        refusal,
        skipAlike,
        landedEarlier,
      });
      return besideDeletion && inserted.isOk()
        ? Result.ok({ ...inserted.value, besideDeletion: true as const })
        : inserted;
    }
    if (operation.type === "setBlockParagraphProperties")
      return resolveParagraphProperties({
        operation,
        paragraph,
        projection: projection.value,
        mintList,
        refusal,
      });
    if (operation.type === "deleteBlock") {
      const gainsFollower =
        !followed &&
        besideDeletions.insertedBeside.has(blockId) &&
        !keepsTablesApart;
      const before = blockBefore;
      const after = blockAfter;
      // Tracked, only paragraphs that pending deletions already take follow
      // it: it ends its container as accepting leaves it, and retires the
      // break before it as a final paragraph's deletion does.
      const endsOnceAccepted =
        tracked &&
        operation.moveId === undefined &&
        followersGoOnAcceptance(
          state.doc,
          paragraph.start + paragraph.node.content.size + 1,
        );
      const followsInsertedTable =
        keepsTablesApart ||
        (!followed &&
          (besideDeletions.tables.before.has(blockId) ||
            (typeof previousId === "string" &&
              besideDeletions.tables.after.has(previousId))));
      // A paragraph the batch puts right after the table before it runs
      // first and stands before this one when the deletion runs.
      const tableBeforeGainsFollower =
        !followed &&
        before !== null &&
        before.type.name !== "paragraph" &&
        paragraphIdsIn(before).some((id) =>
          besideDeletions.paragraphsAfter.has(id),
        );
      const deleted = this.resolveBlockDeletion({
        operation,
        paragraph,
        projection: projection.value,
        projectionStory: story,
        doc: state.doc,
        tracked,
        deletesJoinedWords: deletionTakesJoinedWords(
          options.batch.mode ?? "tracked-changes",
        ),
        gainsFollower,
        keepsTablesApart,
        followsInsertedTable,
        // The table stands right after it only when no paragraph the batch
        // inserts after it lands there too.
        precedesInsertedTable:
          besideDeletions.tables.after.has(blockId) &&
          !besideDeletions.paragraphsAfter.has(blockId),
        bordersRemovedTable:
          (before !== null && tableWillBeRemoved(before)) ||
          (after !== null && tableWillBeRemoved(after)),
        // The paragraph before it is another when the deletion runs: the
        // batch puts one after it, or deletes it first.
        previousGainsFollower:
          (!followed || endsOnceAccepted) &&
          typeof previousId === "string" &&
          (besideDeletions.paragraphsAfter.has(previousId) ||
            besideDeletions.deleted.has(previousId)),
        tableBeforeGainsFollower,
        endsOnceAccepted,
        deletedByBatch: besideDeletions.deleted,
        deletedRun:
          followed && !endsOnceAccepted
            ? { type: "none" }
            : deletedRunBefore({
                doc: state.doc,
                paragraphStart: paragraph.start,
                besideDeletions,
                tracked,
              }),
        refusal,
      });
      return gainsFollower && deleted.isOk()
        ? Result.ok({ ...deleted.value, besideDeletion: true as const })
        : deleted;
    }
    let start = 0;
    let end = clean.text.length;
    let text = "";
    switch (operation.type) {
      case "replaceInBlock": {
        if (operation.find === "") return refusal("emptyOperation");
        start = clean.text.indexOf(operation.find);
        if (start < 0) return refusal("missingFind");
        if (clean.text.indexOf(operation.find, start + 1) >= 0)
          return refusal("ambiguousFind");
        end = start + operation.find.length;
        text = operation.replace;
        break;
      }
      case "commentOnBlock":
        if (operation.quote !== undefined) {
          start = clean.text.indexOf(operation.quote);
          if (start < 0) return refusal("missingFind");
          if (clean.text.indexOf(operation.quote, start + 1) >= 0)
            return refusal("ambiguousFind");
          end = start + operation.quote.length;
        }
        break;
      case "replaceRange":
      case "formatRange":
      case "commentOnRange":
        start = operation.range.startOffset;
        end = operation.range.endOffset;
        if (
          hashFolioAIBlockText(clean.text.slice(start, end)) !==
          operation.range.selectedTextHash
        )
          return refusal("staleRange");
        text = operation.type === "replaceRange" ? operation.replace : "";
        break;
      case "replaceBlock":
        text = operation.text;
        break;
      case "splitBlock":
        start = operation.offset;
        end = start + (operation.separator?.length ?? 0);
        if (
          start <= 0 ||
          start >= clean.text.length ||
          clean.text.slice(start, end) !== (operation.separator ?? "")
        )
          return refusal("staleRange");
        break;
      case "mergeBlockWithNext":
        // A tab or line feed in the separator is written as its element; a
        // carriage return names no element of its own.
        if (/\r/u.test(operation.separator ?? ""))
          return refusal(
            "unsupportedBlock",
            "Canonical public join separators hold no carriage return.",
            INVALID_REQUEST,
          );
        start = clean.text.length;
        end = start;
        break;
      default: {
        const unreachable: never = operation;
        return unreachable;
      }
    }
    // An annotation only adds marks, so it may meet a letter's combining mark
    // or an emoji sequence's joiner; it never halves a surrogate pair. A text
    // change or a break keeps every grapheme cluster whole.
    const characterSplit = describeCharacterSplit(
      clean.text,
      [start, end],
      ANNOTATING_OPERATION_TYPES.has(operation.type) ? "codePoint" : "grapheme",
    );
    if (characterSplit !== null)
      return refusal("splitsCharacter", characterSplit);
    // A note reference is not text: a replacement rewrites the prose around
    // the references its range holds, each kept in place, as the shared
    // segmentation pairs them. A comment or formatting covers whole
    // references as it covers the words beside them, and leaves them in
    // place. A range that cuts a reference and a replacement that drops,
    // moves or invents a reference are refused.
    const replacesText =
      operation.type === "replaceInBlock" ||
      operation.type === "replaceRange" ||
      operation.type === "replaceBlock";
    const references = noteReferenceSpansWithin(clean, start, end);
    if (cutsIntoNoteReference(clean, start, end))
      return refusal("protectedReference");
    const range = resolveCleanTextRange({
      cleanBlock: clean,
      startOffset: start,
      endOffset: end,
    });
    // The text matched above and no reference is cut, so only structure
    // leaves the range unresolved: an edge inside a field result, which
    // names text no run holds, or a page-break run inside it. Both engines
    // skip it alike, with the reason alone, before reading what the
    // replacement writes.
    if (!range)
      return Result.err(
        new CanonicalPublicOperationResolutionError({
          reason: "unsupportedBlock",
          message:
            "The range begins or ends inside a field result or holds a page-break run.",
          detail: undefined,
        }),
      );
    // Inline emphasis or dropped formatting rebuilds the replaced runs whole,
    // the references among them included.
    const dropsFormatting =
      operation.type === "replaceBlock" &&
      operation.preserveFormatting === false;
    const emphasis =
      replacesText && hasInlineEmphasis(text)
        ? parseInlineEmphasisRuns(text)
        : undefined;
    if (
      replacesText &&
      references.length > 0 &&
      (dropsFormatting || emphasis !== undefined)
    )
      return refusal("protectedReference");
    // The replacement writes the words its emphasis markers mark; directly,
    // as the default engine writes them, they are bold or italic, and a
    // replacement that drops the paragraph's formatting writes them with
    // none of their own in a paragraph that states none. A redline is text
    // only, as the default engine records it: it keeps the formatting and
    // writes the words plain.
    if (emphasis !== undefined) text = stripInlineEmphasisMarkers(text);
    const rebuilds = !tracked && dropsFormatting;
    const emphasized = tracked ? [] : emphasizedSpans(emphasis);
    const rewritesWhole = rebuilds || emphasized.length > 0;
    // Only a replacement rewrites the prose between references.
    const kept = segmentsAroundNoteReferences(
      clean.text.slice(start, end),
      replacesText ? text : clean.text.slice(start, end),
      references,
    );
    // A comparison also removes references and writes new ones as their
    // markers (turned back into references once the batch lands): its
    // replacement is one piece, and a change touching a reference takes the
    // whole reference.
    const editsReferences =
      kept === null && replacesText && noteReferenceEditsAllowed();
    const segments = editsReferences
      ? {
          source: [{ offset: 0, text: clean.text.slice(start, end) }],
          replacement: [text],
        }
      : kept;
    if (segments === null) return refusal("protectedReference");
    const prose = replacesText ? segments.replacement.join("") : "";
    // A marker-shaped string the replacement writes of its own would read
    // as a reference the document does not hold; only a comparison turns
    // markers into references.
    if (!editsReferences && /\[\^(?:e)?\d+\]/u.test(prose))
      return refusal(
        "protectedReference",
        "A replacement writes no note marker of its own.",
      );
    // A tab or line feed is written as its element; a carriage return names
    // no element of its own.
    if (/\r/u.test(prose))
      return refusal(
        "unsupportedBlock",
        "Canonical public replacements require plain text.",
        INVALID_REQUEST,
      );
    const from = projection.value.addressAt(range.from);
    const to = projection.value.addressAt(range.to);
    if (from.isErr() || to.isErr())
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    const resolvedComment = resolveComment({
      operation,
      from: from.value,
      to: to.value,
    });
    if (
      operation.type === "commentOnBlock" ||
      operation.type === "commentOnRange"
    ) {
      // A proposed comment is shown, listed and resolved by the words it
      // covers; one over no shown words (a paragraph whose text is empty or
      // all deleted) would stage a suggestion no reviewer can see, as
      // accepting the deletion of a proposed comment's words retires it.
      if (options.batch.mode === "suggested" && start === end)
        return refusal(
          "emptyOperation",
          "A proposed comment needs words of the paragraph to cover.",
        );
      return Result.ok({
        ...resolvedComment,
        intents: [],
        normalizations: [],
        from: range.from,
        to: range.to,
      });
    }
    const restyle =
      restyles &&
      operation.type === "replaceBlock" &&
      operation.styleId !== undefined
        ? statedPropertiesIntents({
            story,
            blockId: paragraph.blockId,
            properties: { styleId: operation.styleId },
            mintList,
          })
        : [];
    if (
      resolvedComment.type === "plain" &&
      operation.type !== "splitBlock" &&
      operation.type !== "formatRange" &&
      operation.type !== "mergeBlockWithNext" &&
      restyle.length === 0 &&
      !rewritesWhole &&
      clean.text.slice(start, end) === text
    )
      return refusal("noopOperation");
    let intent: EditorIntent;
    switch (operation.type) {
      case "splitBlock":
        // The first half keeps the paragraph's identity, as the default
        // engine's split does.
        intent = {
          type: "splitParagraph",
          at: from.value,
          to: to.value,
          newHalf: "second",
          newBlockId: createEditorIntentIdAllocator({
            reservedBlockIds: this.options.session.reservedBlockIds,
          })(
            readModel(this.options.session.document),
            {
              type: "splitParagraph",
              at: from.value,
              to: to.value,
            },
            this.options.session.idLease,
          ).newBlockId,
        };
        break;
      case "mergeBlockWithNext": {
        const next = projection.value.nextParagraphInContainer(blockId);
        // The block exists; it ends its container, so there is nothing to
        // merge it with.
        if (next === undefined)
          return skipAlike(
            "unsupportedBlock",
            "The block ends its container; no paragraph follows it there.",
          );
        // The merged paragraph is the one whose mark remains, the second,
        // as accepting the tracked merge leaves it.
        intent = {
          type: "joinParagraphs",
          story,
          blockId,
          nextBlockId: next.blockId,
          survivor: directJoinSurvivor(),
        };
        break;
      }
      case "formatRange": {
        const formatting = operation.formatting;
        const patch = {
          ...(formatting.bold === undefined ? {} : { bold: formatting.bold }),
          ...(formatting.italic === undefined
            ? {}
            : { italic: formatting.italic }),
          ...(formatting.underline === undefined
            ? {}
            : {
                underline:
                  formatting.underline === null
                    ? null
                    : { style: formatting.underline ? "single" : "none" },
              }),
          ...(formatting.strike === undefined
            ? {}
            : { strike: formatting.strike, doubleStrike: null }),
          ...(formatting.fontSizePt === undefined
            ? {}
            : {
                fontSize:
                  formatting.fontSizePt === null
                    ? null
                    : formatting.fontSizePt * 2,
              }),
          ...(formatting.fontFamily === undefined
            ? {}
            : {
                fontFamily:
                  formatting.fontFamily === null
                    ? null
                    : {
                        ascii: formatting.fontFamily,
                        hAnsi: formatting.fontFamily,
                      },
              }),
          ...(formatting.highlight === undefined
            ? {}
            : { highlight: formatting.highlight }),
          ...(formatting.color === undefined
            ? {}
            : {
                color:
                  formatting.color === null
                    ? null
                    : {
                        rgb: formatting.color.replace(/^#/u, "").toUpperCase(),
                      },
              }),
        } as const satisfies RunPropsPatch;
        intent = { type: "formatRun", from: from.value, to: to.value, patch };
        break;
      }
      case "replaceInBlock":
      case "replaceRange":
      case "replaceBlock":
        intent = {
          type: "replaceText",
          from: from.value,
          to: to.value,
          text,
          ...(options.replacementBackground === "keep"
            ? {}
            : { runPropsPatch: { highlight: null, shading: null } }),
        };
        break;
      default: {
        const unreachable: never = operation;
        return panic(`Unknown canonical public operation ${unreachable}`);
      }
    }
    const intents: EditorIntent[] = [];
    const normalizations: FolioAIEditNormalization[] = [];
    let deletionSpans: PreparedOperation["deletionSpans"];
    if (intent.type === "replaceText") {
      // The shared plan, as the default engine writes it: the direct edit
      // writes the smallest change, a redline cuts its changes for the reader
      // and keeps every field the direct edit keeps. Each piece of prose
      // between kept references changes on its own, so no change reaches a
      // reference; one that edits references takes each it touches whole.
      const replacementSource = {
        source: clean.text.slice(start, end),
        replacement: text,
        noteReferences: references,
        editsReferences,
        fields: clean.structuralBoundaries.flatMap((boundary) =>
          boundary.type === "field"
            ? [{ offset: boundary.offset - start, length: boundary.length }]
            : [],
        ),
      };
      // Emphasis and dropped formatting rewrite the whole match, so the
      // words they write stand together where it stood, as the default
      // engine rebuilds them.
      let planChanges: (
        source: string,
        target: string,
      ) => readonly { start: number; end: number; text: string }[] =
        planTextChanges;
      if (tracked)
        planChanges = (source, target) =>
          changesFromSegments(diff(source, target));
      if (rewritesWhole) planChanges = rewriteWhole;
      let planned = planReplacementChanges({
        ...replacementSource,
        planChanges,
        ...(tracked ? { keepFieldsOf: planTextChanges } : {}),
      });
      // A comment goes with exactly the words the edit takes, so its fate
      // must not depend on the mode: accepting the tracked edit equals making
      // it directly. A word diff can strike unchanged text at a comment edge
      // and reinsert it outside the range, keep a word the direct cuts take,
      // or align a repeated word to its other occurrence, so that new text
      // lands inside a comment the direct edit writes beside, or the other
      // way round ("done" to "done done" inserts after "done" directly and
      // before it in a word diff). Where either plan changes words under a
      // comment or writes text inside or at the edge of one, the tracked edit
      // takes the direct cuts: a comment the paragraph holds, or one the batch
      // adds, which stands before the edit runs.
      const addedComments = tracked
        ? batchCommentSpans(batchComments, blockId, clean.text)
        : [];
      const commentedAt = (index: number) => {
        const position = clean.offsets[index];
        return (
          position !== undefined &&
          paragraph.node
            .nodeAt(position - paragraph.start)
            ?.marks.some((mark) => mark.type.name === "comment") === true
        );
      };
      // A gap between two characters a comment covers, or at its edge.
      const commentTouchesGap = (gap: number) =>
        commentedAt(gap - 1) ||
        commentedAt(gap) ||
        addedComments.some((span) => span.start <= gap && gap <= span.end);
      const touchesComment = (
        changes: readonly { start: number; end: number; text: string }[] | null,
      ) =>
        changes?.some((change) => {
          const changeStart = start + change.start;
          const changeEnd = start + change.end;
          for (let index = changeStart; index < changeEnd; index++)
            if (commentedAt(index)) return true;
          if (
            changeEnd > changeStart &&
            addedComments.some(
              (span) => span.start < changeEnd && changeStart < span.end,
            )
          )
            return true;
          return (
            change.text.length > 0 &&
            (commentTouchesGap(changeStart) || commentTouchesGap(changeEnd))
          );
        }) === true;
      if (tracked && touchesComment([{ start: 0, end: end - start, text }])) {
        const directCuts = planReplacementChanges({
          ...replacementSource,
          planChanges: planTextChanges,
        });
        if (touchesComment(planned) || touchesComment(directCuts))
          planned = directCuts;
      }
      // The segments above keep the references, or the replacement edits
      // them, so a plan always exists and reproduces the replacement.
      if (planned === null)
        return panic("A replacement's planned changes must reproduce it.");
      // Keep repeated tabs and breaks in place. A replacement that changes
      // the separators passes them to the editor as inline elements.
      const separated = cutAroundRepeatedSeparators(
        replacementSource.source,
        planned,
      );
      const changes = separated.map((change) => ({
        start: start + change.start,
        end: start + change.end,
        text: change.text,
      }));
      const placeChanges = (
        pieces: readonly { start: number; end: number; text: string }[],
      ) => {
        const placed = [];
        for (const change of pieces) {
          const gap = resolveCleanTextRange({
            cleanBlock: clean,
            startOffset: change.start,
            endOffset: change.end,
          });
          if (!gap) return undefined;
          // Keep trailing references and struck text outside the replacement;
          // a reference the change ends with goes with it whole.
          const lastCharacter =
            change.end > change.start &&
            !endsWithNoteReference(clean, change.start, change.end)
              ? clean.offsets[change.end - 1]
              : undefined;
          // A character an element reads as, such as a field's result, ends
          // where the element does.
          const lastHolder =
            lastCharacter === undefined
              ? null
              : paragraph.node.nodeAt(lastCharacter - paragraph.start);
          const afterLast =
            lastCharacter === undefined
              ? undefined
              : lastCharacter +
                (lastHolder === null || lastHolder.isText
                  ? 1
                  : lastHolder.nodeSize);
          placed.push({
            ...change,
            from: gap.from,
            to:
              afterLast === undefined
                ? gap.to
                : Math.max(gap.from, Math.min(gap.to, afterLast)),
          });
        }
        return placed;
      };
      const addressable = (
        placed: readonly { from: number; to: number }[],
      ): boolean =>
        placed.every(
          ({ from: changeFrom, to: changeTo }) =>
            projection.value
              .addressAt(
                changeFrom,
                changeFrom < changeTo ? "afterMarkers" : undefined,
              )
              .isOk() &&
            projection.value
              .addressAt(
                changeTo,
                changeFrom < changeTo ? "beforeMarkers" : undefined,
              )
              .isOk(),
        );
      let replacementChanges = placeChanges(changes);
      if (replacementChanges === undefined)
        return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
      // A change the word diff places inside an element edited only whole,
      // such as a field whose result is in more than one format, replaces
      // each piece of prose whole instead: the replaced text holds the
      // element entirely, or the replacement is refused as cutting into it.
      if (!addressable(replacementChanges)) {
        const whole = placeChanges(
          segments.source.flatMap((piece, index) => {
            const replacement = segments.replacement[index] ?? "";
            if (replacement === piece.text) return [];
            const pieceStart = start + piece.offset;
            return [
              {
                start: pieceStart,
                end: pieceStart + piece.text.length,
                text: replacement,
              },
            ];
          }),
        );
        if (whole === undefined)
          return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
        replacementChanges = whole;
      }
      if (tracked && !suggested)
        deletionSpans = replacementChanges.filter(
          (span) => span.from < span.to,
        );
      // The background goes from every stretch a change touches, its kept
      // characters included, as the default engine clears it: changing one
      // digit of a highlighted amount fills the whole amount. The kept
      // characters are formatted first, while the source offsets still name
      // them; the replacements write their own text without it.
      if (intent.runPropsPatch !== undefined) {
        const nodeAtCharacter = (index: number) => {
          const position = clean.offsets[start + index];
          return position === undefined
            ? null
            : paragraph.node.nodeAt(position - paragraph.start);
        };
        const local = replacementChanges.map((change) => ({
          start: change.start - start,
          end: change.end - start,
          text: change.text,
        }));
        const keptTouched = [
          ...touchedBackgroundStretches(local, {
            length: end - start,
            hasBackground: (index) =>
              nodeAtCharacter(index)?.marks.some((mark) =>
                BACKGROUND_MARK_NAMES.has(mark.type.name),
              ) ?? false,
          }),
        ]
          .filter((index) =>
            local.every(
              (change) => index < change.start || index >= change.end,
            ),
          )
          .toSorted((left, right) => left - right);
        if (
          tracked &&
          keptTouched.some((index) =>
            nodeAtCharacter(index)?.marks.some(
              (mark) =>
                mark.type.name === "runPropertyChange" &&
                expectRunPropertyChangeMarkAttrs(mark).changes.length > 0,
            ),
          )
        )
          return refusal("pendingRunPropertyChange");
        const stretches: { start: number; end: number }[] = [];
        for (const index of keptTouched) {
          const last = stretches.at(-1);
          if (last !== undefined && last.end === index) last.end = index + 1;
          else stretches.push({ start: index, end: index + 1 });
        }
        for (const stretch of stretches) {
          const cleared = resolveCleanTextRange({
            cleanBlock: clean,
            startOffset: start + stretch.start,
            endOffset: start + stretch.end,
          });
          const clearedFrom =
            cleared && projection.value.addressAt(cleared.from, "afterMarkers");
          const clearedTo =
            cleared && projection.value.addressAt(cleared.to, "beforeMarkers");
          if (
            !clearedFrom ||
            !clearedTo ||
            clearedFrom.isErr() ||
            clearedTo.isErr()
          )
            return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
          intents.push({
            type: "formatRun",
            from: clearedFrom.value,
            to: clearedTo.value,
            patch: intent.runPropsPatch,
          });
        }
      }
      const directPlan = tracked
        ? planReplacementChanges({
            ...replacementSource,
            planChanges: planTextChanges,
          })
        : null;
      const directChanges =
        directPlan === null
          ? undefined
          : placeChanges(
              cutAroundRepeatedSeparators(
                replacementSource.source,
                directPlan,
              ).map(
                ({ start: localStart, end: localEnd, text: replacement }) => ({
                  start: start + localStart,
                  end: start + localEnd,
                  text: replacement,
                }),
              ),
            );
      const addressedDirectChanges =
        directChanges !== undefined && addressable(directChanges)
          ? directChanges
          : undefined;
      const formattingSources =
        addressedDirectChanges === undefined
          ? undefined
          : replacementRunSources({
              paragraph: paragraph.node,
              paragraphStart: paragraph.start,
              offsets: clean.offsets,
              start,
              end,
              direct: addressedDirectChanges,
              tracked: replacementChanges,
              projection: projection.value,
            });
      for (const change of addressedDirectChanges ?? replacementChanges) {
        const gap = { from: change.from, to: change.to };
        if (change.text.length > 0 && gap.from < gap.to) {
          let first: readonly Mark[] | null = null;
          let uniform = true;
          paragraph.node.nodesBetween(
            gap.from - paragraph.start,
            gap.to - paragraph.start,
            (node) => {
              if (
                !node.isText ||
                node.marks.some((mark) => mark.type.name === "deletion")
              )
                return;
              const marks = node.marks.filter(
                (mark) => mark.type.name !== "insertion",
              );
              if (first === null) first = marks;
              else if (!Mark.sameSet(first, marks)) uniform = false;
            },
          );
          if (!uniform && normalizations.length === 0)
            normalizations.push({
              id: operation.id,
              code: "uniformReplacementFormatting",
            });
        }
      }
      // Where the words the replacement writes begin: its first change's
      // start, which no later change moves.
      let wordsStart: TextPosition | undefined;
      for (const [index, change] of [
        ...replacementChanges.entries(),
      ].reverse()) {
        let gap = { from: change.from, to: change.to };
        const sources = formattingSources?.[index];
        // The direct diff writes new text at the first removed character.
        // When that character is inside a control, delete the tail outside
        // it first so the replacement can continue the original control.
        if (!tracked && change.text !== "" && gap.from < gap.to) {
          const sourceHolder = projection.value.doc.resolve(gap.from);
          if (
            sourceHolder.parent.type.name === "sdt" &&
            gap.to > sourceHolder.after()
          ) {
            const tailStart = sourceHolder.after();
            const tailFrom = projection.value.addressAt(
              tailStart,
              "afterMarkers",
            );
            const tailTo = projection.value.addressAt(gap.to, "beforeMarkers");
            if (tailFrom.isErr() || tailTo.isErr())
              return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
            intents.push({
              type: "replaceText",
              from: tailFrom.value,
              to: tailTo.value,
              text: "",
              inside: controlEdgesAt(projection.value.doc, tailStart, gap.to),
            });
            gap = { from: gap.from, to: sourceHolder.end() };
          }
        }
        // A replaced range takes the editor's selection rule: its edges stay
        // inside the markers beside them, so replacing exactly the words a
        // comment covers keeps the comment on the new text, pending and
        // accepted. An insertion at a point takes the caret's gap.
        const edges =
          gap.from < gap.to
            ? ({ start: "afterMarkers", end: "beforeMarkers" } as const)
            : undefined;
        const changeFrom = projection.value.addressAt(gap.from, edges?.start);
        const changeTo = projection.value.addressAt(gap.to, edges?.end);
        if (changeFrom.isErr() || changeTo.isErr())
          return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
        // A replacement of exactly a content control's text goes in the
        // control, as typing over it does, tracked or not. Text added at the
        // control's edge goes in it only when the operation's range is the
        // control's own: one that reaches past the control adds it beside.
        let inside = controlEdgesAt(projection.value.doc, gap.from, gap.to);
        if (
          gap.from === gap.to &&
          (inside.start === true || inside.end === true)
        ) {
          const edge = projection.value.doc.resolve(gap.from);
          if (range.from < edge.start() || range.to > edge.end()) inside = {};
        }
        if (
          gap.from === gap.to &&
          (change.start === start || change.start === end) &&
          surveyAcceptedReplacedAnnotations(
            state.doc,
            range.from,
            range.to,
          ).carried.some((mark) => mark.type.name === "hyperlink")
        ) {
          // A minimal change at the edge of an exact link match still edits
          // that link. A match spanning linked and unlinked text has no
          // shared hyperlink to continue at its edge.
          inside = {
            ...inside,
            ...(change.start === start ? { start: true } : { end: true }),
          };
        }
        if (index === 0) wordsStart = changeFrom.value;
        intents.push({
          type: "replaceText",
          from: changeFrom.value,
          to: changeTo.value,
          text: change.text,
          ...(sources === undefined ? {} : { formattingSources: sources }),
          ...(Object.keys(inside).length === 0 ? {} : { inside }),
          // Dropped formatting writes the words with none of their own.
          ...(rebuilds ? { runProps: {} } : {}),
          ...(intent.runPropsPatch === undefined
            ? {}
            : { runPropsPatch: intent.runPropsPatch }),
        });
      }
      // The whole match is rewritten, so its words stand together from the
      // first change's start: each span its markers mark takes its emphasis.
      if (wordsStart !== undefined)
        for (const { start: spanStart, end: spanEnd, patch } of emphasized) {
          const { story: wordsStory, blockId: wordsBlockId } = wordsStart;
          intents.push({
            type: "formatRun",
            from:
              spanStart === 0
                ? wordsStart
                : {
                    story: wordsStory,
                    blockId: wordsBlockId,
                    offset: wordsStart.offset + spanStart,
                  },
            to: {
              story: wordsStory,
              blockId: wordsBlockId,
              offset: wordsStart.offset + spanEnd,
            },
            patch,
          });
        }
      if (rebuilds) {
        const patch = droppedParagraphFormatting(paragraph.source.formatting);
        if (Object.keys(patch).length > 0)
          intents.push({
            type: "formatParagraph",
            at: { story, blockId: paragraph.blockId, offset: 0 },
            patch,
          });
      }
      intents.push(...restyle);
    } else if (
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs"
    ) {
      // The separator goes in at the first paragraph's end, before the join;
      // as the merge runs, it goes only where words follow
      // (`mergeIntentsAsItRuns`).
      const separator = operation.separator ?? "";
      if (separator !== "") {
        const at = projection.value.addressAt(
          paragraph.start + paragraph.node.content.size,
        );
        if (at.isErr())
          return refusal("unsupportedBlock", undefined, at.error.ground);
        intents.push({
          type: "replaceText",
          from: at.value,
          to: at.value,
          text: separator,
        });
      }
      intents.push(intent);
      // The merged paragraph's properties: on the paragraph the join leaves,
      // the second, whose mark remains, unless a direct join keeps the first.
      intents.push(
        ...statedPropertiesIntents({
          story,
          blockId:
            !tracked && intent.survivor === "first"
              ? intent.blockId
              : intent.nextBlockId,
          properties: operation.mergedParagraphProperties,
          mintList,
        }),
      );
    } else if (
      operation.type === "splitBlock" &&
      intent.type === "splitParagraph"
    ) {
      if (tracked && !suggested && end > start) deletionSpans = [range];
      intents.push(intent);
      // Each half's properties, as the legacy engine sets them once the
      // split is made: the first half keeps the paragraph's identity, the
      // second takes the new one.
      for (const [half, properties] of [
        [paragraph.blockId, operation.firstParagraphProperties],
        [intent.newBlockId, operation.secondParagraphProperties],
      ] as const)
        intents.push(
          ...statedPropertiesIntents({
            story,
            blockId: half,
            properties,
            mintList,
          }),
        );
    } else intents.push(intent);
    const nextParagraph =
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs"
        ? projection.value.paragraph(intent.nextBlockId)
        : undefined;
    // A merge plans with its second paragraph's address: one the projection
    // cannot align is no paragraph to merge with.
    if (
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs" &&
      nextParagraph === undefined
    )
      return refusal("unsupportedBlock", undefined, INVALID_REQUEST);
    let claimedTo = range.to;
    if (operation.type === "mergeBlockWithNext")
      claimedTo = paragraph.start + paragraph.node.content.size;
    if (nextParagraph)
      claimedTo = nextParagraph.start + nextParagraph.node.content.size;
    return Result.ok({
      ...resolvedComment,
      intents,
      normalizations,
      from: range.from,
      to: claimedTo,
      ...(deletionSpans !== undefined && { deletionSpans }),
    });
  }
}
