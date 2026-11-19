/** Public batches compile against the read projection; only the canonical journal publishes. */
import {
  panic,
  Result,
  TaggedError,
  type TaggedErrorClass,
} from "better-result";
import { Mark, type Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { TableMap } from "prosemirror-tables";

import {
  CONTENT_CONTROL_LOCK,
  lockedControlMessage,
  paragraphNumberingReference,
  type Table,
} from "@stll/docx-core/model";
import {
  applyDocumentOps,
  type BatchBlockAccess,
  type BatchOrderAddress,
  type BlockInsertionPoint,
  compileEditorIntent,
  contentRank,
  createEditorIntentIdAllocator,
  deletionRemovesBlock,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  documentStories,
  type EditorIntent,
  type FinalBreakCandidate,
  finalBreakCarrier,
  findStoryBody,
  formatParaId,
  freshCommentId,
  groupInsertionRuns,
  type IdBlock,
  type IdLease,
  introducedIds,
  type LeasedId,
  leasedIdCensus,
  type NewListInstance,
  type OpStory,
  orderAddressingBeforeRemoval,
  orderBatchOperations,
  packageResourcesOpOf,
  paragraphLength,
  type ParagraphPropsPatch,
  pendingSuggestionIdsIn,
  revisionSpan,
  type RunPropsPatch,
  storyParagraphs,
  type TextPosition,
} from "@stll/docx-core/ops";

import {
  buildSignatureTableNode,
  buildTableNode,
  SUGGESTED_SUPPORTED_OPERATION_TYPES,
} from "../ai-edits/apply";
import { BatchClaims } from "../ai-edits/batch-claims";
import {
  splitsGraphemeCluster,
  splitsSurrogatePair,
} from "../ai-edits/character-boundaries";
import {
  buildCleanBlockText,
  resolveCleanTextRange,
} from "../ai-edits/clean-text";
import { directJoinSurvivor, directJoinTakes } from "../ai-edits/directJoin";
import {
  createExistingRevisionGuard,
  EXISTING_REVISION_CONFLICT_MESSAGE,
} from "../ai-edits/existingRevisions";
import { stripInlineEmphasisMarkers } from "../ai-edits/inline-emphasis";
import { compileMarkdownInsertion } from "../ai-edits/markdownInsertion";
import {
  createBatchListMinter,
  type MintedList,
} from "../ai-edits/newListNumbering";
import {
  cutsIntoNoteReference,
  type NoteReferenceLabels,
  noteReferenceSpansWithin,
  segmentsAroundNoteReferences,
} from "../ai-edits/note-references";
import {
  isDeletedPPrMark,
  joinsIn,
  joinStopsAtTableAfter,
  PENDING_JOIN_INSERTION_REFUSAL,
  pendingJoinParagraphs,
  pendingMarkEditOf,
  pendingMarkEditRefusal,
  pendingParagraphPropertiesRefusal,
} from "../ai-edits/pendingJoin";
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
import type { TableRectangle } from "../ai-edits/table-mutation-plan";
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
import {
  supportsTableInsertionMode,
  type TableInsertionOperation,
} from "../ai-edits/tableInsertionModes";
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
  findUndefinedOperationReference,
  styleDefinitionLookup,
  withOperationHeadingStyles,
} from "../ai-edits/undefinedReferences";
import { wordDiffSessionFromOptions } from "../ai-edits/word-diff";
import type { FolioContentStatedNumbering } from "../compare/content-types";
import {
  type ApplyFolioDocumentOperationsOptions,
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationBatch,
  type FolioDocumentOperationResult,
  type FolioDocumentOperationStory,
  type FolioDocumentOperationUndoHandle,
  type FolioDocumentOperationUndoResult,
  getFolioDocumentOperationIssues,
  getFolioDocumentOperationReceiptsForStory,
  parseFolioDocumentOperationBatch,
  withDistinctSuggestionIds,
} from "../document-operations";
import { expectTableAttrs } from "../prosemirror/attrs";
import { tableAnchor } from "../prosemirror/canonicalTableCommands";
import { paragraphEndsItsContainer } from "../prosemirror/containerFinalParagraph";
import {
  standaloneTableRowsFromProseMirror,
  tableAttrsToFormatting,
} from "../prosemirror/conversion/fromProseDoc";
import { insertionGap } from "../prosemirror/insertionGaps";
import { maxAnnotationIdInDoc } from "../prosemirror/plugins/revisionIds";
import { preservesReplacementCommentAlignment } from "../prosemirror/replacedAnnotations";
import { createStyleResolver } from "../prosemirror/styles/styleResolver";
import {
  insertedColumnWidth,
  omittedSlotTakingColumn,
} from "../prosemirror/tableGridMutation";
import {
  CANONICAL_GAP,
  type CanonicalGap,
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
  canonicalCommentBody,
  compileCanonicalComments,
} from "./canonicalComments";
import { lockViolationBetween } from "./canonicalLockGuard";
import {
  type CanonicalCommit,
  type CanonicalProjection,
  type CanonicalSession,
} from "./canonicalSession";
import { indentationAsRead, spacingSidesAsRead } from "./paragraphPatchAsRead";

type CanonicalBatchPlan = {
  claim: ReturnType<typeof snapshotBatchClaim>;
  address: BatchOrderAddress;
  block: number;
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
    address: { type: "content", position, rank: 1, inputIndex },
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
        address: { type: "annotation", position: from, inputIndex },
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
      return {
        ...plan,
        insertionRun: hasComment || emptyPageBreak ? null : from,
      };
    }
    case "insertTable":
    case "insertSignatureTable":
      return content({
        type: "insertion",
        at: gapAt(operation.position ?? "after"),
      });
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
      return content({
        type: "deleteTableRow",
        deletion: target,
        row: target.table.child(target.rowIndex),
        table: target.table,
      });
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
};

type CanonicalBatchOrderOptions<T extends PlannedOperation> = {
  items: readonly T[];
  doc: PMNode;
};

const orderCanonicalBatch = <T extends PlannedOperation>({
  items,
  doc,
}: CanonicalBatchOrderOptions<T>): T[] => {
  const deleted = new Set(
    items.flatMap(({ operation, plan }) =>
      operation.type === "deleteBlock" ? [plan.block] : [],
    ),
  );
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
  const address = ({ operation, plan }: T): BatchOrderAddress => {
    if (plan.address.type !== "content") return plan.address;
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
  const access = ({ operation, plan }: T): BatchBlockAccess<number> => {
    const claim = plan.claim;
    if (claim.type === "merge")
      return {
        addresses: [plan.block],
        removes: claim.joinsNow
          ? directJoinTakes(plan.block, claim.next)
          : [claim.next],
      };
    if (operation.type === "deleteTable" || operation.type === "deleteTableRow")
      return { addresses: [], removes: [plan.block] };
    if (plan.address.type === "column" || plan.address.type === "cell")
      return { addresses: [], removes: [] };
    if (operation.type === "insertBeforeBlock" && joinedAway.has(plan.block))
      return { addresses: [], removes: [] };
    if (claim.type !== "deleteBlock")
      return { addresses: [plan.block], removes: [] };
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
    const removes = deletionRemovesBlock({
      endsContainer:
        !appended && paragraphEndsItsContainer(at, paragraph.type.name),
      carrier: finalBreakCarrier(before).type,
    });
    return { addresses: [], removes: removes ? [plan.block] : [] };
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
} as const satisfies Record<
  FolioDocumentOperation["type"],
  "compile" | CanonicalGap
>;

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
  if (properties.styleId !== undefined) patch["styleId"] = properties.styleId;
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
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
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
  if (at.isErr()) return refusal("unsupportedBlock");
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

type ProjectedParagraph = NonNullable<
  ReturnType<CanonicalProjection["paragraph"]>
>;

type ResolveBlockInsertionOptions = {
  mintList: ReturnType<typeof createBatchListMinter>;
  operation: BlockInsertion;
  /** The paragraph the insertion goes beside. */
  paragraph: ProjectedParagraph;
  /** The paragraph whose stated properties the new paragraph copies. */
  formattingSource: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

/**
 * Where an insertion beside `paragraph` goes: beside the paragraph, or, for a
 * paragraph in a table cell, beside the outermost table, as the default
 * engine places it. The point names the paragraph on the far side of the
 * table; none stands there when the table starts or ends its container.
 */
const blockInsertionPlace = (
  operation: BlockInsertion,
  paragraph: NonNullable<ReturnType<CanonicalProjection["paragraph"]>>,
  projection: CanonicalProjection,
): { at: BlockInsertionPoint; position: number } | undefined => {
  const side = operation.type === "insertAfterBlock" ? "after" : "before";
  const table = findOutermostTableBoundary(projection.doc, paragraph.start - 1);
  if (table === null)
    return {
      at: { type: side, blockId: paragraph.blockId },
      position:
        side === "after"
          ? paragraph.start + paragraph.node.content.size
          : paragraph.start,
    };
  const position = table[side];
  const resolved = projection.doc.resolve(position);
  const beside = side === "after" ? resolved.nodeAfter : resolved.nodeBefore;
  if (beside?.type.name !== "paragraph") return undefined;
  const neighbour = projection.paragraphStartingAt(
    side === "after" ? position : position - beside.nodeSize,
  );
  if (neighbour === undefined) return undefined;
  // After the table is before the paragraph that follows it, and before it
  // is after the one it follows.
  return {
    at: {
      type: side === "after" ? "before" : "after",
      blockId: neighbour.blockId,
    },
    position,
  };
};

/** Where an insertion beside a paragraph goes, and whose properties it copies. */
type BesidePendingJoin = {
  place: ProjectedParagraph;
  formattingSource: ProjectedParagraph;
};

/** A suggestion reads the document as it stands: beside the paragraph it names. */
const besideAsItStands = (
  paragraph: ProjectedParagraph,
): BesidePendingJoin => ({
  place: paragraph,
  formattingSource: paragraph,
});

type BesidePendingJoinOptions = {
  side: "before" | "after";
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  noteReferences: NoteReferenceLabels;
};

/**
 * Where an insertion beside `paragraph` goes, read as the joined paragraph a
 * pending join makes of it in the document the batch read
 * (`pendingJoinParagraphs`, which the default engine places by too): after
 * the join's last paragraph, or before its first, copying the properties of
 * the paragraph the join leaves, which the joined paragraph has. The batch's
 * other operations do not move it: a split of the paragraph the join leaves
 * puts the block after both halves, as one put after that paragraph goes.
 * Beside a paragraph whose own break is pending removal where no join reads
 * it (after it, when the join leaves no paragraph with words for the block to
 * follow; before it, when the paragraph goes whole), the block goes beside
 * the paragraph itself.
 */
const besidePendingJoin = ({
  side,
  paragraph,
  projection,
  noteReferences,
}: BesidePendingJoinOptions): BesidePendingJoin => {
  const options = {
    doc: projection.doc,
    paragraphPos: paragraph.start - 1,
    noteReferences,
  };
  const join = pendingJoinParagraphs(options);
  const breakGoes = isDeletedPPrMark(paragraph.node.attrs["pPrMark"]);
  const after = side === "after";
  // Where no join reads the paragraph, the block goes beside the paragraph
  // itself, as the default engine places it.
  if (
    breakGoes &&
    (after ? join.last === options.paragraphPos : !joinsIn(options))
  )
    return besideAsItStands(paragraph);
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
 * One new paragraph beside the anchor, as the default engine writes it: the
 * anchor's stated paragraph properties unless `inheritFormatting` is false,
 * with the operation's own on top, holding the text as one plain run.
 */
const resolveBlockInsertion = ({
  operation,
  paragraph,
  formattingSource,
  projection,
  story,
  mintList,
  refusal,
  skipAlike,
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
  if (markdown === undefined && stripInlineEmphasisMarkers(text) !== text)
    return refusal(
      "unsupportedBlock",
      "Canonical public emphasis replacements are unavailable.",
    );
  const inline = operation.lineBreakMode === "inline";
  if (markdown === undefined && /\[\^(?:e)?\d+\]/u.test(text))
    return refusal(
      "unsupportedBlock",
      "Canonical block insertion requires plain text.",
    );
  // Each line its own paragraph, blank lines dropped, as the default engine
  // splits them; inline, line breaks stay in the one paragraph. A tab is a
  // tab element either way.
  const split =
    markdown === undefined && !inline && LINE_BREAK_PATTERN.test(text);
  const lines = split
    ? text.split(LINE_BREAK_PATTERN).filter((line) => line.trim().length > 0)
    : [text];
  const texts = lines.length === 0 ? [""] : lines;
  const formatting: Record<string, unknown> =
    operation.inheritFormatting === false
      ? {}
      : { ...formattingSource.source.formatting };
  for (const [key, value] of Object.entries(
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
  ))
    if (value === null) Reflect.deleteProperty(formatting, key);
    else formatting[key] = value;
  const everyParagraph = operation.formattingScope === "allParagraphs";
  const markdownNumbering = stated.numbering;
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
  const paragraphs =
    markdownParagraphs ??
    texts.map((line, index): Paragraph => {
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
      else if (line !== "")
        content = [{ type: "run", content: inlineText(line) }];
      return {
        type: "paragraph",
        content,
        // The anchor's properties with the operation's set or cleared over them.
        ...(Object.keys(own).length === 0
          ? {}
          : { formatting: own as ParagraphFormatting }),
      };
    });
  const place = blockInsertionPlace(operation, paragraph, projection);
  if (place === undefined)
    return refusal(
      "unsupportedBlock",
      "Canonical insertion beside a table needs a paragraph on that side of it.",
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
  // nothing for it to cover, as the default engine finds.
  if (
    operation.comment !== undefined &&
    texts.every((line) => line.length === 0)
  )
    return refusal("unsupportedBlock");
  const normalizations: FolioAIEditNormalization[] = split
    ? [
        {
          id: operation.id,
          code: "splitMultilineText",
          paragraphCount: texts.length,
        },
      ]
    : [];
  const resolved = { intents, normalizations, from: position, to: position };
  if (operation.comment === undefined)
    return Result.ok({ type: "plain" as const, operation, ...resolved });
  // Where the inserted paragraphs are is known once they are compiled.
  const anchor = { story, blockId: place.at.blockId, offset: 0 };
  return Result.ok({
    type: "commented" as const,
    operation,
    comment: {
      text: operation.comment.text,
      from: anchor,
      to: anchor,
      inserted: true as const,
    },
    ...resolved,
  });
};

type InsertionAnchor =
  | { type: "paragraph"; at: BlockInsertionPoint }
  | { type: "table"; at: BlockInsertionPoint };

type InsertionGapOptions = {
  operation: FolioDocumentOperation;
  intents: readonly EditorIntent[];
  projection: CanonicalProjection;
};

/** The physical anchor and accepted source gap, before this batch inserts. */
const canonicalInsertionGap = ({
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
    if (type === undefined || side === undefined)
      panic("A resolved table insertion needs an explicit anchor and side.");
    anchor = { type, at: { type: side, blockId } };
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
  const boundary =
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
  const requestedSide = isBlockInsertion(operation)
    ? blockSide
    : (operation.position ?? "after");
  return {
    anchor,
    boundary,
    // Lowering an insertion after a table to before its next paragraph
    // changes its physical anchor, not its accepted source gap.
    gap:
      sourceTable === null
        ? insertionGap(projection.doc, boundary, anchor.at.type)
        : insertionGap(
            projection.doc,
            sourceTable[requestedSide],
            requestedSide,
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
      return {
        ...intent,
        at:
          anchor.type === "paragraph"
            ? anchor.at
            : {
                type:
                  anchor.at.type === "before" ? "beforeTable" : "afterTable",
                blockId: anchor.at.blockId,
              },
      };
    }
    if (
      intent.type !== "tableBlock" ||
      intent.operation.type !== DOCUMENT_OP_TYPES.INSERT_TABLE
    )
      return intent;
    const { operation } = intent;
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
  paragraph: ProjectedParagraph;
  projection: CanonicalProjection;
  story: OpStory;
  noteReferences: NoteReferenceLabels;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

/**
 * A table beside the anchor, built as the default engine builds it (from the
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
  paragraph,
  projection,
  story,
  noteReferences,
  refusal,
  skipAlike,
}: ResolveTableInsertionOptions) => {
  const side = operation.position ?? "after";
  const { schema } = projection.doc.type;
  let node: PMNode | null;
  if (operation.type === "insertTable")
    node = buildTableNode({ schema, rows: operation.rows });
  else {
    if (operation.parties.length === 0)
      return skipAlike("emptyOperation", "A signature table needs a party.");
    // A comment across the whole table would cross its cells, as the
    // default engine finds.
    if (operation.comment !== undefined)
      return skipAlike(
        "unsupportedBlock",
        "A signature table takes no comment.",
      );
    node = buildSignatureTableNode({ schema, parties: operation.parties });
  }
  if (node === null) return refusal("unsupportedBlock");
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
  const outer = findOutermostTableBoundary(projection.doc, paragraph.start - 1);
  if (outer !== null) {
    if (
      joinStopsAtTableAfter({
        insertedBlock: "table",
        doc: projection.doc,
        paragraphPos: outer[side],
        noteReferences,
      })
    )
      return refusal(
        PENDING_JOIN_INSERTION_REFUSAL.reason,
        PENDING_JOIN_INSERTION_REFUSAL.message,
      );
    // Beside the outermost table, touching it, as the default engine puts
    // it: addressed by a paragraph of that table's own.
    const outerTable = projection.doc.nodeAt(outer.before);
    const blockId =
      outerTable === null
        ? undefined
        : tableBlockAddress(projection, {
            table: outerTable,
            tableStart: outer.before + 1,
          });
    if (blockId === undefined) return refusal("unsupportedBlock");
    const intent: EditorIntent = {
      type: "tableBlock",
      operation: {
        type: DOCUMENT_OP_TYPES.INSERT_TABLE,
        story,
        blockId,
        table,
        side,
        anchor: "table",
      },
    };
    return Result.ok({
      type: "plain" as const,
      operation,
      intents: [intent],
      normalizations: [],
      from: outer[side],
      to: outer[side],
    });
  }
  const { place } = besidePendingJoin({
    side,
    paragraph,
    projection,
    noteReferences,
  });
  const $place = projection.doc.resolve(place.start - 1);
  const neighbour = $place.parent.maybeChild(
    $place.index() + (side === "after" ? 1 : -1),
  );
  const gap =
    side === "after" ? place.start - 1 + place.node.nodeSize : place.start - 1;
  if (
    joinStopsAtTableAfter({
      insertedBlock: "table",
      doc: projection.doc,
      paragraphPos: gap,
      noteReferences,
    })
  )
    return refusal(
      PENDING_JOIN_INSERTION_REFUSAL.reason,
      PENDING_JOIN_INSERTION_REFUSAL.message,
    );
  // After a container's last paragraph the container still ends with one.
  return inserted({
    place,
    side,
    paragraphAfter: side === "after" && neighbour === null ? "empty" : "none",
  });
};

/** A paragraph of the table's own that addresses a table edit, by its id. */
const tableBlockAddress = (
  projection: CanonicalProjection,
  row: { table: PMNode; tableStart: number },
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
  return addressed === undefined || addressed.isErr()
    ? undefined
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
 * instead, as the default engine grows it.
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

/**
 * Whether the rectangle is one column of one-cell rows, every cell below
 * the first holding nothing but empty paragraphs: the merge a suggestion
 * records, as merge state alone.
 */
const mergesDownOverEmptyCells = (
  table: PMNode,
  map: TableMap,
  rectangle: TableRectangle,
) => {
  if (rectangle.right - rectangle.left !== 1) return false;
  for (let row = rectangle.top; row < rectangle.bottom; row += 1) {
    const offset = map.map[row * map.width + rectangle.left];
    const cell = offset === undefined ? null : table.nodeAt(offset);
    if (offset === undefined || cell === null) return false;
    const rect = map.findCell(offset);
    if (rect.top !== row || rect.bottom !== row + 1) return false;
    if (row === rectangle.top) continue;
    let empty = true;
    cell.forEach((child) => {
      if (!child.isTextblock || child.content.size > 0) empty = false;
    });
    if (!empty) return false;
  }
  return true;
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
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
  /** Whether the batch records its changes as suggestions. */
  tracked: boolean;
  /** Resolve another paragraph the operation names, in the same story. */
  paragraphOf: (blockId: string) => ProjectedParagraph | undefined;
  /** The default engine skips an operation whose block is in no table alike. */
  notInTable: () => Result<never, CanonicalPublicOperationResolutionError>;
  /** A skip the default engine gives alike: no capability is missing. */
  skipAlike: (
    reason: FolioAIEditSkipReason,
    message: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

/**
 * A table, row or column change, addressed through the paragraph the
 * operation names, as the table edit the editor's own command lowers to. The
 * table it acts on is the innermost one holding that paragraph, as in the
 * default engine. A new row is the row the default engine builds
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
}: ResolveTableStructureOptions) => {
  const { doc } = projection;
  const blockFrom = paragraph.start - 1;
  const row = findEnclosingTableRow(doc, blockFrom);
  if (row === null) return notInTable();
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
            blockId: paragraph.blockId,
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
            blockId: paragraph.blockId,
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
              blockId: paragraph.blockId,
            },
          },
          table.from,
          table.to,
        );
      const column = {
        left: cell.leftColumnIndex,
        right: cell.leftColumnIndex + 1,
      };
      // A suggestion marks whole cells deleted; a cell spanning the column
      // and another has no suggested form narrower than itself, in either
      // engine.
      if (
        tracked &&
        Array.from({ length: map.height }, (_, rowIndex) => {
          const position = map.map[rowIndex * map.width + column.left];
          return position === undefined ? null : row.table.nodeAt(position);
        }).some((node) => Number(node?.attrs["colspan"] ?? 1) !== 1)
      )
        return skipAlike(
          "unsupportedBlock",
          "A suggested column deletion cannot narrow a cell that spans other columns.",
        );
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
        return refusal("unsupportedBlock");
      const address = addressed.value;
      return resolved(
        {
          type: "table",
          operation: {
            type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
            story,
            blockId: address.blockId,
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
      const overflow = describeUnplacedCellTexts({
        cellTexts: operation.cellTexts,
        capacity: getInsertedPhysicalColumnRows(
          TableMap.get(row.table),
          row.table,
          insertion.columnIndex,
        ).length,
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
        return skipAlike(
          "unsupportedBlock",
          "A suggested column cannot widen a cell that spans other columns.",
        );
      const blockId = tableBlockAddress(projection, row);
      if (blockId === undefined) return refusal("unsupportedBlock");
      const intent: EditorIntent = {
        type: "table",
        operation: {
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          story,
          blockId,
          column: insertion.columnIndex,
          // The neighbour's width, as the default engine gives it; a table
          // that states no grid keeps none, whatever the width.
          width: insertedColumnWidth(row.table, insertion.columnIndex) ?? 1,
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
      // A suggestion records a merge down over empty cells; a merge that
      // moves content between cells has no tracked form.
      if (tracked && !mergesDownOverEmptyCells(row.table, map, rectangle))
        return skipAlike(
          "unsupportedBlock",
          "A suggested merge takes one column over empty cells.",
        );
      const blockId = tableBlockAddress(projection, row);
      if (blockId === undefined) return refusal("unsupportedBlock");
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
      // A suggested split gives a merge down its rows back; a cell spanning
      // columns has no suggested split in either engine.
      if (tracked && cell.rightColumnIndex - cell.leftColumnIndex !== 1)
        return skipAlike(
          "unsupportedBlock",
          "A suggested split takes a cell merged down one column.",
        );
      // The cell is addressed by its own paragraph, the merge's first.
      return resolved(
        {
          type: "table",
          operation: {
            type: DOCUMENT_OP_TYPES.SPLIT_CELL,
            story,
            blockId: paragraph.blockId,
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
      });
      if (built.type !== "applied") return refusal("unsupportedBlock");
      const table = built.transaction.doc.nodeAt(
        built.transaction.mapping.map(row.tablePosition, -1),
      );
      if (table === null || table.type.name !== "table")
        return refusal("unsupportedBlock");
      const index = side === "before" ? row.rowIndex : row.rowIndex + 1;
      const inserted = standaloneTableRowsFromProseMirror(table)[index];
      if (inserted === undefined) return refusal("unsupportedBlock");
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
            blockId: paragraph.blockId,
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
 * The operations a canonical suggestion can carry: the default engine's;
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
  /**
   * Whether the batch inserts a paragraph beside this one, which no
   * paragraph follows: the insertion applies first, and the deletion then
   * finds this paragraph by identity.
   */
  gainsFollower: boolean;
  refusal: (
    reason: FolioAIEditSkipReason,
    message?: string,
  ) => Result<never, CanonicalPublicOperationResolutionError>;
};

type CanonicalPublicOperationsOptions = {
  session: CanonicalSession;
  getState: (story: OpStory) => EditorState | null;
  publish: (commit: CanonicalCommit) => boolean;
};

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
  allocations?: readonly ReturnType<
    ReturnType<typeof createEditorIntentIdAllocator>
  >[];
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
  allocations: admittedAllocations,
}: CompilePublicIntentsOptions) => {
  const ops: DocumentOp[] = [];
  const revisions: number[] = [];
  // A split's new half takes the id allocated here, not the one it resolved
  // with: a later intent naming that half names the allocated one.
  const allocated = new Map<string, string>();
  const allocations: ReturnType<
    ReturnType<typeof createEditorIntentIdAllocator>
  >[] = [];
  for (const [intentIndex, sourceIntent] of intents.entries()) {
    const ids =
      admittedAllocations === undefined
        ? allocate(document, sourceIntent, idLease)
        : admittedAllocations.at(intentIndex);
    if (ids === undefined) panic("An admitted intent lost its allocation.");
    allocations.push(ids);
    const leased = idLease === undefined ? {} : { idLease };
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
    const compiled = compileEditorIntent(document, {
      intent,
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
              newIds: {
                control: ids.newIds.control,
                revision:
                  idLease === undefined
                    ? ids.newIds.revision.map(
                        (_, index) => firstRevisionId + index + 1,
                      )
                    : ids.newIds.revision,
              },
            },
    });
    if (compiled.isErr()) return Result.err(compiled.error);
    const applied = applyDocumentOps(document, compiled.value.ops);
    if (applied.isErr()) return Result.err(applied.error);
    document = applied.value.document;
    ops.push(...compiled.value.ops);
    revisions.push(...applied.value.revisions);
    for (const id of applied.value.revisions)
      nextRevisionId = Math.max(nextRevisionId, id + 1);
  }
  return Result.ok({ document, ops, revisions, nextRevisionId, allocations });
};

const CanonicalPublicOperationResolutionErrorBase: TaggedErrorClass<"CanonicalPublicOperationResolutionError"> =
  TaggedError("CanonicalPublicOperationResolutionError");

class CanonicalPublicOperationResolutionError extends CanonicalPublicOperationResolutionErrorBase<{
  message: string;
  reason: FolioAIEditSkipReason;
  /**
   * The capability the refusal waits for, when it is not the compiler's
   * inline one; `null` for an operation the default engine skips alike, such
   * as a row operation whose block is in no table.
   */
  gap?: CanonicalGap | null;
}> {}

type MergeFormattingConflictOptions = {
  before: NonNullable<ReturnType<typeof findStoryBody>>;
  after: NonNullable<ReturnType<typeof findStoryBody>>;
  ops: readonly DocumentOp[];
  author: string;
};

/** Both the tracked join and a subsequent property alignment may write the survivor. */
const overwritesForeignMergeFormatting = ({
  before,
  after,
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
  const pending = storyParagraphs(before).flatMap(({ paragraph }) => {
    if (paragraph.paraId === undefined || !targets.has(paragraph.paraId))
      return [];
    const changes = foreignChanges(paragraph);
    return changes.length === 0 ? [] : [[paragraph.paraId, changes] as const];
  });
  if (pending.length === 0) return false;
  const survivors = new Map(
    storyParagraphs(after).map(({ paragraph }) => [
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
  const body = findStoryBody(document, story);
  const paragraph =
    body === undefined
      ? undefined
      : storyParagraphs(body).find(
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
  options: CanonicalPublicOperationOptions;
  state: EditorState;
  story: OpStory;
  diff: ReturnType<typeof wordDiffSessionFromOptions>["diff"];
  /** Whether the batch records its edits as revisions. */
  tracked: boolean;
  /**
   * Whether the batch records its edits as suggestions, which read the
   * document as it stands rather than as accepting leaves it.
   */
  suggested: boolean;
  /** A new list of one kind for this operation, minted after the batch's earlier ones. */
  mintList: ReturnType<typeof createBatchListMinter>;
  /** The blocks the batch deletes, and those it inserts beside. */
  besideDeletions: BesideDeletions;
};

type BesideDeletions = {
  deleted: ReadonlySet<string>;
  insertedBeside: ReadonlySet<string>;
};

/**
 * The stated numbering a request names, minting its new list once for the
 * operation, as the default engine allocates one instance per operation.
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
    version: number;
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
    const mintList = createBatchListMinter(() => readModel(session.document));
    const { diff } = wordDiffSessionFromOptions(options.wordDiff);
    let nextRevisionId = options.revisionStamp?.idSeed ?? 1;
    if (!session.isComposing && options.revisionStamp === undefined) {
      for (const currentStory of documentStories(readModel(session.document))) {
        const currentProjection = session.projectStory(currentStory);
        if (currentProjection.isOk())
          nextRevisionId = Math.max(
            nextRevisionId,
            maxAnnotationIdInDoc(currentProjection.value.doc) + 1,
          );
      }
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
    // A refused resolution names its gap: its own, or for a block the
    // compiler does not support, the unsupported public payload gap.
    const refuseResolution = (
      id: string,
      error: CanonicalPublicOperationResolutionError,
    ) => {
      const gap =
        error.gap === null
          ? undefined
          : (error.gap ??
            (error.reason === "unsupportedBlock"
              ? CANONICAL_GAP.publicUnsupportedInline
              : undefined));
      if (gap !== undefined) refusals.set(id, { gap });
      skip(id, error.reason, error.message);
    };
    const result = (
      applied: FolioDocumentOperationResult["applied"],
      undoHandle: FolioDocumentOperationUndoHandle | null,
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
        status: batchStatus(batch, ordered.length),
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
    const besideDeletions: BesideDeletions = {
      deleted: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "deleteBlock" ? [operation.blockId] : [],
        ),
      ),
      insertedBeside: new Set(
        batch.operations.flatMap((operation) =>
          operation.type === "insertAfterBlock" ||
          operation.type === "insertBeforeBlock"
            ? [operation.blockId]
            : [],
        ),
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
    const overlapsResolved = (candidate: ReturnType<typeof planned>) =>
      claims.conflictOf(candidate.plan.claim) !== null;
    const admit = (candidate: ReturnType<typeof planned>) => {
      claims.add(candidate.operation.id, candidate.plan.claim);
      resolved.push(candidate);
    };
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
      if (
        options.tableTemplates !== undefined ||
        options.undefinedReferences === "keep"
      ) {
        refusals.set(operation.id, {
          gap: CANONICAL_GAP.publicUnsupportedInline,
        });
        skip(
          operation.id,
          "unsupportedBlock",
          "This canonical compiler does not support the supplied execution options.",
        );
        continue;
      }
      // The same reference check the default engine runs, in the same
      // place: before the operation compiles.
      const referenceSet = references();
      const headings = withOperationHeadingStyles(referenceSet.definitions, [
        operation,
      ]);
      const undefinedReference = findUndefinedOperationReference({
        operation,
        numberingInstanceIds: referenceSet.numberingInstanceIds,
        styles:
          headings === referenceSet.definitions
            ? referenceSet.styles
            : styleDefinitionLookup(createStyleResolver(headings)),
        undefinedReferences: "refuse",
      });
      if (undefinedReference !== undefined) {
        skip(
          operation.id,
          undefinedReference.reason,
          undefinedReference.message,
        );
        continue;
      }
      const disposition =
        CANONICAL_PUBLIC_OPERATION_DISPOSITIONS[operation.type];
      if (disposition !== "compile") {
        refusals.set(operation.id, {
          gap: disposition,
        });
        skip(operation.id, "unsupportedBlock");
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
      });
      if (compiled.isErr()) {
        refuseResolution(operation.id, compiled.error);
        continue;
      }
      const candidate = planned(compiled.value);
      if (refusesExistingRevision(candidate)) continue;
      if (overlapsResolved(candidate)) {
        skip(operation.id, "overlappingOperation");
        continue;
      }
      admit(candidate);
    }
    // An operation resolved beside a deleted paragraph that ends its run
    // counts on its other half; where that half did not resolve, it resolves
    // again on its own, against the operations that did.
    for (const item of [...resolved]) {
      if (item.besideDeletion !== true || !state) continue;
      const { operation } = item;
      if (
        operation.type !== "deleteBlock" &&
        operation.type !== "insertAfterBlock" &&
        operation.type !== "insertBeforeBlock"
      )
        continue;
      const partnered = resolved.some(({ operation: other }) =>
        operation.type === "deleteBlock"
          ? (other.type === "insertAfterBlock" ||
              other.type === "insertBeforeBlock") &&
            other.blockId === operation.blockId
          : other.type === "deleteBlock" && other.blockId === operation.blockId,
      );
      if (partnered) continue;
      resolved.splice(resolved.indexOf(item), 1);
      // Conditional resolution may retire an accepted claim; rebuild only at
      // this boundary, so its old target cannot reject the independent retry.
      claims = new BatchClaims();
      for (const earlier of resolved)
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
        besideDeletions: {
          deleted: new Set(),
          insertedBeside: new Set(),
        },
      });
      if (alone.isErr()) {
        refuseResolution(operation.id, alone.error);
        continue;
      }
      const candidate = planned(alone.value);
      if (refusesExistingRevision(candidate)) continue;
      if (overlapsResolved(candidate)) {
        skip(operation.id, "overlappingOperation");
        continue;
      }
      admit(candidate);
    }
    // A move id names one relocation, a move only when it reaches exactly one
    // source and one destination, as the default engine pairs them; an
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
    });
    let document = readModel(session.document);
    const ops: DocumentOp[] = [];
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
      counts.set(key, (counts.get(key) ?? 0) + 1);
    };
    type PrepareOperationOptions = {
      document: Document;
      resolvedOperation: AdmittedOperation;
      columnInsertions: ReadonlyMap<string, number>;
      nextRevisionId: number;
      idLease?: IdLease;
      allocations?: CompilePublicIntentsOptions["allocations"];
      reservedBlockIds?: ReadonlySet<string>;
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
    // Preparation owns no session state: even comments and lock refusals are
    // facts until the batch chooses this attempt.
    const prepareOperation = ({
      document: sourceDocument,
      resolvedOperation,
      columnInsertions,
      nextRevisionId: firstRevisionId,
      idLease = session.idLease,
      allocations,
      reservedBlockIds = session.reservedBlockIds,
      commentAllocation: admittedComment,
    }: PrepareOperationOptions) => {
      const { operation, intents } = resolvedOperation;
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
      // A suggested operation's records are one suggestion: its own id, or the
      // one it names.
      const suggestionId =
        batch.mode === "suggested"
          ? (("suggestionId" in operation
              ? operation.suggestionId
              : undefined) ?? operation.id)
          : undefined;
      const column = resolvedOperation.plan.address;
      const mappedIntents =
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
      const headings = withOperationHeadingStyles(
        sourceDocument.package.styles,
        [operation],
      );
      const withHeadings =
        headings === undefined || headings === sourceDocument.package.styles
          ? sourceDocument
          : {
              ...sourceDocument,
              package: { ...sourceDocument.package, styles: headings },
            };
      const compiled = compilePublicIntents({
        document: withHeadings,
        intents: mappedIntents,
        allocate,
        reservedBlockIds,
        idLease: idLease,
        nextRevisionId: firstRevisionId,
        mode: batch.mode === "direct" ? "direct" : "tracked-changes",
        suggestionId,
        author: options.author ?? "AI",
        date,
        ...(allocations === undefined ? {} : { allocations }),
      });
      if (compiled.isErr()) {
        return declined({
          reason: "unsupportedBlock",
          message: compiled.error.message,
          refusal: { gap: CANONICAL_GAP.publicUnsupportedInline },
        });
      }
      let nextDocument = compiled.value.document;
      const operationOps = [...compiled.value.ops];
      if (
        operation.type === "mergeBlockWithNext" &&
        conflictsWithRevision !== undefined
      ) {
        const beforeStory = findStoryBody(sourceDocument, story);
        const afterStory = findStoryBody(nextDocument, story);
        if (
          beforeStory !== undefined &&
          afterStory !== undefined &&
          overwritesForeignMergeFormatting({
            before: beforeStory,
            after: afterStory,
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
      switch (resolvedOperation.type) {
        case "plain":
          break;
        case "commented": {
          const { comment } = resolvedOperation;
          const id =
            admittedComment === undefined
              ? freshCommentId(nextDocument, idLease)
              : Result.ok(admittedComment.id);
          // A suggested operation's comment is proposed with it: pending,
          // never saved, accepted and rejected with the suggestion.
          const proposedBy = suggestionId === undefined ? {} : { suggestionId };
          if (comment.inserted === true) {
            // The comment covers every paragraph the operation inserted, at
            // any depth: in a content control or a table cell as much as in
            // the story's own block list.
            const paragraphsIn = (current: Document) => {
              const body = findStoryBody(current, story);
              return body === undefined
                ? []
                : storyParagraphs(body).map(({ paragraph }) => paragraph);
            };
            const known = new Set(
              paragraphsIn(sourceDocument).flatMap(({ paraId }) =>
                paraId === undefined ? [] : [paraId],
              ),
            );
            const added = paragraphsIn(nextDocument).filter(
              ({ paraId }) => paraId !== undefined && !known.has(paraId),
            );
            const firstAdded = added.at(0);
            const lastAdded = added.at(-1);
            const commentContent =
              admittedComment?.content ??
              canonicalCommentBody(nextDocument, comment.text, idLease);
            const commented =
              id.isErr() ||
              firstAdded?.paraId === undefined ||
              lastAdded?.paraId === undefined
                ? undefined
                : compileCanonicalComments({
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
                      anchor: {
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
            if (commented === undefined || commented.isErr() || id.isErr()) {
              return declined({
                reason: "unsupportedBlock",
                message:
                  commented?.isErr() === true
                    ? commented.error.message
                    : "The inserted paragraphs cannot carry the comment.",
                refusal: { gap: CANONICAL_GAP.publicComments },
              });
            }
            nextDocument = commented.value.document;
            operationOps.push(...commented.value.ops);
            commentId = id.value;
            commentAllocation = { id: id.value, content: commentContent };
            break;
          }
          // The anchor paragraph at any depth: in a table cell or a content
          // control as much as in the story's own block list.
          const anchorIn = (current: Document) => {
            const body = findStoryBody(current, story);
            return body === undefined
              ? undefined
              : storyParagraphs(body).find(
                  ({ paragraph }) => paragraph.paraId === comment.from.blockId,
                )?.paragraph;
          };
          const before = anchorIn(sourceDocument);
          const after = anchorIn(nextDocument);
          if (id.isErr() || before === undefined || after === undefined) {
            return declined({
              reason: "unsupportedBlock",
              message: id.isErr()
                ? id.error.message
                : "The comment anchor no longer exists.",
              refusal: { gap: CANONICAL_GAP.publicComments },
            });
          }
          // A tracked edit's comment covers what the edit changed, from its
          // first revision to its last, as the default engine anchors it.
          const changed =
            batch.mode === "direct"
              ? undefined
              : revisionSpan(after.content, new Set(compiled.value.revisions));
          const from =
            changed === undefined
              ? comment.from
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
                    comment.to.offset +
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
            canonicalCommentBody(nextDocument, comment.text, idLease);
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
          if (commented.isErr()) {
            return declined({
              reason: "unsupportedBlock",
              message: commented.error.message,
              refusal: { gap: CANONICAL_GAP.publicComments },
            });
          }
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
      const locked = lockViolationBetween(sourceDocument, nextDocument);
      if (locked !== null) {
        return declined({
          reason: CONTENT_CONTROL_LOCK,
          message: `Refused: ${lockedControlMessage(locked)}`,
        });
      }
      return {
        type: "prepared" as const,
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
        revisionIds: compiled.value.revisions,
        commentId,
        suggestionId,
        allocations: compiled.value.allocations,
        commentAllocation,
      };
    };
    type OperationAttempt = ReturnType<typeof prepareOperation>;
    const commitOperation = (
      resolvedOperation: AdmittedOperation,
      prepared: OperationAttempt,
    ) => {
      const {
        operation,
        intents,
        normalizations: adjustments,
      } = resolvedOperation;
      if (prepared.type === "declined") {
        if (prepared.refusal !== undefined)
          refusals.set(operation.id, prepared.refusal);
        skip(operation.id, prepared.reason, prepared.message);
        return;
      }
      const { commentId, suggestionId, revisionIds } = prepared;
      normalizations.push(...adjustments);
      recordColumnInsertion(resolvedOperation, insertedColumnCounts);
      document = prepared.document;
      ops.push(...prepared.ops);
      nextRevisionId = Math.max(nextRevisionId, prepared.nextRevisionId);
      const revisionId = revisionIds.at(0);
      const joined =
        batch.mode === "direct"
          ? joinedParagraphsOf(operation.type, intents)
          : undefined;
      appliedById.set(operation.id, {
        id: operation.id,
        ...(commentId === undefined ? {} : { commentId }),
        ...(joined === undefined ? {} : { paragraphIdentity: joined }),
        ...(revisionId === undefined ? {} : { revisionId, revisionIds }),
        ...(suggestionId === undefined ? {} : { suggestionId }),
      });
    };
    const inputOrder = new Map(
      batch.operations.map(({ id }, index) => [id, index]),
    );
    const orderOf = ({ operation }: AdmittedOperation) => {
      const order = inputOrder.get(operation.id);
      if (order === undefined)
        panic("A resolved insertion lost its batch order.");
      return order;
    };
    const gaps = new Map<number, AdmittedOperation[]>();
    for (const member of executionOrder) {
      if (!placesBlock(member.operation)) continue;
      const place = canonicalInsertionGap({
        operation: member.operation,
        intents: member.intents,
        projection: projected.value,
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
    const windows = new Map<AdmittedOperation, readonly AdmittedOperation[]>();
    for (const interval of merged) {
      const first = executionOrder.at(interval.first);
      if (first === undefined)
        panic("An insertion admission window lost its first operation.");
      windows.set(
        first,
        executionOrder.slice(interval.first, interval.last + 1),
      );
    }
    const completed = new Set<string>();
    for (const resolvedOperation of executionOrder) {
      if (completed.has(resolvedOperation.operation.id)) continue;
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
          }),
        );
        continue;
      }
      type AdmittedMember = {
        resolvedOperation: AdmittedOperation;
        firstRevisionId: number;
        ids: ReturnType<typeof introducedIds>;
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
      for (const member of window) {
        completed.add(member.operation.id);
        const prepared = prepareOperation({
          document: admissionDocument,
          columnInsertions: admissionColumns,
          resolvedOperation: member,
          nextRevisionId: admissionRevisionId,
        });
        if (prepared.type === "declined") {
          declinedMembers.push({ resolvedOperation: member, prepared });
          continue;
        }
        recordColumnInsertion(member, admissionColumns);
        admitted.push({
          resolvedOperation: member,
          firstRevisionId: admissionRevisionId,
          allocations: prepared.allocations,
          ops: prepared.ops,
          commentAllocation: prepared.commentAllocation,
          ids: introducedIds(
            leasedIdCensus(admissionDocument),
            leasedIdCensus(prepared.document),
          ),
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
        const firstPlace = canonicalInsertionGap({
          operation: first.resolvedOperation.operation,
          intents: first.resolvedOperation.intents,
          projection: projected.value,
        });
        if (firstPlace === undefined)
          panic("An admitted insertion lost its source placement.");
        const anchor = firstPlace.anchor;
        const placed = partners.toSorted(
          (left, right) =>
            (orderOf(left.resolvedOperation) -
              orderOf(right.resolvedOperation)) *
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
      const replayColumns = new Map(insertedColumnCounts);
      let replayDocument = document;
      const replayed: {
        resolvedOperation: AdmittedOperation;
        prepared: OperationAttempt;
      }[] = [];
      for (const { member, operation: replayOperation } of ordered) {
        const idLease =
          session.idLease === undefined
            ? undefined
            : insertionMemberIdLease({
                lease: session.idLease,
                siblings: admitted
                  .filter((other) => other !== member)
                  .map(({ ids }) => ids),
              });
        const reservedBlockIds = new Set(session.reservedBlockIds);
        for (const other of admitted) {
          if (other === member) continue;
          for (const { kind, id } of other.ids) {
            if (kind === "paragraph") reservedBlockIds.add(formatParaId(id));
          }
        }
        const droppedCarrier = droppedInsertionCarrier({
          admitted: member.resolvedOperation.intents,
          replayed: replayOperation.intents,
          ops: member.ops,
        });
        if (droppedCarrier !== undefined) reservedBlockIds.add(droppedCarrier);
        const prepared = prepareOperation({
          document: replayDocument,
          columnInsertions: replayColumns,
          resolvedOperation: replayOperation,
          nextRevisionId: member.firstRevisionId,
          reservedBlockIds,
          allocations: member.allocations,
          ...(member.commentAllocation === undefined
            ? {}
            : { commentAllocation: member.commentAllocation }),
          ...(idLease === undefined ? {} : { idLease }),
        });
        if (prepared.type === "declined")
          panic(`Admitted insertion failed when ordered: ${prepared.message}`);
        recordColumnInsertion(replayOperation, replayColumns);
        replayDocument = prepared.document;
        replayed.push({ resolvedOperation: replayOperation, prepared });
      }
      for (const declinedMember of declinedMembers)
        commitOperation(
          declinedMember.resolvedOperation,
          declinedMember.prepared,
        );
      for (const replay of replayed)
        commitOperation(replay.resolvedOperation, replay.prepared);
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
    if (batch.atomic && skipped.length > 0) {
      for (const { id } of appliedById.values())
        skip(id, "atomicBatchRejected");
      return result([], null);
    }
    if (ops.length === 0) return result([], null);
    const selection = projected.value.selectionAt(state);
    const prepared = selection.isErr()
      ? selection
      : session.prepareOps(state, ops, selection.value);
    if (prepared.isErr()) {
      // A batch that changes nothing is a no-op, not a capability gap.
      const noChange = prepared.error.reason === "noChange";
      for (const { id } of appliedById.values()) {
        if (!noChange)
          refusals.set(id, {
            gap: CANONICAL_GAP.publicUnsupportedInline,
          });
        skip(
          id,
          noChange ? "noopOperation" : "unsupportedBlock",
          prepared.error.message,
        );
      }
      return result([], null);
    }
    if (batch.mode === "direct") {
      const finalBody = findStoryBody(prepared.value.document, story);
      if (finalBody === undefined)
        panic("A committed public operation lost its story.");
      const finalIds = new Set(
        storyParagraphs(finalBody).map(({ paragraph }) => paragraph.paraId),
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
    }
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
            }
          : entry,
      ];
    });
    if (batch.dryRun) return result(applied, null);
    if (!this.options.publish(prepared.value)) {
      for (const { id } of applied) skip(id, "documentNotEditable");
      return result([], null);
    }
    const handle = {
      type: "documentOperationUndo",
      id: `canonical-${this.handleNamespace}-${String(this.nextHandle++)}`,
    } as const;
    this.handles.push({ handle, version: session.version, story });
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
    // Undo advances the version while exposing the preceding unchanged journal group.
    const previous = this.handles.at(-1);
    if (previous && previous.version === entry.version - 1)
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
    gainsFollower,
    refusal,
  }: ResolveBlockDeletionOptions) {
    // A direct deletion leaves no text for a comment to cover, as the default
    // engine finds; a tracked one covers the deleted words.
    if (operation.comment !== undefined && !tracked)
      return refusal("unsupportedBlock");
    const end = paragraph.start + paragraph.node.content.size;
    const after = doc.resolve(end + 1).nodeAfter;
    const before = doc.resolve(paragraph.start - 1).nodeBefore;
    const neighbour = (node: PMNode | null) => {
      const id: unknown = node?.attrs["paraId"];
      return node?.type.name === "paragraph" && typeof id === "string"
        ? projection.paragraph(id)
        : undefined;
    };
    const next = neighbour(after);
    const previous = neighbour(before);
    // A tracked deletion's comment covers the deleted paragraph's words.
    const commentFrom = projection.addressAt(paragraph.start);
    const commentTo = projection.addressAt(end);
    if (commentFrom.isErr() || commentTo.isErr())
      return refusal("unsupportedBlock");
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
    if (
      next !== undefined ||
      gainsFollower ||
      (tracked && (after !== null || operation.moveId !== undefined))
    ) {
      const intents: EditorIntent[] = [
        {
          type: "deleteParagraphs",
          story: projectionStory,
          blockIds: [paragraph.blockId],
        },
      ];
      return tracking({
        intents,
        normalizations: [],
        from: paragraph.start,
        to: end,
        ...(after === null && !gainsFollower
          ? {}
          : { removesParagraphBeforeBlock: true }),
      });
    }
    let range: { from: number; to: number } | undefined;
    // Removed directly, a paragraph no paragraph follows loses its words and
    // then joins the paragraph before it, whose properties it keeps; the
    // joined paragraph is the one whose mark remains, this one, as accepting
    // the tracked deletion leaves it. With none before it, it only loses its
    // words.
    if (!tracked) {
      const own = (
        from: number,
        to: number,
      ): Result<EditorIntent[], CanonicalPublicOperationResolutionError> => {
        const start = projection.addressAt(from);
        const stop = projection.addressAt(to);
        return start.isErr() || stop.isErr()
          ? refusal("unsupportedBlock")
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
      const cleared = own(paragraph.start, end);
      if (cleared.isErr()) return cleared;
      const intents: EditorIntent[] = [...cleared.value];
      // The paragraph before it takes over: before a table or a content
      // control the deleted paragraph's mark goes with it, and the one before
      // it keeps its identity (`directJoinSurvivor`).
      if (previous !== undefined)
        intents.push({
          type: "joinParagraphs",
          story: projectionStory,
          blockId: previous.blockId,
          nextBlockId: paragraph.blockId,
          survivor: directJoinSurvivor({ laterMarkRemains: after === null }),
        });
      if (intents.length === 0) return refusal("noopOperation");
      return Result.ok({
        type: "plain" as const,
        operation,
        intents,
        normalizations: [],
        from: previous === undefined ? paragraph.start : previous.start,
        to: end,
      });
    } else if (previous !== undefined)
      range = {
        from: previous.start + previous.node.content.size,
        to: end,
      };
    if (range === undefined)
      return refusal(
        "unsupportedBlock",
        "Canonical deletion of the only paragraph in its container is unavailable.",
      );
    const from = projection.addressAt(range.from);
    const to = projection.addressAt(range.to);
    if (from.isErr() || to.isErr()) return refusal("unsupportedBlock");
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

  private resolve({
    operation,
    options,
    state,
    story,
    diff,
    tracked,
    suggested,
    mintList,
    besideDeletions,
  }: ResolveCanonicalPublicOperationOptions) {
    const refusal = (reason: FolioAIEditSkipReason, message?: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message: message ?? `The operation was refused: ${reason}.`,
        }),
      );
    const skipAlike = (reason: FolioAIEditSkipReason, message: string) =>
      Result.err(
        new CanonicalPublicOperationResolutionError({
          reason,
          message,
          gap: null,
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
            "The canonical session holds no address in this paragraph, such as one in a text box.",
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
    if (operation.type === "insertSignatureTable" && suggested)
      return Result.err(
        new CanonicalPublicOperationResolutionError({
          reason: "unsupportedBlock",
          message:
            "Whole-table signature suggestions need direct acceptance semantics.",
          gap: CANONICAL_GAP.publicTableStructure,
        }),
      );
    const noteReferences = collectNoteReferenceLabels(state.doc);
    const clean = buildCleanBlockText(paragraph.node, paragraph.start - 1, {
      fieldResults: "text",
      noteReferences,
    });
    // A replacement that names a style the paragraph does not have restyles
    // it, as the default engine does, even when its words stay the same.
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
          ...(pendingRefusal.canonicalRefusal === undefined
            ? {}
            : { gap: pendingRefusal.canonicalRefusal.gap }),
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
    const followed =
      state.doc.resolve(paragraph.start + paragraph.node.content.size + 1)
        .nodeAfter?.type.name === "paragraph";
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
      });
    if (
      operation.type === "insertTable" ||
      operation.type === "insertSignatureTable"
    )
      return resolveTableInsertion({
        operation,
        paragraph,
        projection: projection.value,
        story,
        noteReferences,
        refusal,
        skipAlike,
      });
    if (isBlockInsertion(operation)) {
      const besideDeletion =
        !tracked &&
        operation.type === "insertBeforeBlock" &&
        !followed &&
        besideDeletions.deleted.has(blockId);
      // A container's final mark remains on a paragraph inserted before it.
      // Before a table or control, its own mark goes: insert after it first
      // so the block deletion has a following paragraph to receive its text.
      const after = state.doc.resolve(
        paragraph.start - 1 + paragraph.node.nodeSize,
      ).nodeAfter;
      const insertion: BlockInsertion =
        besideDeletion && after !== null
          ? { ...operation, type: "insertAfterBlock" }
          : operation;
      const placed = suggested
        ? besideAsItStands(paragraph)
        : besidePendingJoin({
            side: insertion.type === "insertAfterBlock" ? "after" : "before",
            paragraph,
            projection: projection.value,
            noteReferences,
          });
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
        formattingSource: placed.formattingSource,
        projection: projection.value,
        story,
        mintList,
        refusal,
        skipAlike,
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
        !followed && besideDeletions.insertedBeside.has(blockId);
      const deleted = this.resolveBlockDeletion({
        operation,
        paragraph,
        projection: projection.value,
        projectionStory: story,
        doc: state.doc,
        tracked,
        gainsFollower,
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
        if (operation.preserveFormatting === false)
          return refusal(
            "unsupportedBlock",
            "Canonical replacements that drop the paragraph's formatting are unavailable.",
          );
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
        if (/[\t\r\n]/u.test(operation.separator ?? ""))
          return refusal(
            "unsupportedBlock",
            "Canonical public join separators require plain text.",
          );
        start = clean.text.length;
        end = start;
        break;
      default: {
        const unreachable: never = operation;
        return unreachable;
      }
    }
    if (
      [start, end].some(
        (offset) =>
          splitsSurrogatePair(clean.text, offset) ||
          splitsGraphemeCluster(clean.text, offset),
      )
    )
      return refusal("splitsCharacter");
    // A note reference is not text: a replacement rewrites the prose around
    // the references its range holds, each kept in place, as the shared
    // segmentation pairs them. A range that cuts a reference, an operation
    // that does not replace text, and a replacement that drops, moves or
    // invents a reference are refused.
    const replacesText =
      operation.type === "replaceInBlock" ||
      operation.type === "replaceRange" ||
      operation.type === "replaceBlock";
    const references = noteReferenceSpansWithin(clean, start, end);
    if (
      cutsIntoNoteReference(clean, start, end) ||
      (references.length > 0 && !replacesText)
    )
      return refusal("protectedReference");
    const segments = segmentsAroundNoteReferences(
      clean.text.slice(start, end),
      text,
      references,
    );
    if (segments === null) return refusal("protectedReference");
    const prose = segments.replacement.join("");
    if (stripInlineEmphasisMarkers(text) !== text)
      return refusal(
        "unsupportedBlock",
        "Canonical public emphasis replacements are unavailable.",
      );
    if (/\[\^(?:e)?\d+\]/u.test(prose) || /[\t\r\n]/u.test(prose))
      return refusal(
        "unsupportedBlock",
        "Canonical public replacements require plain text.",
      );
    const range = resolveCleanTextRange({
      cleanBlock: clean,
      startOffset: start,
      endOffset: end,
    });
    if (!range) return refusal("staleRange");
    const from = projection.value.addressAt(range.from);
    const to = projection.value.addressAt(range.to);
    if (from.isErr() || to.isErr()) return refusal("unsupportedBlock");
    const resolvedComment = resolveComment({
      operation,
      from: from.value,
      to: to.value,
    });
    if (
      operation.type === "commentOnBlock" ||
      operation.type === "commentOnRange"
    )
      return Result.ok({
        ...resolvedComment,
        intents: [],
        normalizations: [],
        from: range.from,
        to: range.to,
      });
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
        if (next === undefined) return refusal("missingBlock");
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
    if (intent.type === "replaceText") {
      const changes: { start: number; end: number; text: string }[] = [];
      // Each piece of prose between kept references changes on its own, so
      // no change reaches a reference.
      for (const [index, piece] of segments.source.entries()) {
        let offset = start + piece.offset;
        let pending: { start: number; end: number; text: string } | null = null;
        for (const segment of diff(
          piece.text,
          segments.replacement[index] ?? "",
        )) {
          switch (segment.type) {
            case "equal":
              if (pending) changes.push(pending);
              pending = null;
              offset += segment.text.length;
              break;
            case "del":
              pending ??= { start: offset, end: offset, text: "" };
              offset += segment.text.length;
              pending.end = offset;
              break;
            case "ins":
              pending ??= { start: offset, end: offset, text: "" };
              pending.text += segment.text;
              break;
            default: {
              const unreachable: never = segment.type;
              return panic(`Unknown replacement diff ${unreachable}`);
            }
          }
        }
        if (pending) changes.push(pending);
      }
      const replacementChanges = [];
      for (const change of changes) {
        const gap = resolveCleanTextRange({
          cleanBlock: clean,
          startOffset: change.start,
          endOffset: change.end,
        });
        if (!gap) return refusal("unsupportedBlock");
        // Keep trailing references and struck text outside the replacement.
        const lastCharacter =
          change.end > change.start ? clean.offsets[change.end - 1] : undefined;
        replacementChanges.push({
          ...change,
          from: gap.from,
          to:
            lastCharacter === undefined
              ? gap.to
              : Math.max(gap.from, Math.min(gap.to, lastCharacter + 1)),
          markers: "retained" as const,
        });
      }
      if (
        !preservesReplacementCommentAlignment({
          doc: state.doc,
          source: {
            from: paragraph.start,
            to: paragraph.start + paragraph.node.content.size,
          },
          changes: replacementChanges,
        })
      )
        return Result.err(
          new CanonicalPublicOperationResolutionError({
            reason: "unsupportedBlock",
            gap: CANONICAL_GAP.commentReplacementAlignment,
            message:
              "Canonical replacement cannot preserve the aligned comment ranges.",
          }),
        );
      for (const change of replacementChanges.toReversed()) {
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
          return refusal("unsupportedBlock");
        intents.push({
          type: "replaceText",
          from: changeFrom.value,
          to: changeTo.value,
          text: change.text,
          ...(intent.runPropsPatch === undefined
            ? {}
            : { runPropsPatch: intent.runPropsPatch }),
        });
      }
      intents.push(...restyle);
    } else if (
      operation.type === "mergeBlockWithNext" &&
      intent.type === "joinParagraphs"
    ) {
      // The separator goes in at the first paragraph's end, before the join.
      const separator = operation.separator ?? "";
      if (separator !== "") {
        const at = projection.value.addressAt(
          paragraph.start + paragraph.node.content.size,
        );
        if (at.isErr()) return refusal("unsupportedBlock");
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
      intents.push(intent);
      // Each half's properties, as the default engine sets them once the
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
    });
  }
}
