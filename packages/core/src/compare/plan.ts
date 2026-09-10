/**
 * Turn two block snapshots of one story into the change list a caller reads
 * and the {@link FolioAIEditOperation}s that reproduce the target when every
 * generated tracked change is accepted.
 *
 * Pure: no parsing, no serialization, no clock. Every alignment runs through
 * {@link alignFolioContentStructure}, shared with representation-neutral comparison.
 *
 * ## Aligning by container, not by block
 *
 * A single alignment over every paragraph in the story cannot see structure.
 * It pairs on text and document order, so it will put a paragraph inside a
 * table cell opposite one outside it, or a cell of one row opposite a cell of
 * the next — and rewriting either pair in place leaves the target's text in
 * the wrong container. A whole row that was added or removed likewise arrives
 * as a scatter of unmatched paragraphs, which no block operation can turn back
 * into a row: a block insertion anchored in a table lands beside the table,
 * never as a new row in it.
 *
 * So the story is aligned in three nested passes, each over things that can
 * actually stand in for one another:
 *
 * 1. **Segments.** Maximal runs of body text and of one table each, so a table
 *    is only ever compared against a table.
 * 2. **Rows**, within a paired table segment, so a whole-row change stays whole
 *    and becomes `insertTableRow` / `deleteTableRow`.
 * 3. **Cells**, within a paired row, matched by physical cell index.
 *
 * ## Move detection
 *
 * Move, split/merge, and formatting classification come from the same neutral
 * comparison core used by non-DOCX callers. The operations remain specific to
 * the tracked-change format, but the meaning of the comparison does not fork.
 */

import { panic } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import { createFolioAITextRangeHandle, trailingBodyBlockId } from "../ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIBlockTableLocation,
  FolioAIEditOperation,
  FolioAIEditSnapshot,
} from "../ai-edits/types";
import type { TableCellCoordinate, TableGeometryPairing } from "../ai-edits/table-geometry";
import { getFolioParaIdFromBlockId } from "../types/block-id";
import {
  alignFolioContentStructure,
  contentBlocksShareContainer,
  createFolioContentAlignmentWorkSession,
  groupFolioContentTableRows as groupRows,
  type FolioContentAlignmentStep,
  type FolioContentAlignmentWorkSession,
} from "./content-alignment";
import { inlineFormattingSegments } from "./formatting";
import {
  changedFolioContentParagraphFormatting,
  createContentComparisonWorkSession,
  detectFolioContentMoves,
  detectFolioContentParagraphMarkPlans,
  type FolioContentComparisonWorkSession,
} from "./content";
import type { CompareChange, CompareChangeLocation } from "./types";

/**
 * One step of the interpreted alignment. The DOCX planner retains its
 * historical target-side names while the shared core uses revised-side names.
 */
type CompareStep =
  | { type: "pair"; baseBlock: FolioAIBlock; targetBlock: FolioAIBlock }
  | {
      type: "baseOnly";
      block: FolioAIBlock;
      moveScope: Extract<
        FolioContentAlignmentStep<FolioAIBlock>,
        { type: "baseOnly" }
      >["moveScope"];
    }
  | {
      type: "targetOnly";
      block: FolioAIBlock;
      moveScope: Extract<
        FolioContentAlignmentStep<FolioAIBlock>,
        { type: "revisedOnly" }
      >["moveScope"];
    }
  | { type: "baseRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "targetRow"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "baseTable"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | { type: "targetTable"; blocks: readonly FolioAIBlock[]; location: FolioAIBlockTableLocation }
  | {
      type: "baseColumn";
      blocks: readonly FolioAIBlock[];
      location: FolioAIBlockTableLocation;
      columnIndex: number;
    }
  | {
      type: "targetColumn";
      blocks: readonly FolioAIBlock[];
      location: FolioAIBlockTableLocation;
      columnIndex: number;
      anchor: { blockId: string; position: "after" | "before" };
    };

type BuildStepsOptions = {
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  wholeTableReplacement: "allow" | "avoid";
  workSession: FolioContentAlignmentWorkSession;
};

const folioAIBlockIdStability = ({
  id,
  idStability,
}: FolioAIBlock): "stable" | "positional" =>
  idStability ?? (getFolioParaIdFromBlockId(id) === null ? "positional" : "stable");

const toCompareStep = (
  step: FolioContentAlignmentStep<FolioAIBlock>,
): CompareStep => {
  switch (step.type) {
    case "pair":
      return { type: "pair", baseBlock: step.baseBlock, targetBlock: step.revisedBlock };
    case "baseOnly":
      return { type: "baseOnly", block: step.block, moveScope: step.moveScope };
    case "revisedOnly":
      return { type: "targetOnly", block: step.block, moveScope: step.moveScope };
    case "baseRow":
      return { type: "baseRow", blocks: step.blocks, location: step.location };
    case "revisedRow":
      return { type: "targetRow", blocks: step.blocks, location: step.location };
    case "baseTable":
      return { type: "baseTable", blocks: step.blocks, location: step.location };
    case "revisedTable":
      return { type: "targetTable", blocks: step.blocks, location: step.location };
    case "baseColumn":
      return {
        type: "baseColumn",
        blocks: step.blocks,
        location: step.location,
        columnIndex: step.columnIndex,
      };
    case "revisedColumn":
      return {
        type: "targetColumn",
        blocks: step.blocks,
        location: step.location,
        columnIndex: step.columnIndex,
        anchor: step.anchor,
      };
    default: {
      const unreachable: never = step;
      return panic("Unhandled content alignment step", { step: unreachable });
    }
  }
};

const toContentAlignmentStep = (
  step: CompareStep,
): FolioContentAlignmentStep<FolioAIBlock> => {
  switch (step.type) {
    case "pair":
      return { type: "pair", baseBlock: step.baseBlock, revisedBlock: step.targetBlock };
    case "baseOnly":
      return step;
    case "targetOnly":
      return { type: "revisedOnly", block: step.block, moveScope: step.moveScope };
    case "baseRow":
    case "baseTable":
    case "baseColumn":
      return step;
    case "targetRow":
      return { type: "revisedRow", blocks: step.blocks, location: step.location };
    case "targetTable":
      return { type: "revisedTable", blocks: step.blocks, location: step.location };
    case "targetColumn":
      return {
        type: "revisedColumn",
        blocks: step.blocks,
        location: step.location,
        columnIndex: step.columnIndex,
        anchor: step.anchor,
      };
    default: {
      const unreachable: never = step;
      return panic("Unhandled compare step", { step: unreachable });
    }
  }
};

const isEmptyParagraphNode = (block: FolioAIBlock, snapshot: FolioAIEditSnapshot): boolean => {
  const anchor =
    snapshot.anchors[block.id] ??
    panic("A comparison snapshot block has no matching anchor", { blockId: block.id });
  return block.text === "" && block.table === undefined && anchor.to - anchor.from === 2;
};

type TerminalCarrierPair = Extract<CompareStep, { type: "pair" }>;

/** Align both stories, optionally holding the two last paragraphs out as a pair. */
const alignedSteps = (
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
  terminalCarrierPair: TerminalCarrierPair | null,
  wholeTableReplacement: BuildStepsOptions["wholeTableReplacement"],
  workSession: FolioContentAlignmentWorkSession = createFolioContentAlignmentWorkSession(),
): CompareStep[] => {
  const alignedBaseBlocks = terminalCarrierPair ? baseBlocks.slice(0, -1) : baseBlocks;
  const alignedTargetBlocks = terminalCarrierPair ? targetBlocks.slice(0, -1) : targetBlocks;
  const steps = alignFolioContentStructure({
    baseBlocks: alignedBaseBlocks,
    revisedBlocks: alignedTargetBlocks,
    wholeTableReplacement,
    workSession,
    stableIdMismatch: "pair",
    idStability: folioAIBlockIdStability,
  }).map(toCompareStep);
  if (terminalCarrierPair) {
    steps.push(terminalCarrierPair);
  }
  return steps;
};

/**
 * Whether an alignment left the story's last paragraph removed with no
 * paragraph of its own container surviving in front of it.
 *
 * That paragraph's mark is the one a story cannot lose, and a deleted mark
 * joins two paragraphs of ONE container, so with nothing of that container in
 * front of it there is no chain a removal could run down. The paragraph then
 * stays whatever the plan says, and the story ends on a blank line the target
 * does not have — unless the words the target ends with are written into it,
 * which is what reserving it for the target's last paragraph does.
 */
const terminalCarrierIsStranded = (
  steps: readonly CompareStep[],
  baseBlocks: readonly FolioAIBlock[],
): boolean => {
  const carrier = baseBlocks.at(-1);
  if (carrier === undefined) {
    return false;
  }
  const removed = new Set(
    steps.flatMap((step) => (step.type === "baseOnly" ? [step.block.id] : [])),
  );
  if (!removed.has(carrier.id)) {
    return false;
  }
  const container = containerKeyOf(carrier);
  for (let index = baseBlocks.length - 2; index >= 0; index--) {
    const block = baseBlocks[index];
    if (block === undefined || containerKeyOf(block) !== container) {
      return true;
    }
    if (!removed.has(block.id)) {
      return false;
    }
  }
  return true;
};

/**
 * Whether an alignment ADDED the story's last paragraph with something other
 * than paragraphs between it and the base's.
 *
 * An inserted mark at a container's end has the same problem its deletion
 * does: nothing follows it, so rejecting the addition cannot close the break
 * back over the next paragraph. The applier rotates instead — the paragraph
 * the run was appended after takes the inserted mark, and the last one takes
 * the free mark — and that rotation reaches only across paragraphs. A table
 * among them stops it, and the words the target ends with have to be written
 * into the base's own last paragraph instead.
 */
const addedTerminalCarrierIsStranded = (
  steps: readonly CompareStep[],
  baseBlocks: readonly FolioAIBlock[],
  targetBlocks: readonly FolioAIBlock[],
): boolean => {
  const baseLast = baseBlocks.at(-1);
  const targetLast = targetBlocks.at(-1);
  if (baseLast === undefined || targetLast === undefined) {
    return false;
  }
  if (!steps.some((step) => step.type === "targetOnly" && step.block.id === targetLast.id)) {
    return false;
  }
  const baseLastStep = steps.findIndex(
    (step) =>
      (step.type === "pair" && step.baseBlock.id === baseLast.id) ||
      (step.type === "baseOnly" && step.block.id === baseLast.id),
  );
  return (
    baseLastStep === -1 || steps.slice(baseLastStep + 1).some((step) => step.type !== "targetOnly")
  );
};

const buildSteps = ({
  baseSnapshot,
  targetSnapshot,
  wholeTableReplacement,
  workSession,
}: BuildStepsOptions): CompareStep[] => {
  const baseBlocks = baseSnapshot.blocks;
  const targetBlocks = targetSnapshot.blocks;
  const baseLast = baseBlocks.at(-1);
  const targetLast = targetBlocks.at(-1);
  const carrierPair =
    baseLast !== undefined && targetLast !== undefined
      ? ({ type: "pair", baseBlock: baseLast, targetBlock: targetLast } as const)
      : null;
  // A body's final paragraph mark survives accepting or rejecting a deletion
  // on it: nothing follows it to merge into. When both sides end in a truly
  // empty paragraph node, reserve that structural carrier before general
  // alignment and compare its supported properties normally.
  const bothStoriesEndBlank =
    carrierPair !== null &&
    isEmptyParagraphNode(carrierPair.baseBlock, baseSnapshot) &&
    isEmptyParagraphNode(carrierPair.targetBlock, targetSnapshot);
  const alignmentBudgetBeforeAttempt = workSession.remainingLcsCells;
  const steps = alignedSteps(
    baseBlocks,
    targetBlocks,
    bothStoriesEndBlank ? carrierPair : null,
    wholeTableReplacement,
    workSession,
  );
  // Otherwise the reservation is a repair, not a preference, and the alignment
  // is what says whether it is needed: pairing the two last paragraphs where
  // the ordinary alignment reaches the carrier would trade a plain "this
  // paragraph was removed" for a removal plus a rewrite that reads nothing
  // like the edit. Both ends can need it, because the mark that ends a
  // container carries no revision in either direction.
  //
  // The carrier can only take the target's last paragraph when that paragraph
  // is its opposite number: a story ends with a body paragraph, and a target
  // whose last block sits in a table cell has nothing to hand it.
  if (
    carrierPair === null ||
    bothStoriesEndBlank ||
    !shareAContainer(carrierPair.baseBlock, carrierPair.targetBlock) ||
    !(
      terminalCarrierIsStranded(steps, baseBlocks) ||
      addedTerminalCarrierIsStranded(steps, baseBlocks, targetBlocks)
    )
  ) {
    return steps;
  }
  // The first alignment was only a probe for the terminal-carrier repair.
  // Charge the shared package budget for the plan we keep, not both attempts.
  workSession.remainingLcsCells = alignmentBudgetBeforeAttempt;
  return alignedSteps(
    baseBlocks,
    targetBlocks,
    carrierPair,
    wholeTableReplacement,
    workSession,
  );
};

/**
 * Two blocks are in the same container when a paragraph mark between them
 * exists at all: two body paragraphs, or two paragraphs of one table cell. A
 * mark cannot span a cell boundary, so a split or a merge across one is not a
 * paragraph-mark edit however similar the text looks.
 */
const shareAContainer = (left: FolioAIBlock, right: FolioAIBlock): boolean => {
  return contentBlocksShareContainer(left, right);
};

/**
 * For each step, the id of the next base block at or after it — the anchor a
 * target-only insertion is placed before. `null` once no base block follows.
 */
const nextBaseBlockIdByStep = (steps: readonly CompareStep[]): (string | null)[] => {
  const anchors = Array.from<string | null>({ length: steps.length });
  let next: string | null = null;
  for (let index = steps.length - 1; index >= 0; index--) {
    anchors[index] = next;
    const step = steps[index];
    if (step?.type === "pair") {
      next = step.baseBlock.id;
    } else if (step?.type === "baseOnly") {
      next = step.block.id;
    } else if (step?.type === "baseRow") {
      next = step.blocks[0]?.id ?? next;
    } else if (step?.type === "baseColumn") {
      next = step.blocks[0]?.id ?? next;
    }
  }
  return anchors;
};

/** A base block sitting inside a table, and whether it precedes or follows the step. */
type RowAnchor = { blockId: string; position: "after" | "before" };

const baseTableBlockOf = (step: CompareStep): FolioAIBlock | null => {
  switch (step.type) {
    case "pair":
      return step.baseBlock.table ? step.baseBlock : null;
    case "baseOnly":
      return step.block.table ? step.block : null;
    case "baseRow":
    case "baseTable":
    case "baseColumn":
      return step.blocks[0] ?? null;
    case "targetOnly":
    case "targetRow":
    case "targetTable":
    case "targetColumn":
      return null;
    default: {
      const unreachable: never = step;
      return panic("Unhandled compare step", { step: unreachable });
    }
  }
};

/**
 * The base-document row a new row is inserted next to: the nearest base block
 * inside a table, preferring the one before the insertion so a run of new rows
 * keeps its order.
 */
const findRowAnchor = (steps: readonly CompareStep[], stepIndex: number): RowAnchor | null => {
  for (let index = stepIndex - 1; index >= 0; index--) {
    const step = steps[index];
    const block = step ? baseTableBlockOf(step) : null;
    if (block) {
      return { blockId: block.id, position: "after" };
    }
  }
  for (let index = stepIndex + 1; index < steps.length; index++) {
    const step = steps[index];
    const block = step ? baseTableBlockOf(step) : null;
    if (block) {
      return { blockId: block.id, position: "before" };
    }
  }
  return null;
};

/**
 * A row's text per physical cell, indexed BY cell so an empty cell keeps its
 * slot. An empty cell carries no block at all, so packing only the cells that
 * have text would shift every later cell one column left.
 */
/** One table's cell texts, row by row, for a whole-table change. */
/**
 * One table's cell texts, row by row, padded to the widest row.
 *
 * A row's own width is its highest occupied cell index, so a table with
 * merged cells or a short last row produces a ragged grid — and a ragged grid
 * is not a table any consumer can lay out. Padding states the grid the table
 * actually occupies; the empty strings are the cells a `w:gridSpan` covers.
 */
const tableCellTexts = (blocks: readonly FolioAIBlock[]): string[][] => {
  const rows = groupRows(blocks).map((row) => rowCellTexts(row));
  let width = 0;
  for (const row of rows) {
    width = Math.max(width, row.length);
  }
  for (const row of rows) {
    while (row.length < width) {
      row.push("");
    }
  }
  return rows;
};

const rowCellTexts = (blocks: readonly FolioAIBlock[]): string[] => {
  const byCell: (string | undefined)[] = [];
  for (const block of blocks) {
    const cellIndex = block.table?.cellIndex ?? 0;
    const existing = byCell[cellIndex];
    byCell[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return Array.from(byCell, (text) => text ?? "");
};

const columnCellTexts = (blocks: readonly FolioAIBlock[]): string[] => {
  const byCell = new Map<string, string>();
  for (const block of blocks) {
    const table = block.table;
    if (!table) {
      continue;
    }
    const key = `${String(table.rowIndex)}:${String(table.cellIndex)}`;
    const existing = byCell.get(key);
    byCell.set(key, existing === undefined ? block.text : `${existing}\n${block.text}`);
  }
  return [...byCell.values()];
};

const locationOf = (story: FolioDocumentStoryHandle, block: FolioAIBlock): CompareChangeLocation =>
  block.table ? { story, cell: block.table } : { story };

/**
 * The container a paragraph mark ends: the story's body, or one table cell. A
 * mark joins the paragraph it ends with the next paragraph of the SAME
 * container, so two blocks on either side of a boundary are not neighbours for
 * this purpose however adjacent they read.
 */
const containerKeyOf = ({ table }: FolioAIBlock): string =>
  table
    ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}`
    : "body";

/** Each container's last block, in the order the story holds them. */
const lastBlockByContainer = (
  blocks: readonly FolioAIBlock[],
): ReadonlyMap<string, FolioAIBlock> => {
  const last = new Map<string, FolioAIBlock>();
  for (const block of blocks) {
    last.set(containerKeyOf(block), block);
  }
  return last;
};

type TrailingDeletionOptions = {
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  /** The plan so far, rewritten around each container it empties. */
  operations: readonly FolioAIEditOperation[];
  tableTemplates: readonly CompareTableTemplateRequest[];
  nextOperationId: () => string;
};

/**
 * Where the plan's inserts land for one container's trailing deleted run, and
 * which paragraph the merge chain starts from.
 */
type TrailingRun = {
  /** The deleted blocks that end the container, carrier first. */
  blockIds: readonly string[];
  /** The last SURVIVING paragraph of the container, when its mark is free. */
  chainStart: FolioAIBlock | null;
  /** Indexes into the plan of the paragraph insertions placed in the run. */
  insertIndexes: readonly number[];
  /** Indexes into the plan of tables placed inside the run. */
  tableInsertIndexes: readonly number[];
};

/**
 * Rewrite the plan so a container's final paragraph mark survives the removal
 * of the paragraphs that end it.
 *
 * A deleted paragraph mark means "merge this paragraph into the following
 * one". A container's final paragraph has no following one, so the mark cannot
 * say it: the applier leaves that mark alone, and the removal has to be
 * expressed one paragraph earlier. The carrier is that final paragraph, and
 * every trailing removal resolves onto it.
 *
 * Nothing is inserted where the run was: the chain runs from the last
 * SURVIVING paragraph forward. Its mark goes, each removed paragraph's mark
 * goes with it, and the carrier loses its words and keeps its mark as the
 * paragraph the merged text lands in. A carrier that has no words to lose is
 * not deleted at all: its removal is entirely the marks in front of it, and an
 * operation that would write no revision belongs nowhere in the plan.
 *
 * Something IS inserted there: the inserted paragraphs stand where the removed
 * ones did, and the last of them lands INSIDE the carrier as inserted runs
 * before its kept mark, so the carrier is the target's last paragraph rather
 * than a blank line after it. The other inserted paragraphs carry inserted
 * marks of their own and sit before the carrier. The mark count then works out
 * on its own — one mark deleted for every removed paragraph but the carrier,
 * one inserted for every added paragraph but the last — so the chain is not
 * rotated as well.
 *
 * The carrier keeps its own mark either way, and a paragraph's properties live
 * on its mark, so the merged paragraph ends up with the carrier's. Those
 * properties therefore become the target's, written as `w:pPrChange` so
 * rejecting restores what the carrier had. That is bookkeeping for the merge
 * rather than an edit of its own, so it adds no entry to the change list: what
 * the reader is told is that paragraphs were removed and paragraphs added.
 */
const withTrailingDeletionRules = ({
  baseSnapshot,
  targetSnapshot,
  operations,
  tableTemplates,
  nextOperationId,
}: TrailingDeletionOptions): FolioAIEditOperation[] => {
  const plan = [...operations];
  const deletionIndexByBlockId = new Map<string, number>();
  const insertIndexesByAnchor = new Map<string, number[]>();
  const tableInsertIndexesByAnchor = new Map<string, number[]>();
  // A block whose own mark the plan already moves is not a chain start: a
  // split has put a second paragraph after it, and a merge has spent its mark.
  const markedBlockIds = new Set<string>();
  for (const [index, operation] of plan.entries()) {
    switch (operation.type) {
      case "deleteBlock": {
        deletionIndexByBlockId.set(operation.blockId, index);
        break;
      }
      case "insertBeforeBlock":
      case "insertAfterBlock": {
        const placed = insertIndexesByAnchor.get(operation.blockId) ?? [];
        placed.push(index);
        insertIndexesByAnchor.set(operation.blockId, placed);
        break;
      }
      case "insertTable": {
        const placed = tableInsertIndexesByAnchor.get(operation.blockId) ?? [];
        placed.push(index);
        tableInsertIndexesByAnchor.set(operation.blockId, placed);
        break;
      }
      case "splitBlock":
      case "mergeBlockWithNext": {
        markedBlockIds.add(operation.blockId);
        break;
      }
      default: {
        break;
      }
    }
  }
  if (deletionIndexByBlockId.size === 0) {
    return plan;
  }

  const blocks = baseSnapshot.blocks;
  const indexById = new Map(blocks.map((block, index) => [block.id, index]));
  /**
   * Walk back from a container's last block over the run of deleted
   * paragraphs it ends with. A block of another container in between — a
   * table between two body paragraphs, a table nested in a cell — ends the
   * walk: no mark joins across it.
   */
  const trailingRunOf = (container: string, carrier: FolioAIBlock): TrailingRun => {
    const blockIds: string[] = [];
    const insertIndexes: number[] = [];
    const tableInsertIndexes: number[] = [];
    let chainStart: FolioAIBlock | null = null;
    let index =
      indexById.get(carrier.id) ??
      panic("A container's last block is not in the snapshot it came from", {
        blockId: carrier.id,
      });
    for (; index >= 0; index--) {
      const block = blocks[index];
      if (block === undefined || containerKeyOf(block) !== container) {
        break;
      }
      if (!deletionIndexByBlockId.has(block.id)) {
        chainStart = markedBlockIds.has(block.id) ? null : block;
        break;
      }
      blockIds.push(block.id);
      insertIndexes.push(...(insertIndexesByAnchor.get(block.id) ?? []));
      tableInsertIndexes.push(...(tableInsertIndexesByAnchor.get(block.id) ?? []));
    }
    // The plan emits operations in target order, so the plan's own order is
    // the order the inserted paragraphs have to end up in. Numerically: the
    // walk collects them container-block by container-block, and the default
    // comparator would sort index 10 in front of index 2.
    return {
      blockIds,
      chainStart,
      insertIndexes: insertIndexes.toSorted((left, right) => left - right),
      tableInsertIndexes: tableInsertIndexes.toSorted((left, right) => left - right),
    };
  };

  const targetLastByContainer = lastBlockByContainer(targetSnapshot.blocks);
  const terminalTargetTableIndex = targetSnapshot.blocks.at(-1)?.table?.outerTableIndex;
  const targetTableIndexByOperationId = new Map(
    tableTemplates.map(({ operationId, targetTableIndex }) => [operationId, targetTableIndex]),
  );
  const dropped = new Set<number>();
  const appended: FolioAIEditOperation[] = [];
  for (const [container, carrier] of lastBlockByContainer(blocks)) {
    const carrierDeletionIndex = deletionIndexByBlockId.get(carrier.id);
    if (carrierDeletionIndex === undefined) {
      continue;
    }
    const run = trailingRunOf(container, carrier);
    if (run.tableInsertIndexes.length > 0) {
      const tableInsertIndex =
        run.tableInsertIndexes.length === 1 ? run.tableInsertIndexes.at(0) : undefined;
      const tableInsert = tableInsertIndex === undefined ? undefined : plan[tableInsertIndex];
      if (
        container !== "body" ||
        !isEmptyParagraphNode(carrier, baseSnapshot) ||
        tableInsertIndex === undefined ||
        terminalTargetTableIndex === undefined ||
        tableInsert?.type !== "insertTable" ||
        targetTableIndexByOperationId.get(tableInsert.id) !== terminalTargetTableIndex
      ) {
        // A table within a trailing paragraph run otherwise interrupts the
        // deletion chain. Keep the conservative plan and let verification
        // report that its ownership could not be reproduced.
        continue;
      }
      // The body schema permits the target table to precede `sectPr`
      // directly. Put the base's otherwise-final carrier immediately before
      // that table: its deleted mark can then remove the empty paragraph,
      // while rejecting the table insertion restores the original ending.
      plan[tableInsertIndex] = {
        ...tableInsert,
        blockId: carrier.id,
        position: "after",
      };
      continue;
    }
    const lastInsertIndex = run.insertIndexes.at(-1);
    if (lastInsertIndex === undefined) {
      if (run.chainStart) {
        appended.push({
          id: nextOperationId(),
          type: "mergeBlockWithNext",
          blockId: run.chainStart.id,
        });
      }
      if (isEmptyParagraphNode(carrier, baseSnapshot)) {
        dropped.add(carrierDeletionIndex);
      }
    } else {
      const lastInsert = plan[lastInsertIndex];
      if (
        lastInsert === undefined ||
        (lastInsert.type !== "insertBeforeBlock" && lastInsert.type !== "insertAfterBlock") ||
        lastInsert.moveId !== undefined
      ) {
        // A relocation's destination has to stay an insertion for its source
        // to stay a `w:moveFrom`, so this run keeps the shape it has.
        continue;
      }
      const carrierDeletion = plan[carrierDeletionIndex];
      if (carrierDeletion?.type !== "deleteBlock" || carrierDeletion.moveId !== undefined) {
        continue;
      }
      dropped.add(lastInsertIndex);
      if (carrier.text === lastInsert.text) {
        // The carrier already holds the words the last inserted paragraph
        // brings, so neither half of the exchange is a revision.
        dropped.add(carrierDeletionIndex);
      } else {
        plan[carrierDeletionIndex] = {
          id: carrierDeletion.id,
          type: "replaceBlock",
          blockId: carrier.id,
          text: lastInsert.text,
        };
      }
      for (const index of run.insertIndexes.slice(0, -1)) {
        const insert = plan[index];
        if (insert?.type !== "insertBeforeBlock" && insert?.type !== "insertAfterBlock") {
          continue;
        }
        plan[index] = {
          id: insert.id,
          type: "insertBeforeBlock",
          blockId: carrier.id,
          text: insert.text,
          ...(insert.moveId !== undefined && { moveId: insert.moveId }),
          styleId: insert.styleId ?? null,
          listLevel: insert.listLevel ?? null,
          alignment: insert.alignment ?? null,
          spacing: insert.spacing ?? null,
        };
      }
    }
    // The paragraph the carrier's mark now ends is the target's last one in
    // this container, so the carrier is where its properties have to be.
    const targetCarrier = targetLastByContainer.get(container);
    const properties =
      targetCarrier && changedFolioContentParagraphFormatting(carrier, targetCarrier);
    if (properties) {
      appended.push({
        id: nextOperationId(),
        type: "setBlockParagraphProperties",
        blockId: carrier.id,
        properties,
      });
    }
  }
  return [...plan.filter((_, index) => !dropped.has(index)), ...appended];
};

/**
 * A table the target document already holds, which the operation naming it
 * should place verbatim instead of rebuilding from its cell texts.
 *
 * The plan is pure and sees only block snapshots, so it names the table by the
 * index the snapshot numbers tables with; the caller, which has both
 * documents, resolves the index to the node.
 */
export type CompareTableTemplateRequest = {
  /** The `insertTable` or `insertTableRow` operation this table belongs to. */
  operationId: string;
  /** Index of the table in the TARGET story. */
  targetTableIndex: number;
  /** Set for a row insertion: which of that table's rows to place. */
  targetRowIndex?: number;
};

export type CompareStoryPlan = {
  changes: CompareChange[];
  operations: FolioAIEditOperation[];
  /**
   * Where an operation's table comes from. Kept beside the operations rather
   * than inside them because a table node is not JSON, and the serialized
   * operation contract describes a table by its cell texts.
   */
  tableTemplates: CompareTableTemplateRequest[];
  /**
   * Base cells the alignment put opposite a target cell. Their tables, rows
   * and cells are the ones whose `w:tblPr` / `w:trPr` / `w:tcPr` the caller
   * matches: a table that stayed in place while its widths, shading or header
   * row changed moves no block, so no operation carries the difference.
   */
  tableGeometryPairings: TableGeometryPairing[];
};

const cellCoordinate = ({
  tableIndex,
  rowIndex,
  cellIndex,
}: FolioAIBlockTableLocation): TableCellCoordinate => ({ tableIndex, rowIndex, cellIndex });

/**
 * The cells the alignment paired, one entry each. A cell holds several
 * paragraphs and each pairs on its own, so the first pairing of a cell is the
 * one kept: later ones would name the same two cells.
 */
const tableGeometryPairingsOf = (steps: readonly CompareStep[]): TableGeometryPairing[] => {
  const pairings: TableGeometryPairing[] = [];
  const seen = new Set<string>();
  for (const step of steps) {
    if (step.type !== "pair") {
      continue;
    }
    const base = step.baseBlock.table;
    const target = step.targetBlock.table;
    if (!base || !target) {
      continue;
    }
    const key = `${String(base.tableIndex)}:${String(base.rowIndex)}:${String(base.cellIndex)}`;
    if (seen.has(key)) {
      continue;
    }
    seen.add(key);
    pairings.push({ base: cellCoordinate(base), target: cellCoordinate(target) });
  }
  return pairings;
};

export type PlanStoryCompareOptions = {
  story: FolioDocumentStoryHandle;
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  /** Cap on generated operations; the caller turns `null` into its own error. */
  maxOperations: number;
  /** Avoid structural fallback when copying the target table would lose package-bound content. */
  wholeTableReplacement?: "allow" | "avoid";
  /** Shared comparison budget across every story in one package. */
  workSession?: FolioContentComparisonWorkSession;
};

/**
 * Plan one story's comparison, or `null` when it needs more operations than
 * `maxOperations`.
 */
export const planStoryCompare = ({
  story,
  baseSnapshot,
  targetSnapshot,
  maxOperations,
  wholeTableReplacement = "allow",
  workSession = createContentComparisonWorkSession(),
}: PlanStoryCompareOptions): CompareStoryPlan | null => {
  // A tracked table insertion needs a surviving body paragraph as its anchor.
  // This is an application constraint, not a neutral alignment constraint.
  const effectiveWholeTableReplacement =
    wholeTableReplacement === "allow" && trailingBodyBlockId(baseSnapshot) === null
      ? "avoid"
      : wholeTableReplacement;
  const steps = buildSteps({
    baseSnapshot,
    targetSnapshot,
    wholeTableReplacement: effectiveWholeTableReplacement,
    workSession: workSession.alignment,
  });
  const contentSteps = steps.map(toContentAlignmentStep);
  const paragraphMarkPlans = detectFolioContentParagraphMarkPlans(contentSteps);
  // The step after each paragraph-mark plan is part of it, so neither the move
  // pass nor the main loop may claim it again.
  const consumedSteps = new Set([...paragraphMarkPlans.keys()].map((index) => index + 1));
  const movesByBaseBlockId = new Map(
    detectFolioContentMoves({
      steps: contentSteps,
      consumedStepIndexes: consumedSteps,
      workSession,
      idStability: folioAIBlockIdStability,
    }).map(({ baseBlock, revisedBlock }) => [baseBlock.id, revisedBlock.id] as const),
  );
  const moveSourceByTargetBlockId = new Map<string, string>();
  for (const [baseBlockId, targetBlockId] of movesByBaseBlockId) {
    moveSourceByTargetBlockId.set(targetBlockId, baseBlockId);
  }

  const anchorIds = nextBaseBlockIdByStep(steps);
  const changes: CompareChange[] = [];
  const operations: FolioAIEditOperation[] = [];
  const tableTemplates: CompareTableTemplateRequest[] = [];
  /**
   * The anchor everything past the base document's content hangs from: its
   * last BODY-LEVEL paragraph, when it has one. Anchoring to the last block
   * instead puts the anchor inside a terminal table, and a paragraph insertion
   * then escapes to the table's boundary where no paragraph mark can express
   * the break it added.
   *
   * The applier orders insertions that resolve to one position by their order
   * in this array, so the tail is emitted where its step sits rather than
   * collected and appended — a table and a paragraph both added after the last
   * base block otherwise come out in operation order, which is not target
   * order.
   */
  const tailAnchorId = trailingBodyBlockId(baseSnapshot);

  let operationSequence = 0;

  const nextOperationId = (): string => `compare-${++operationSequence}`;

  /**
   * The relocation this block belongs to, named so the applier can link the
   * deletion at the source with the insertion at the destination as
   * `w:moveFrom` and `w:moveTo` instead of writing two unrelated revisions.
   */
  const moveIdOf = (baseBlockId: string): string => `move-${baseBlockId}`;

  const pushInsertOperation = (block: FolioAIBlock, anchorId: string | null): void => {
    const moveSourceId = moveSourceByTargetBlockId.get(block.id);
    // All four always explicit, `null` included: an inserted paragraph that
    // says nothing takes the anchor's paragraph properties, and the anchor is
    // whichever block happened to follow it. A new ordinary paragraph beside
    // a styled, aligned list item is not implicitly the same kind of paragraph.
    const shared = {
      text: block.text,
      ...(moveSourceId !== undefined && { moveId: moveIdOf(moveSourceId) }),
      styleId: block.styleId ?? null,
      listLevel: block.listLevel ?? null,
      alignment: block.directAlignment ?? null,
      spacing: block.directSpacing ?? null,
    };
    if (anchorId !== null) {
      operations.push({
        id: nextOperationId(),
        type: "insertBeforeBlock",
        blockId: anchorId,
        ...shared,
      });
      return;
    }
    if (tailAnchorId === null) {
      // An empty base with no anchor paragraph cannot receive tracked
      // insertions at all.
      return;
    }
    operations.push({
      id: nextOperationId(),
      type: "insertAfterBlock",
      blockId: tailAnchorId,
      ...shared,
    });
  };

  for (const [stepIndex, step] of steps.entries()) {
    if (consumedSteps.has(stepIndex)) {
      continue;
    }
    const paragraphMarkPlan = paragraphMarkPlans.get(stepIndex);
    if (paragraphMarkPlan?.type === "split") {
      const { baseBlock, revisedBlocks: splitInto, offset, separator } = paragraphMarkPlan;
      changes.push({
        kind: "split",
        location: locationOf(story, baseBlock),
        baseBlockId: baseBlock.id,
        targetBlockIds: splitInto.map(({ id }) => id),
        text: baseBlock.text,
      });
      operations.push({
        id: nextOperationId(),
        type: "splitBlock",
        blockId: baseBlock.id,
        offset,
        ...(separator.length > 0 && { separator }),
      });
      continue;
    }
    if (paragraphMarkPlan?.type === "merge") {
      const { baseBlocks, revisedBlock: targetBlock, separator } = paragraphMarkPlan;
      changes.push({
        kind: "merge",
        location: locationOf(story, baseBlocks[0]),
        baseBlockIds: baseBlocks.map(({ id }) => id),
        targetBlockId: targetBlock.id,
        text: targetBlock.text,
      });
      operations.push({
        id: nextOperationId(),
        type: "mergeBlockWithNext",
        blockId: baseBlocks[0].id,
        ...(separator.length > 0 && { separator }),
      });
      continue;
    }
    switch (step.type) {
      case "pair": {
        const { baseBlock, targetBlock } = step;
        const properties = changedFolioContentParagraphFormatting(baseBlock, targetBlock);
        if (properties) {
          changes.push({
            kind: "paragraph-format",
            location: locationOf(story, baseBlock),
            baseBlockId: baseBlock.id,
            targetBlockId: targetBlock.id,
            properties,
          });
          operations.push({
            id: nextOperationId(),
            type: "setBlockParagraphProperties",
            blockId: baseBlock.id,
            properties,
          });
        }
        if (baseBlock.text !== targetBlock.text) {
          changes.push({
            kind: "replace",
            location: locationOf(story, baseBlock),
            baseBlockId: baseBlock.id,
            targetBlockId: targetBlock.id,
            before: baseBlock.text,
            after: targetBlock.text,
          });
          operations.push({
            id: nextOperationId(),
            type: "replaceBlock",
            blockId: baseBlock.id,
            text: targetBlock.text,
          });
          break;
        }
        const segments = inlineFormattingSegments({
          baseBlock,
          targetBlock,
          maxSegments: maxOperations,
        });
        if (segments === null) {
          return null;
        }
        if (segments.length === 0) {
          break;
        }
        changes.push({
          kind: "format",
          location: locationOf(story, baseBlock),
          baseBlockId: baseBlock.id,
          targetBlockId: targetBlock.id,
          text: baseBlock.text,
          ranges: segments,
        });
        for (const { startOffset, endOffset, formatting } of segments) {
          const range = createFolioAITextRangeHandle({
            blockId: baseBlock.id,
            text: baseBlock.text,
            startOffset,
            endOffset,
          });
          if (!range) {
            panic("An aligned formatting range could not be represented");
          }
          operations.push({ id: nextOperationId(), type: "formatRange", range, formatting });
        }
        break;
      }
      case "baseOnly": {
        const targetBlockId = movesByBaseBlockId.get(step.block.id);
        if (targetBlockId === undefined) {
          changes.push({
            kind: "delete",
            location: locationOf(story, step.block),
            baseBlockId: step.block.id,
            before: step.block.text,
          });
        }
        operations.push({
          id: nextOperationId(),
          type: "deleteBlock",
          blockId: step.block.id,
          ...(targetBlockId !== undefined && { moveId: moveIdOf(step.block.id) }),
        });
        break;
      }
      case "targetOnly": {
        const baseBlockId = moveSourceByTargetBlockId.get(step.block.id);
        changes.push(
          baseBlockId === undefined
            ? {
                kind: "insert",
                location: locationOf(story, step.block),
                targetBlockId: step.block.id,
                after: step.block.text,
              }
            : {
                kind: "move",
                location: locationOf(story, step.block),
                baseBlockId,
                targetBlockId: step.block.id,
                text: step.block.text,
              },
        );
        pushInsertOperation(step.block, anchorIds[stepIndex] ?? null);
        break;
      }
      case "baseRow": {
        const anchorBlockId = step.blocks[0]?.id;
        if (anchorBlockId === undefined) {
          panic("A collapsed table row carried no blocks");
        }
        changes.push({
          kind: "table-row-delete",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rowIndex: step.location.rowIndex,
          cells: rowCellTexts(step.blocks),
          baseBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({ id: nextOperationId(), type: "deleteTableRow", blockId: anchorBlockId });
        break;
      }
      case "baseColumn": {
        const anchorBlockId = step.blocks.at(0)?.id;
        if (anchorBlockId === undefined) {
          panic("A table column carried no blocks");
        }
        changes.push({
          kind: "table-column-delete",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          columnIndex: step.columnIndex,
          cells: columnCellTexts(step.blocks),
          baseBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({
          id: nextOperationId(),
          type: "deleteTableColumn",
          blockId: anchorBlockId,
        });
        break;
      }
      case "baseTable": {
        const anchorBlockId = step.blocks[0]?.id;
        if (anchorBlockId === undefined) {
          panic("An unpaired table segment carried no blocks");
        }
        changes.push({
          kind: "table-delete",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rows: tableCellTexts(step.blocks),
          baseBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({ id: nextOperationId(), type: "deleteTable", blockId: anchorBlockId });
        break;
      }
      case "targetTable": {
        const rows = tableCellTexts(step.blocks);
        changes.push({
          kind: "table-insert",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rows,
          targetBlockIds: step.blocks.map(({ id }) => id),
        });
        const before = anchorIds[stepIndex] ?? null;
        const anchorBlockId = before ?? tailAnchorId;
        if (anchorBlockId === null) {
          break;
        }
        const operationId = nextOperationId();
        operations.push({
          id: operationId,
          type: "insertTable",
          blockId: anchorBlockId,
          position: before === null ? "after" : "before",
          rows,
        });
        // The rows above are the change list's summary; the table itself comes
        // from the target, so its grid, row and cell properties survive.
        tableTemplates.push({ operationId, targetTableIndex: step.location.tableIndex });
        break;
      }
      case "targetRow": {
        const anchor = findRowAnchor(steps, stepIndex);
        const cells = rowCellTexts(step.blocks);
        changes.push({
          kind: "table-row-insert",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          rowIndex: step.location.rowIndex,
          cells,
          targetBlockIds: step.blocks.map(({ id }) => id),
        });
        if (anchor === null) {
          // The base story has no table to grow, so the row's text can only
          // land as ordinary paragraphs.
          for (const block of step.blocks) {
            pushInsertOperation(block, anchorIds[stepIndex] ?? null);
          }
          break;
        }
        const rowOperationId = nextOperationId();
        operations.push({
          id: rowOperationId,
          type: "insertTableRow",
          blockId: anchor.blockId,
          position: anchor.position,
          cellTexts: cells,
        });
        tableTemplates.push({
          operationId: rowOperationId,
          targetTableIndex: step.location.tableIndex,
          targetRowIndex: step.location.rowIndex,
        });
        break;
      }
      case "targetColumn": {
        changes.push({
          kind: "table-column-insert",
          location: { story, cell: step.location },
          tableIndex: step.location.tableIndex,
          columnIndex: step.columnIndex,
          cells: columnCellTexts(step.blocks),
          targetBlockIds: step.blocks.map(({ id }) => id),
        });
        operations.push({
          id: nextOperationId(),
          type: "insertTableColumn",
          blockId: step.anchor.blockId,
          position: step.anchor.position,
          cellTexts: columnCellTexts(step.blocks),
        });
        break;
      }
      default: {
        const unreachable: never = step;
        panic("Unhandled compare step", { step: unreachable });
      }
    }
    if (operations.length > maxOperations) {
      return null;
    }
  }

  const planned = withTrailingDeletionRules({
    baseSnapshot,
    targetSnapshot,
    operations,
    tableTemplates,
    nextOperationId,
  });

  return planned.length > maxOperations
    ? null
    : {
        changes,
        operations: planned,
        tableTemplates,
        tableGeometryPairings: tableGeometryPairingsOf(steps),
      };
};
