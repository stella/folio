/**
 * Turn one completed neutral story comparison into the change list a caller
 * reads and the closed DOCX instruction program that realizes it.
 *
 * Pure: no parsing, no serialization, no clock. Alignment, moves, split/merge,
 * text segments, and formatting classification have already been completed by
 * the representation-neutral comparison core. This module only lowers that
 * immutable semantic product into tracked-document operations.
 */

import { panic, Result } from "better-result";

import type { FolioDocumentStoryHandle } from "../ai-edits/headless";
import { trailingBodyBlockId } from "../ai-edits/snapshot";
import type {
  FolioAIBlock,
  FolioAIBlockParagraphProperties,
  FolioAIBlockTableLocation,
  FolioAIEditSnapshot,
} from "../ai-edits/types";
import type {
  TableCellCoordinate,
  TableGeometryPairing,
} from "../internal/compare/table-geometry-program";
import { groupFolioContentTableRows as groupRows } from "./content-alignment";
import {
  docxAuthoredRunsForBlock,
  docxAuthoredRunsForRange,
  docxParagraphPropertiesFromBlock,
  docxParagraphPropertiesFromChanges,
  docxTableLocationFromContent,
} from "./docx-content-adapter";
import {
  DocxComparisonProgram,
  type DocxAuthoredChangeRange,
  type DocxComparisonBlockExpectation,
  type DocxComparisonInsertionAnchor,
  type DocxComparisonInstructionInput,
  type DocxComparisonParagraphTargetInput,
  type DocxComparisonRangePlanInput,
} from "../internal/compare/docx-program";
import {
  type FolioContentBlockChange,
  type FolioContentComparison,
  type FolioContentComparisonEvent,
  type FolioContentFormattingChange,
  type FolioContentPairRelation,
  type FolioContentRangePairRelation,
  type FolioContentStructuralChange,
  type FolioContentWholePairRelation,
} from "./content";
import type { FolioContentBlock, FolioContentTableLocation } from "./content-types";
import {
  CompareDocxLoweringError,
  CompareDocxOperationLimitError,
  type CompareChange,
  type CompareChangeLocation,
  type CompareUnsupportedPart,
} from "./types";

const isEmptyParagraphNode = (block: FolioAIBlock, snapshot: FolioAIEditSnapshot): boolean => {
  const anchor =
    snapshot.anchors[block.id] ??
    panic("A comparison snapshot block has no matching anchor", { blockId: block.id });
  return block.text === "" && block.table === undefined && anchor.to - anchor.from === 2;
};

const baseBlocksOfEvent = (event: FolioContentComparisonEvent): readonly FolioContentBlock[] => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return [event.relation.base.block];
    case "deleted":
      return [event.block];
    case "inserted":
    case "movedTo":
      return [];
    case "movedFrom":
      return [event.move.relation.base.block];
    case "split":
      return [event.relations[0].base.block];
    case "merge":
      return [event.relations[0].base.block, event.relations[1].base.block];
    case "tableReplacement":
      return event.replacement.baseBlocks;
    case "structural":
      return event.change.type === "table-delete" ||
        event.change.type === "table-row-delete" ||
        event.change.type === "table-column-delete"
        ? [event.change.blocks[event.memberIndex] ?? panic("Missing structural base member")]
        : [];
    default: {
      const unreachable: never = event;
      return panic("Unhandled comparison event while projecting base blocks", {
        event: unreachable,
      });
    }
  }
};

const revisedBlocksOfEvent = (event: FolioContentComparisonEvent): readonly FolioContentBlock[] => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return [event.relation.revised.block];
    case "inserted":
      return [event.block];
    case "deleted":
    case "movedFrom":
      return [];
    case "movedTo":
      return [event.move.relation.revised.block];
    case "split":
      return [event.relations[0].revised.block, event.relations[1].revised.block];
    case "merge":
      return [event.relations[0].revised.block];
    case "tableReplacement":
      return event.replacement.revisedBlocks;
    case "structural":
      return event.change.type === "table-insert" ||
        event.change.type === "table-row-insert" ||
        event.change.type === "table-column-insert"
        ? [event.change.blocks[event.memberIndex] ?? panic("Missing structural revised member")]
        : [];
    default: {
      const unreachable: never = event;
      return panic("Unhandled comparison event while projecting revised blocks", {
        event: unreachable,
      });
    }
  }
};

/** The next surviving base block at each canonical stream position. */
const nextBaseBlockIdByEvent = (
  events: readonly FolioContentComparisonEvent[],
): (string | null)[] => {
  const anchors = Array.from<string | null>({ length: events.length });
  let next: string | null = null;
  for (let index = events.length - 1; index >= 0; index--) {
    anchors[index] = next;
    const block = events[index] ? baseBlocksOfEvent(events[index]).at(0) : undefined;
    next = block?.identity.id ?? next;
  }
  return anchors;
};

/** A base block sitting inside a table, and whether it precedes or follows the step. */
type RowAnchor = { blockId: string; position: "after" | "before" };

const baseTableBlockOf = (event: FolioContentComparisonEvent): FolioContentBlock | null =>
  baseBlocksOfEvent(event).find(({ table }) => table !== undefined) ?? null;

/**
 * The base-document row a new row is inserted next to: the nearest base block
 * inside a table, preferring the one before the insertion so a run of new rows
 * keeps its order.
 */
const findRowAnchor = (
  events: readonly FolioContentComparisonEvent[],
  eventIndex: number,
): RowAnchor | null => {
  for (let index = eventIndex - 1; index >= 0; index--) {
    const event = events[index];
    const block = event ? baseTableBlockOf(event) : null;
    if (block) {
      return { blockId: block.identity.id, position: "after" };
    }
  }
  for (let index = eventIndex + 1; index < events.length; index++) {
    const event = events[index];
    const block = event ? baseTableBlockOf(event) : null;
    if (block) {
      return { blockId: block.identity.id, position: "before" };
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
const tableCellTexts = (blocks: readonly FolioContentBlock[]): string[][] => {
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

const rowCellTexts = (blocks: readonly FolioContentBlock[]): string[] => {
  const byCell: (string | undefined)[] = [];
  for (const block of blocks) {
    const cellIndex = block.table?.cellIndex ?? 0;
    const existing = byCell[cellIndex];
    byCell[cellIndex] = existing === undefined ? block.text : `${existing}\n${block.text}`;
  }
  return Array.from(byCell, (text) => text ?? "");
};

const columnCellTexts = (blocks: readonly FolioContentBlock[]): string[] => {
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

const locationOf = (
  story: FolioDocumentStoryHandle,
  block: FolioContentBlock,
): CompareChangeLocation =>
  block.table ? { story, cell: docxTableLocationFromContent(block.table) } : { story };

/**
 * The container a paragraph mark ends: the story's body, or one table cell. A
 * mark joins the paragraph it ends with the next paragraph of the SAME
 * container, so two blocks on either side of a boundary are not neighbours for
 * this purpose however adjacent they read.
 */
type ContainerLocatedBlock = {
  readonly table?: {
    readonly tableIndex: number;
    readonly rowIndex: number;
    readonly cellIndex: number;
  };
};

const containerKeyOf = ({ table }: ContainerLocatedBlock): string =>
  table
    ? `t${String(table.tableIndex)}r${String(table.rowIndex)}c${String(table.cellIndex)}`
    : "body";

/** Each container's last block, in the order the story holds them. */
const lastBlockByContainer = <Block extends ContainerLocatedBlock>(
  blocks: readonly Block[],
): ReadonlyMap<string, Block> => {
  const last = new Map<string, Block>();
  for (const block of blocks) {
    last.set(containerKeyOf(block), block);
  }
  return last;
};

type TrailingDeletionOptions = {
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  baseContentBlocks: readonly FolioContentBlock[];
  targetContentBlocks: readonly FolioContentBlock[];
  /** The plan so far, rewritten around each container it empties. */
  instructions: readonly DocxComparisonInstructionInput[];
};

type TrailingDeletionPlan = {
  readonly instructions: DocxComparisonInstructionInput[];
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
  baseContentBlocks,
  targetContentBlocks,
  instructions,
}: TrailingDeletionOptions): TrailingDeletionPlan => {
  const plan = [...instructions];
  const deletionIndexByBlockId = new Map<string, number>();
  const insertIndexesByAnchor = new Map<string, number[]>();
  const tableInsertIndexesByAnchor = new Map<string, number[]>();
  // A block whose own mark the plan already moves is not a chain start: a
  // split has put a second paragraph after it, and a merge has spent its mark.
  const markedBlockIds = new Set<string>();
  for (const [index, instruction] of plan.entries()) {
    switch (instruction.type) {
      case "deleteParagraph": {
        deletionIndexByBlockId.set(instruction.source.blockId, index);
        break;
      }
      case "insertParagraph": {
        const placed = insertIndexesByAnchor.get(instruction.anchor.blockId) ?? [];
        placed.push(index);
        insertIndexesByAnchor.set(instruction.anchor.blockId, placed);
        break;
      }
      case "insertTable": {
        const placed = tableInsertIndexesByAnchor.get(instruction.anchor.blockId) ?? [];
        placed.push(index);
        tableInsertIndexesByAnchor.set(instruction.anchor.blockId, placed);
        break;
      }
      case "splitParagraph": {
        markedBlockIds.add(instruction.source.blockId);
        break;
      }
      case "mergeParagraphs": {
        markedBlockIds.add(instruction.firstSource.blockId);
        break;
      }
      case "moveParagraph": {
        markedBlockIds.add(instruction.source.blockId);
        break;
      }
      default: {
        break;
      }
    }
  }
  if (deletionIndexByBlockId.size === 0) {
    return { instructions: plan };
  }

  const blocks = baseSnapshot.blocks;
  const baseContentById = new Map(
    baseContentBlocks.map((block) => [block.identity.id, block] as const),
  );
  const baseDocxById = targetBlockById(baseSnapshot);
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

  const targetLastByContainer = lastBlockByContainer(targetContentBlocks);
  const terminalTargetTableIndex = targetSnapshot.blocks.at(-1)?.table?.outerTableIndex;
  const dropped = new Set<number>();
  const appended: DocxComparisonInstructionInput[] = [];
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
        tableInsert.targetTableIndex !== terminalTargetTableIndex
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
        anchor: { blockId: carrier.id, position: "after" },
      };
      continue;
    }
    const lastInsertIndex = run.insertIndexes.at(-1);
    if (lastInsertIndex === undefined) {
      if (run.chainStart) {
        const chainStart = baseContentById.get(run.chainStart.id);
        if (!chainStart) {
          return panic("A terminal merge carrier lost its canonical base block", {
            blockId: run.chainStart.id,
          });
        }
        appended.push({
          type: "mergeTerminalCarrier",
          source: blockExpectation(chainStart),
        });
      }
      if (isEmptyParagraphNode(carrier, baseSnapshot)) {
        dropped.add(carrierDeletionIndex);
      }
    } else {
      const lastInsert = plan[lastInsertIndex];
      if (lastInsert?.type !== "insertParagraph") {
        // A relocation is one atomic move instruction, so it is not rewritten
        // as a terminal insertion without its source half.
        continue;
      }
      const carrierDeletion = plan[carrierDeletionIndex];
      if (carrierDeletion?.type !== "deleteParagraph") {
        continue;
      }
      dropped.add(lastInsertIndex);
      if (carrier.text === lastInsert.target.text) {
        // The carrier already holds the words the last inserted paragraph
        // brings, so neither half of the exchange is a revision.
        dropped.add(carrierDeletionIndex);
      } else {
        const baseContent = baseContentById.get(carrier.id);
        if (!baseContent) {
          return panic("A terminal carrier lost its canonical base block", {
            blockId: carrier.id,
          });
        }
        plan[carrierDeletionIndex] = {
          type: "replaceText",
          source: carrierDeletion.source,
          sourceStartOffset: 0,
          range: {
            sourceText: carrier.text,
            targetText: lastInsert.target.text,
            segments: Object.freeze([
              ...(carrier.text.length > 0
                ? [
                    Object.freeze({
                      type: "del" as const,
                      text: carrier.text,
                      baseStart: 0,
                      baseEnd: carrier.text.length,
                      revisedStart: 0,
                      revisedEnd: 0,
                    }),
                  ]
                : []),
              ...(lastInsert.target.text.length > 0
                ? [
                    Object.freeze({
                      type: "ins" as const,
                      text: lastInsert.target.text,
                      baseStart: carrier.text.length,
                      baseEnd: carrier.text.length,
                      revisedStart: 0,
                      revisedEnd: lastInsert.target.text.length,
                    }),
                  ]
                : []),
            ]),
            sourceRuns: docxAuthoredRunsForBlock(
              baseContent,
              targetDocxBlock(baseDocxById, baseContent),
            ),
            targetRuns: lastInsert.target.runs,
            authoredChanges: [],
          },
        };
      }
      for (const index of run.insertIndexes.slice(0, -1)) {
        const insert = plan[index];
        if (insert?.type !== "insertParagraph") {
          continue;
        }
        plan[index] = {
          type: "insertParagraph",
          anchor: { blockId: carrier.id, position: "before" },
          target: insert.target,
        };
      }
    }
    // The paragraph the carrier's mark now ends is the target's last one in
    // this container, so the carrier is where its properties have to be.
    const targetCarrier = targetLastByContainer.get(container);
    if (targetCarrier) {
      const baseCarrier = baseContentById.get(carrier.id);
      if (!baseCarrier) {
        return panic("A terminal property carrier lost its canonical base block", {
          blockId: carrier.id,
        });
      }
      appended.push({
        type: "setParagraphProperties",
        source: blockExpectation(baseCarrier),
        targetProperties: docxParagraphPropertiesFromBlock(targetCarrier),
      });
    }
  }
  return {
    instructions: [...plan.filter((_, index) => !dropped.has(index)), ...appended],
  };
};

export type CompareStoryPlan = {
  readonly changes: readonly CompareChange[];
  readonly unsupported: readonly CompareUnsupportedPart[];
  /** The sole transport semantics for this story, consumed once by apply. */
  readonly program: DocxComparisonProgram;
};

const cellCoordinate = ({
  tableIndex,
  rowIndex,
  cellIndex,
}: FolioAIBlockTableLocation): TableCellCoordinate => ({ tableIndex, rowIndex, cellIndex });

const pairedRelationsOfEvent = (
  event: FolioContentComparisonEvent,
): readonly FolioContentPairRelation[] => {
  switch (event.type) {
    case "unchanged":
    case "modified":
    case "formatting":
      return [event.relation];
    case "movedFrom":
    case "movedTo":
      return [event.move.relation];
    case "split":
    case "merge":
      return [...event.relations, event.separator];
    case "inserted":
    case "deleted":
    case "tableReplacement":
    case "structural":
      return [];
    default: {
      const unreachable: never = event;
      return panic("Unhandled comparison event while projecting table geometry", {
        event: unreachable,
      });
    }
  }
};

/** Cells paired by the canonical relation graph, once per base cell identity. */
const tableGeometryPairingsOf = (
  comparison: FolioContentComparison,
): TableGeometryPairing[] => {
  const pairings: TableGeometryPairing[] = [];
  const seen = new Set<string>();
  for (const event of comparison.events) {
    for (const relation of pairedRelationsOfEvent(event)) {
      const base = relation.base.block.table;
      const target = relation.revised.block.table;
      if (!base || !target || seen.has(base.cellIdentity.id)) continue;
      seen.add(base.cellIdentity.id);
      pairings.push({ base: cellCoordinate(base), target: cellCoordinate(target) });
    }
  }
  return pairings;
};

export type PlanStoryCompareOptions = {
  story: FolioDocumentStoryHandle;
  baseSnapshot: FolioAIEditSnapshot;
  targetSnapshot: FolioAIEditSnapshot;
  comparison: FolioContentComparison;
  /** Cap on generated operations, separate from semantic comparison budgets. */
  maxOperations: number;
};

type PlanStoryCompareError = CompareDocxLoweringError | CompareDocxOperationLimitError;

const operationLimit = (maxOperations: number): CompareDocxOperationLimitError =>
  new CompareDocxOperationLimitError({
    message: "The comparison needs more operations than the engine generates.",
    limit: maxOperations,
  });

const loweringError = ({
  story,
  reason,
  message,
  baseBlockId,
  targetBlockId,
  tableIndex,
}: {
  story: FolioDocumentStoryHandle;
  reason: ConstructorParameters<typeof CompareDocxLoweringError>[0]["reason"];
  message: string;
  baseBlockId?: string;
  targetBlockId?: string;
  tableIndex?: number;
}): CompareDocxLoweringError =>
  new CompareDocxLoweringError({
    message,
    reason,
    story,
    ...(baseBlockId !== undefined && { baseBlockId }),
    ...(targetBlockId !== undefined && { targetBlockId }),
    ...(tableIndex !== undefined && { tableIndex }),
  });

const targetBlockById = (snapshot: FolioAIEditSnapshot): ReadonlyMap<string, FolioAIBlock> =>
  new Map(snapshot.blocks.map((block) => [block.id, block]));

const targetDocxBlock = (
  blocks: ReadonlyMap<string, FolioAIBlock>,
  block: FolioContentBlock,
): FolioAIBlock =>
  blocks.get(block.identity.id) ??
  panic("A canonical target block has no DOCX projection", { blockId: block.identity.id });

const blockExpectation = (block: FolioContentBlock): DocxComparisonBlockExpectation =>
  Object.freeze({
    blockId: block.identity.id,
    kind: block.kind,
    text: block.text,
    paragraphProperties: docxParagraphPropertiesFromBlock(block),
    structuralBoundaries: block.structuralBoundaries,
    containerPath: block.containerPath.map(({ kind, identity }) => ({ kind, id: identity.id })),
    ...(block.table !== undefined && { table: docxTableLocationFromContent(block.table) }),
  });

const paragraphTarget = (
  block: FolioContentBlock,
  docxBlocks: ReadonlyMap<string, FolioAIBlock>,
): DocxComparisonParagraphTargetInput =>
  Object.freeze({
    text: block.text,
    runs: docxAuthoredRunsForBlock(block, targetDocxBlock(docxBlocks, block)),
    properties: docxParagraphPropertiesFromBlock(block),
    ...(block.table !== undefined && { table: docxTableLocationFromContent(block.table) }),
  });

type RangePlanInputOptions = {
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation;
  baseDocxBlocks: ReadonlyMap<string, FolioAIBlock>;
  targetDocxBlocks: ReadonlyMap<string, FolioAIBlock>;
};

const rangePlanInput = ({
  relation,
  baseDocxBlocks,
  targetDocxBlocks,
}: RangePlanInputOptions): DocxComparisonRangePlanInput => {
  const baseStart = relation.base.startOffset;
  const revisedStart = relation.revised.startOffset;
  const baseBlock = relation.base.block;
  const revisedBlock = relation.revised.block;
  const baseDocx = targetDocxBlock(baseDocxBlocks, baseBlock);
  const revisedDocx = targetDocxBlock(targetDocxBlocks, revisedBlock);
  return {
    sourceText: baseBlock.text.slice(baseStart, relation.base.endOffset),
    targetText: revisedBlock.text.slice(revisedStart, relation.revised.endOffset),
    segments: relation.segments.map((segment) =>
      Object.freeze({
        ...segment,
        baseStart: segment.baseStart - baseStart,
        baseEnd: segment.baseEnd - baseStart,
        revisedStart: segment.revisedStart - revisedStart,
        revisedEnd: segment.revisedEnd - revisedStart,
      }),
    ),
    sourceRuns: docxAuthoredRunsForRange(
      baseBlock,
      baseDocx,
      relation.base.startOffset,
      relation.base.endOffset,
    ),
    targetRuns: docxAuthoredRunsForRange(
      revisedBlock,
      revisedDocx,
      relation.revised.startOffset,
      relation.revised.endOffset,
    ),
    authoredChanges:
      relation.formatting?.ranges
        .filter(({ formatting }) => formatting.authored.length > 0)
        .map(
          ({ baseStart: rangeBaseStart, baseEnd, revisedStart: rangeRevisedStart, revisedEnd, formatting }) =>
            Object.freeze({
              baseStart: rangeBaseStart - baseStart,
              baseEnd: baseEnd - baseStart,
              revisedStart: rangeRevisedStart - revisedStart,
              revisedEnd: revisedEnd - revisedStart,
              properties: Object.freeze(formatting.authored.map(({ key }) => key)),
            }) satisfies DocxAuthoredChangeRange,
        ) ?? [],
  };
};

type FormattingRangePlanInputOptions = RangePlanInputOptions & {
  range: NonNullable<FolioContentWholePairRelation["formatting"]>["ranges"][number];
};

const formattingRangePlanInput = ({
  relation,
  baseDocxBlocks,
  targetDocxBlocks,
  range,
}: FormattingRangePlanInputOptions): DocxComparisonRangePlanInput => {
  const sourceText = relation.base.block.text.slice(range.baseStart, range.baseEnd);
  const targetText = relation.revised.block.text.slice(range.revisedStart, range.revisedEnd);
  return {
    sourceText,
    targetText,
    segments: Object.freeze([
      Object.freeze({
        type: "equal",
        text: sourceText,
        baseStart: 0,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: targetText.length,
      }),
    ]),
    sourceRuns: docxAuthoredRunsForRange(
      relation.base.block,
      targetDocxBlock(baseDocxBlocks, relation.base.block),
      range.baseStart,
      range.baseEnd,
    ),
    targetRuns: docxAuthoredRunsForRange(
      relation.revised.block,
      targetDocxBlock(targetDocxBlocks, relation.revised.block),
      range.revisedStart,
      range.revisedEnd,
    ),
    authoredChanges: Object.freeze([
      Object.freeze({
        baseStart: 0,
        baseEnd: sourceText.length,
        revisedStart: 0,
        revisedEnd: targetText.length,
        properties: Object.freeze(range.formatting.authored.map(({ key }) => key)),
      }),
    ]),
  };
};

const formattingChangeOf = (
  relation: FolioContentWholePairRelation,
): FolioContentFormattingChange | null => relation.formatting;

const textChanged = (relation: FolioContentPairRelation): boolean =>
  relation.segments.some(({ type }) => type !== "equal");

type UnsupportedRelationField = {
  readonly reason:
    | "block-semantics"
    | "container-change"
    | "effective-inline-formatting"
    | "effective-paragraph-formatting"
    | "structural-boundary-change";
  readonly field: string;
};

const unsupportedRelationFields = (
  relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
): readonly UnsupportedRelationField[] => {
  const unsupported: UnsupportedRelationField[] = [];
  for (const change of relation.blockChanges) {
    switch (change.field) {
      case "kind":
      case "blockProperties":
      case "table":
        unsupported.push({ reason: "block-semantics", field: change.field });
        break;
      case "containerPath":
        unsupported.push({ reason: "container-change", field: change.field });
        break;
      case "structuralBoundaries":
        unsupported.push({ reason: "structural-boundary-change", field: change.field });
        break;
      default: {
        const unreachable: never = change;
        return panic("Unhandled canonical block change", { change: unreachable });
      }
    }
  }
  const formatting = relation.formatting;
  if (!formatting) return Object.freeze(unsupported);
  const authoredParagraphKeys = new Set(
    formatting.paragraph.authored.map(({ key }) => key),
  );
  for (const change of formatting.paragraph.effective) {
    if (!authoredParagraphKeys.has(change.key)) {
      unsupported.push({
        reason: "effective-paragraph-formatting",
        field: `paragraph.${change.key}`,
      });
    }
  }
  for (const range of formatting.ranges) {
    const authoredKeys = new Set(range.formatting.authored.map(({ key }) => key));
    for (const change of range.formatting.effective) {
      if (!authoredKeys.has(change.key)) {
        unsupported.push({
          reason: "effective-inline-formatting",
          field: `runs.${change.key}`,
        });
      }
    }
  }
  return Object.freeze(unsupported);
};

/** Exhaustively lower one completed semantic story comparison. */
export const planStoryCompare = ({
  story,
  baseSnapshot,
  targetSnapshot,
  comparison,
  maxOperations,
}: PlanStoryCompareOptions): Result<CompareStoryPlan, PlanStoryCompareError> => {
  const events = comparison.events;
  const anchorIds = nextBaseBlockIdByEvent(events);
  const changes: CompareChange[] = [];
  const unsupported: CompareUnsupportedPart[] = [];
  const instructions: DocxComparisonInstructionInput[] = [];
  const baseBlocks = targetBlockById(baseSnapshot);
  const targetBlocks = targetBlockById(targetSnapshot);

  const recordUnsupportedRelation = (
    eventType: Extract<CompareUnsupportedPart, { story: FolioDocumentStoryHandle }>["eventType"],
    relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  ): boolean => {
    const fields = unsupportedRelationFields(relation);
    for (const field of fields) {
      unsupported.push({
        reason: field.reason,
        story,
        eventType,
        field: field.field,
        baseBlockId: relation.base.block.identity.id,
        targetBlockId: relation.revised.block.identity.id,
      });
    }
    return fields.length > 0;
  };
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

  const insertionAnchor = (
    anchorId: string | null,
    block: FolioContentBlock,
  ): DocxComparisonInsertionAnchor | CompareDocxLoweringError => {
    if (anchorId !== null) {
      return Object.freeze({ blockId: anchorId, position: "before" });
    }
    if (tailAnchorId === null) {
      return loweringError({
        story,
        reason: "missing-insertion-anchor",
        message: "The target adds content to a story with no tracked insertion anchor.",
        targetBlockId: block.identity.id,
      });
    }
    return Object.freeze({ blockId: tailAnchorId, position: "after" });
  };

  const pushInsertInstruction = (
    block: FolioContentBlock,
    anchorId: string | null,
  ): CompareDocxLoweringError | null => {
    const anchor = insertionAnchor(anchorId, block);
    if (anchor instanceof CompareDocxLoweringError) return anchor;
    instructions.push({
      type: "insertParagraph",
      anchor,
      target: paragraphTarget(block, targetBlocks),
    });
    return null;
  };

  const pushParagraphFormattingReport = (
    relation: FolioContentWholePairRelation | FolioContentRangePairRelation,
  ): void => {
    const formatting = relation.formatting;
    if (!formatting || formatting.paragraph.authored.length === 0) return;
    const properties = docxParagraphPropertiesFromChanges(formatting.paragraph.authored);
    changes.push({
      kind: "paragraph-format",
      location: locationOf(story, relation.base.block),
      baseBlockId: relation.base.block.identity.id,
      targetBlockId: relation.revised.block.identity.id,
      properties,
    });
  };

  const pushFormatting = (relation: FolioContentWholePairRelation): void => {
    const formatting = formattingChangeOf(relation);
    if (!formatting) return;
    if (formatting.paragraph.authored.length > 0) {
      pushParagraphFormattingReport(relation);
      instructions.push({
        type: "setParagraphProperties",
        source: blockExpectation(relation.base.block),
        targetProperties: docxParagraphPropertiesFromBlock(relation.revised.block),
      });
    }
    if (formatting.ranges.length === 0) return;
    changes.push({
      kind: "format",
      location: locationOf(story, relation.base.block),
      baseBlockId: relation.base.block.identity.id,
      targetBlockId: relation.revised.block.identity.id,
      text: relation.base.block.text,
      ranges: formatting.ranges.map(({ baseStart, baseEnd, formatting: delta }) => ({
        startOffset: baseStart,
        endOffset: baseEnd,
        formatting: delta,
      })),
    });
    for (const range of formatting.ranges) {
      if (range.formatting.authored.length === 0) continue;
      instructions.push({
        type: "formatText",
        source: blockExpectation(relation.base.block),
        sourceStartOffset: range.baseStart,
        range: formattingRangePlanInput({
          relation,
          baseDocxBlocks: baseBlocks,
          targetDocxBlocks: targetBlocks,
          range,
        }),
      });
    }
  };

  for (const [eventIndex, event] of events.entries()) {
    switch (event.type) {
      case "unchanged":
        break;
      case "modified":
      case "formatting": {
        const { relation } = event;
        recordUnsupportedRelation(event.type, relation);
        if (textChanged(relation)) {
          changes.push({
            kind: "replace",
            location: locationOf(story, relation.base.block),
            baseBlockId: relation.base.block.identity.id,
            targetBlockId: relation.revised.block.identity.id,
            before: relation.base.block.text,
            after: relation.revised.block.text,
          });
          instructions.push({
            type: "replaceText",
            source: blockExpectation(relation.base.block),
            sourceStartOffset: 0,
            range: rangePlanInput({
              relation,
              baseDocxBlocks: baseBlocks,
              targetDocxBlocks: targetBlocks,
            }),
          });
          if (
            (relation.formatting?.paragraph.authored.length ?? 0) > 0 ||
            (relation.formatting?.paragraph.effective.length ?? 0) > 0
          ) {
            const paragraphOnly = Object.freeze({
              ...relation,
              formatting: Object.freeze({
                paragraph: relation.formatting?.paragraph ??
                  Object.freeze({ authored: [], effective: [] }),
                ranges: [],
              }),
            });
            pushFormatting(paragraphOnly);
          }
        } else {
          pushFormatting(relation);
        }
        break;
      }
      case "deleted":
        changes.push({
          kind: "delete",
          location: locationOf(story, event.block),
          baseBlockId: event.block.identity.id,
          before: event.block.text,
        });
        instructions.push({
          type: "deleteParagraph",
          source: blockExpectation(event.block),
        });
        break;
      case "inserted": {
        changes.push({
          kind: "insert",
          location: locationOf(story, event.block),
          targetBlockId: event.block.identity.id,
          after: event.block.text,
        });
        if (event.block.structuralBoundaries.length > 0) {
          unsupported.push({
            reason: "structural-boundary-change",
            story,
            eventType: "inserted",
            field: "structuralBoundaries",
            targetBlockId: event.block.identity.id,
          });
        }
        if (event.block.containerPath.length > 0) {
          unsupported.push({
            reason: "container-change",
            story,
            eventType: "inserted",
            field: "containerPath",
            targetBlockId: event.block.identity.id,
          });
        }
        const error = pushInsertInstruction(event.block, anchorIds[eventIndex] ?? null);
        if (error) return Result.err(error);
        break;
      }
      case "movedFrom":
        break;
      case "movedTo": {
        const relation = event.move.relation;
        recordUnsupportedRelation("moved", relation);
        changes.push({
          kind: "move",
          location: locationOf(story, relation.revised.block),
          baseBlockId: relation.base.block.identity.id,
          targetBlockId: relation.revised.block.identity.id,
          text: relation.revised.block.text,
        });
        const anchor = insertionAnchor(anchorIds[eventIndex] ?? null, relation.revised.block);
        if (anchor instanceof CompareDocxLoweringError) return Result.err(anchor);
        instructions.push({
          type: "moveParagraph",
          source: blockExpectation(relation.base.block),
          anchor,
          target: paragraphTarget(relation.revised.block, targetBlocks),
        });
        break;
      }
      case "split": {
        const [first, second] = event.relations;
        recordUnsupportedRelation("split", first);
        recordUnsupportedRelation("split", second);
        const baseBlock = first.base.block;
        const firstBlock = first.revised.block;
        const secondBlock = second.revised.block;
        changes.push({
          kind: "split",
          location: locationOf(story, baseBlock),
          baseBlockId: baseBlock.identity.id,
          targetBlockIds: [firstBlock.identity.id, secondBlock.identity.id],
          text: baseBlock.text,
        });
        const separatorText = event.separator.base.block.text.slice(
          event.separator.base.startOffset,
          event.separator.base.endOffset,
        );
        instructions.push({
          type: "splitParagraph",
          source: blockExpectation(baseBlock),
          offset: first.base.endOffset - first.base.startOffset,
          first: rangePlanInput({
            relation: first,
            baseDocxBlocks: baseBlocks,
            targetDocxBlocks: targetBlocks,
          }),
          second: rangePlanInput({
            relation: second,
            baseDocxBlocks: baseBlocks,
            targetDocxBlocks: targetBlocks,
          }),
          separatorText,
          separatorRuns: docxAuthoredRunsForRange(
            event.separator.base.block,
            targetDocxBlock(baseBlocks, event.separator.base.block),
            event.separator.base.startOffset,
            event.separator.base.endOffset,
          ),
          firstTarget: paragraphTarget(firstBlock, targetBlocks),
          secondTarget: paragraphTarget(secondBlock, targetBlocks),
        });
        pushParagraphFormattingReport(first);
        pushParagraphFormattingReport(second);
        break;
      }
      case "merge": {
        const [first, second] = event.relations;
        recordUnsupportedRelation("merge", first);
        recordUnsupportedRelation("merge", second);
        const firstBlock = first.base.block;
        const secondBlock = second.base.block;
        const targetBlock = first.revised.block;
        changes.push({
          kind: "merge",
          location: locationOf(story, firstBlock),
          baseBlockIds: [firstBlock.identity.id, secondBlock.identity.id],
          targetBlockId: targetBlock.identity.id,
          text: targetBlock.text,
        });
        const separatorText = event.separator.revised.block.text.slice(
          event.separator.revised.startOffset,
          event.separator.revised.endOffset,
        );
        instructions.push({
          type: "mergeParagraphs",
          firstSource: blockExpectation(firstBlock),
          secondSource: blockExpectation(secondBlock),
          first: rangePlanInput({
            relation: first,
            baseDocxBlocks: baseBlocks,
            targetDocxBlocks: targetBlocks,
          }),
          second: rangePlanInput({
            relation: second,
            baseDocxBlocks: baseBlocks,
            targetDocxBlocks: targetBlocks,
          }),
          separatorText,
          separatorRuns: docxAuthoredRunsForRange(
            event.separator.revised.block,
            targetDocxBlock(targetBlocks, event.separator.revised.block),
            event.separator.revised.startOffset,
            event.separator.revised.endOffset,
          ),
          target: paragraphTarget(targetBlock, targetBlocks),
        });
        pushParagraphFormattingReport(first);
        break;
      }
      case "tableReplacement": {
        const { replacement } = event;
        const baseBlock = replacement.baseBlocks[0];
        const targetBlock = replacement.revisedBlocks[0];
        const baseTable = baseBlock.table;
        const targetTable = targetBlock.table;
        if (!baseTable || !targetTable) {
          return panic("A canonical table replacement has no table location");
        }
        changes.push({
          kind: "table-delete",
          location: { story, cell: docxTableLocationFromContent(baseTable) },
          tableIndex: replacement.baseTableIndex,
          rows: tableCellTexts(replacement.baseBlocks),
          baseBlockIds: replacement.baseBlocks.map(({ identity }) => identity.id),
        });
        changes.push({
          kind: "table-insert",
          location: { story, cell: docxTableLocationFromContent(targetTable) },
          tableIndex: replacement.revisedTableIndex,
          rows: tableCellTexts(replacement.revisedBlocks),
          targetBlockIds: replacement.revisedBlocks.map(({ identity }) => identity.id),
        });
        const before = anchorIds[eventIndex] ?? null;
        const anchorBlockId = before ?? tailAnchorId;
        if (anchorBlockId === null) {
          return Result.err(
            loweringError({
              story,
              reason: "missing-insertion-anchor",
              message: "The target table replacement has no tracked insertion anchor.",
              tableIndex: replacement.revisedTableIndex,
            }),
          );
        }
        instructions.push({
          type: "replaceTable",
          source: blockExpectation(baseBlock),
          baseTableIndex: replacement.baseTableIndex,
          anchor: {
            blockId: anchorBlockId,
            position: before === null ? "after" : "before",
          },
          targetTableIndex: replacement.revisedTableIndex,
        });
        break;
      }
      case "structural": {
        if (event.memberIndex !== 0) break;
        const structural = event.change;
        const firstBlock = structural.blocks[0];
        const location = firstBlock.table;
        if (!location) return panic("A canonical table change has no table location");
        const cell = docxTableLocationFromContent(location);
        switch (structural.type) {
          case "table-row-delete":
            changes.push({
              kind: "table-row-delete",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              rowIndex: structural.rowIndex,
              cells: rowCellTexts(structural.blocks),
              baseBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            instructions.push({
              type: "deleteTableRow",
              source: blockExpectation(firstBlock),
              baseTableIndex: structural.tableIndex,
              baseRowIndex: structural.rowIndex,
            });
            break;
          case "table-column-delete":
            changes.push({
              kind: "table-column-delete",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              columnIndex: structural.columnIndex,
              cells: columnCellTexts(structural.blocks),
              baseBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            instructions.push({
              type: "deleteTableColumn",
              source: blockExpectation(firstBlock),
              baseTableIndex: structural.tableIndex,
              baseColumnIndex: structural.columnIndex,
            });
            break;
          case "table-delete":
            changes.push({
              kind: "table-delete",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              rows: tableCellTexts(structural.blocks),
              baseBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            instructions.push({
              type: "deleteTable",
              source: blockExpectation(firstBlock),
              baseTableIndex: structural.tableIndex,
            });
            break;
          case "table-insert": {
            const rows = tableCellTexts(structural.blocks);
            changes.push({
              kind: "table-insert",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              rows,
              targetBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            const before = anchorIds[eventIndex] ?? null;
            const anchorBlockId = before ?? tailAnchorId;
            if (anchorBlockId === null) {
              return Result.err(
                loweringError({
                  story,
                  reason: "missing-insertion-anchor",
                  message: "The target adds a table to a story with no tracked insertion anchor.",
                  tableIndex: structural.tableIndex,
                }),
              );
            }
            instructions.push({
              type: "insertTable",
              anchor: {
                blockId: anchorBlockId,
                position: before === null ? "after" : "before",
              },
              targetTableIndex: structural.tableIndex,
            });
            break;
          }
          case "table-row-insert": {
            const anchor = findRowAnchor(events, eventIndex);
            if (!anchor) {
              return Result.err(
                loweringError({
                  story,
                  reason: "table-row-anchor",
                  message: "The inserted table row has no surviving table anchor.",
                  tableIndex: structural.tableIndex,
                }),
              );
            }
            const cells = rowCellTexts(structural.blocks);
            changes.push({
              kind: "table-row-insert",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              rowIndex: structural.rowIndex,
              cells,
              targetBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            instructions.push({
              type: "insertTableRow",
              anchor,
              targetTableIndex: structural.tableIndex,
              targetRowIndex: structural.rowIndex,
            });
            break;
          }
          case "table-column-insert":
            unsupported.push({
              reason: "table-column-content",
              story,
              eventType: "structural",
              tableIndex: structural.tableIndex,
            });
            changes.push({
              kind: "table-column-insert",
              location: { story, cell },
              tableIndex: structural.tableIndex,
              columnIndex: structural.columnIndex,
              cells: columnCellTexts(structural.blocks),
              targetBlockIds: structural.blocks.map(({ identity }) => identity.id),
            });
            instructions.push({
              type: "insertTableColumn",
              anchor: structural.anchor,
              targetTableIndex: structural.tableIndex,
              targetColumnIndex: structural.columnIndex,
              cellTexts: columnCellTexts(structural.blocks),
            });
            break;
          default: {
            const unreachable: never = structural;
            return panic("Unhandled canonical structural change", { change: unreachable });
          }
        }
        break;
      }
      default: {
        const unreachable: never = event;
        return panic("Unhandled canonical comparison event", { event: unreachable });
      }
    }
    if (instructions.length > maxOperations) {
      return Result.err(operationLimit(maxOperations));
    }
  }

  const tableGeometryPairings = tableGeometryPairingsOf(comparison);
  if (tableGeometryPairings.length > 0) {
    instructions.push({ type: "matchTableGeometry", pairings: tableGeometryPairings });
  }

  const baseContentBlocks = events.flatMap(baseBlocksOfEvent);
  const targetContentBlocks = events.flatMap(revisedBlocksOfEvent);

  const planned = withTrailingDeletionRules({
    baseSnapshot,
    targetSnapshot,
    baseContentBlocks,
    targetContentBlocks,
    instructions,
  });

  if (planned.instructions.length > maxOperations) {
    return Result.err(operationLimit(maxOperations));
  }
  return Result.ok({
    changes: Object.freeze(changes),
    unsupported: Object.freeze(unsupported),
    program: DocxComparisonProgram.create(planned.instructions),
  });
};
