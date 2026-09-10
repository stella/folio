/**
 * Pure alignment for representation-neutral content snapshots.
 *
 * The block walk is split by structural container before any textual
 * alignment runs. Body blocks are compared with body blocks; tables are
 * aligned table by table, then row by row and cell by cell. All quadratic
 * text alignment in one caller-owned work session shares one hard budget.
 */

import { panic } from "better-result";

import type {
  FolioContentBlock,
  FolioContentIdStability,
  FolioContentTableLocation,
} from "./content-types";
import { alignTableColumns, type TableColumnAlignmentStep } from "./column-alignment";

const MAX_CONTENT_ALIGNMENT_LCS_CELLS = 4_000_000;

export type FolioContentAlignmentWorkSession = {
  remainingLcsCells: number;
};

export type CreateFolioContentAlignmentWorkSessionOptions = {
  /** @internal A lower allowance used by focused boundary tests. */
  lcsCells?: number;
};

export const createFolioContentAlignmentWorkSession = (
  options: CreateFolioContentAlignmentWorkSessionOptions = {},
): FolioContentAlignmentWorkSession => {
  const lcsCells = options.lcsCells ?? MAX_CONTENT_ALIGNMENT_LCS_CELLS;
  if (
    !Number.isSafeInteger(lcsCells) ||
    lcsCells < 0 ||
    lcsCells > MAX_CONTENT_ALIGNMENT_LCS_CELLS
  ) {
    return panic("A content-alignment LCS allowance is outside the supported range", {
      lcsCells,
    });
  }
  return { remainingLcsCells: lcsCells };
};

export const exceedsFolioContentLcsBudget = (
  baseCount: number,
  revisedCount: number,
): boolean =>
  baseCount > 0 &&
  revisedCount > 0 &&
  baseCount > Math.floor(MAX_CONTENT_ALIGNMENT_LCS_CELLS / revisedCount);

export type FolioContentBlockPair = { baseIndex: number; revisedIndex: number };
type IndexedBlock<Block extends FolioContentBlock> = { block: Block; index: number };

/**
 * Longest strictly increasing subsequence by revised index.
 *
 * The Fenwick tree stores the earliest end index for equal-length candidates,
 * preserving the historical deterministic tie break while reducing the
 * stable-anchor pass from quadratic to O(n log n).
 */
export const longestIncreasingFolioContentPairs = (
  pairs: readonly FolioContentBlockPair[],
): FolioContentBlockPair[] => {
  if (pairs.length === 0) {
    return [];
  }
  const coordinates = [...new Set(pairs.map(({ revisedIndex }) => revisedIndex))].toSorted(
    (left, right) => left - right,
  );
  const rankByCoordinate = new Map(
    coordinates.map((coordinate, index) => [coordinate, index + 1] as const),
  );
  const treeLengths = new Int32Array(coordinates.length + 1);
  const treeIndexes = new Int32Array(coordinates.length + 1).fill(-1);
  const lengths = new Int32Array(pairs.length).fill(1);
  const predecessors = new Int32Array(pairs.length).fill(-1);

  const earlier = (
    leftLength: number,
    leftIndex: number,
    rightLength: number,
    rightIndex: number,
  ): boolean =>
    leftLength > rightLength ||
    (leftLength === rightLength && leftLength > 0 && leftIndex < rightIndex);

  let bestEnd = 0;
  for (let pairIndex = 0; pairIndex < pairs.length; pairIndex++) {
    const pair = pairs[pairIndex];
    const rank = pair === undefined ? undefined : rankByCoordinate.get(pair.revisedIndex);
    if (rank === undefined) {
      continue;
    }

    let predecessorLength = 0;
    let predecessorIndex = -1;
    for (let cursor = rank - 1; cursor > 0; cursor -= cursor & -cursor) {
      const candidateLength = treeLengths[cursor] ?? 0;
      const candidateIndex = treeIndexes[cursor] ?? -1;
      if (
        earlier(candidateLength, candidateIndex, predecessorLength, predecessorIndex)
      ) {
        predecessorLength = candidateLength;
        predecessorIndex = candidateIndex;
      }
    }
    lengths[pairIndex] = predecessorLength + 1;
    predecessors[pairIndex] = predecessorIndex;

    for (let cursor = rank; cursor < treeLengths.length; cursor += cursor & -cursor) {
      const currentLength = treeLengths[cursor] ?? 0;
      const currentIndex = treeIndexes[cursor] ?? -1;
      if (earlier(lengths[pairIndex] ?? 0, pairIndex, currentLength, currentIndex)) {
        treeLengths[cursor] = lengths[pairIndex] ?? 0;
        treeIndexes[cursor] = pairIndex;
      }
    }
    if ((lengths[pairIndex] ?? 0) > (lengths[bestEnd] ?? 0)) {
      bestEnd = pairIndex;
    }
  }

  const ordered: FolioContentBlockPair[] = [];
  for (let cursor = bestEnd; cursor !== -1; cursor = predecessors[cursor] ?? -1) {
    const pair = pairs[cursor];
    if (pair) {
      ordered.push(pair);
    }
  }
  return ordered.toReversed();
};

const pairByStableId = <Block extends FolioContentBlock>(
  base: readonly Block[],
  revised: readonly Block[],
  idStability: (block: Block) => FolioContentIdStability,
): FolioContentBlockPair[] => {
  const revisedIndexById = new Map<string, number>();
  revised.forEach((block, revisedIndex) => {
    if (idStability(block) === "stable") {
      revisedIndexById.set(block.id, revisedIndex);
    }
  });

  const candidates: FolioContentBlockPair[] = [];
  base.forEach((block, baseIndex) => {
    if (idStability(block) !== "stable") {
      return;
    }
    const revisedIndex = revisedIndexById.get(block.id);
    if (revisedIndex !== undefined) {
      candidates.push({ baseIndex, revisedIndex });
    }
  });
  return longestIncreasingFolioContentPairs(candidates);
};

const pairByExactText = <Block extends FolioContentBlock>(
  base: readonly IndexedBlock<Block>[],
  revised: readonly IndexedBlock<Block>[],
  workSession: FolioContentAlignmentWorkSession,
  idStability: (block: Block) => FolioContentIdStability,
): FolioContentBlockPair[] => {
  const baseCount = base.length;
  const revisedCount = revised.length;
  if (baseCount === 0 || revisedCount === 0) {
    return [];
  }
  if (
    exceedsFolioContentLcsBudget(baseCount, revisedCount) ||
    baseCount > Math.floor(workSession.remainingLcsCells / revisedCount)
  ) {
    return [];
  }
  workSession.remainingLcsCells -= baseCount * revisedCount;

  const baseTexts = base.map(({ block }) => block.text);
  const revisedTexts = revised.map(({ block }) => block.text);
  const entriesCanPair = (baseIndex: number, revisedIndex: number): boolean => {
    const baseBlock = base[baseIndex]?.block;
    const revisedBlock = revised[revisedIndex]?.block;
    if (!baseBlock || !revisedBlock || baseBlock.text !== revisedBlock.text) {
      return false;
    }
    return !(
      idStability(baseBlock) === "stable" &&
      idStability(revisedBlock) === "stable" &&
      baseBlock.id !== revisedBlock.id
    );
  };
  const stride = revisedCount + 1;
  const lengths = new Int32Array((baseCount + 1) * stride);
  for (let baseIndex = baseCount - 1; baseIndex >= 0; baseIndex--) {
    const rowOffset = baseIndex * stride;
    const nextRowOffset = (baseIndex + 1) * stride;
    for (let revisedIndex = revisedCount - 1; revisedIndex >= 0; revisedIndex--) {
      lengths[rowOffset + revisedIndex] =
        entriesCanPair(baseIndex, revisedIndex)
          ? (lengths[nextRowOffset + revisedIndex + 1] ?? 0) + 1
          : Math.max(
              lengths[nextRowOffset + revisedIndex] ?? 0,
              lengths[rowOffset + revisedIndex + 1] ?? 0,
            );
    }
  }

  const pairs: FolioContentBlockPair[] = [];
  let baseIndex = 0;
  let revisedIndex = 0;
  while (baseIndex < baseCount && revisedIndex < revisedCount) {
    const baseEntry = base[baseIndex];
    const revisedEntry = revised[revisedIndex];
    if (!baseEntry || !revisedEntry) {
      break;
    }
    if (entriesCanPair(baseIndex, revisedIndex)) {
      pairs.push({ baseIndex: baseEntry.index, revisedIndex: revisedEntry.index });
      baseIndex += 1;
      revisedIndex += 1;
      continue;
    }
    const down = lengths[(baseIndex + 1) * stride + revisedIndex] ?? 0;
    const right = lengths[baseIndex * stride + revisedIndex + 1] ?? 0;
    if (down >= right) {
      baseIndex += 1;
    } else {
      revisedIndex += 1;
    }
  }
  return pairs;
};

export type FolioContentAlignedBlockEvent<
  Block extends FolioContentBlock = FolioContentBlock,
> =
  | { type: "pair"; baseBlock: Block; revisedBlock: Block }
  | { type: "baseOnly"; block: Block }
  | { type: "revisedOnly"; block: Block };

export type AlignFolioContentBlocksOptions<Block extends FolioContentBlock> = {
  workSession?: FolioContentAlignmentWorkSession;
  /** @internal Compatibility hook for adapters whose source model predates `idStability`. */
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

export const alignFolioContentBlocks = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
  options: AlignFolioContentBlocksOptions<Block> = {},
): FolioContentAlignedBlockEvent<Block>[] => {
  const workSession = options.workSession ?? createFolioContentAlignmentWorkSession();
  const idStability =
    options.idStability ?? ((block: Block): FolioContentIdStability => block.idStability ?? "stable");
  const stableIdAnchors = pairByStableId(baseBlocks, revisedBlocks, idStability);
  const usedBaseIndexes = new Set(stableIdAnchors.map(({ baseIndex }) => baseIndex));
  const usedRevisedIndexes = new Set(stableIdAnchors.map(({ revisedIndex }) => revisedIndex));
  const baseRemaining = baseBlocks.flatMap((block, index) =>
    usedBaseIndexes.has(index) ? [] : [{ block, index }],
  );
  const revisedRemaining = revisedBlocks.flatMap((block, index) =>
    usedRevisedIndexes.has(index) ? [] : [{ block, index }],
  );
  const exactTextAnchors = pairByExactText(
    baseRemaining,
    revisedRemaining,
    workSession,
    idStability,
  );
  const anchors = longestIncreasingFolioContentPairs(
    [...stableIdAnchors, ...exactTextAnchors].toSorted(
      (left, right) => left.baseIndex - right.baseIndex,
    ),
  );
  const events: FolioContentAlignedBlockEvent<Block>[] = [];

  const emitGap = (
    baseFrom: number,
    baseTo: number,
    revisedFrom: number,
    revisedTo: number,
  ): void => {
    const pairedCount = Math.min(baseTo - baseFrom, revisedTo - revisedFrom);
    for (let offset = 0; offset < pairedCount; offset++) {
      const baseBlock = baseBlocks[baseFrom + offset];
      const revisedBlock = revisedBlocks[revisedFrom + offset];
      if (baseBlock && revisedBlock) {
        const differentStableIdentities =
          idStability(baseBlock) === "stable" &&
          idStability(revisedBlock) === "stable" &&
          baseBlock.id !== revisedBlock.id;
        if (differentStableIdentities) {
          events.push({ type: "baseOnly", block: baseBlock });
          events.push({ type: "revisedOnly", block: revisedBlock });
        } else {
          events.push({ type: "pair", baseBlock, revisedBlock });
        }
      }
    }
    for (let index = baseFrom + pairedCount; index < baseTo; index++) {
      const block = baseBlocks[index];
      if (block) {
        events.push({ type: "baseOnly", block });
      }
    }
    for (let index = revisedFrom + pairedCount; index < revisedTo; index++) {
      const block = revisedBlocks[index];
      if (block) {
        events.push({ type: "revisedOnly", block });
      }
    }
  };

  let baseCursor = 0;
  let revisedCursor = 0;
  for (const anchor of anchors) {
    emitGap(baseCursor, anchor.baseIndex, revisedCursor, anchor.revisedIndex);
    const baseBlock = baseBlocks[anchor.baseIndex];
    const revisedBlock = revisedBlocks[anchor.revisedIndex];
    if (baseBlock && revisedBlock) {
      events.push({ type: "pair", baseBlock, revisedBlock });
    }
    baseCursor = anchor.baseIndex + 1;
    revisedCursor = anchor.revisedIndex + 1;
  }
  emitGap(baseCursor, baseBlocks.length, revisedCursor, revisedBlocks.length);
  return events;
};

export type FolioContentAlignmentStep<
  Block extends FolioContentBlock = FolioContentBlock,
> =
  | { type: "pair"; baseBlock: Block; revisedBlock: Block }
  | { type: "baseOnly"; block: Block }
  | { type: "revisedOnly"; block: Block }
  | { type: "baseRow"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "revisedRow"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "baseTable"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | { type: "revisedTable"; blocks: readonly Block[]; location: FolioContentTableLocation }
  | TableColumnAlignmentStep<Block>;

type DocumentSegment<Block extends FolioContentBlock> =
  | { kind: "body"; blocks: Block[]; containerPathKey: string | null }
  | { kind: "table"; blocks: Block[]; containerPathKey: null };

const containerPathKeyOf = (block: FolioContentBlock): string | null =>
  block.containerPath === undefined
    ? null
    : JSON.stringify(block.containerPath.map(({ kind, id }) => [kind, id]));

const splitSegments = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): DocumentSegment<Block>[] => {
  const segments: DocumentSegment<Block>[] = [];
  let currentTableIndex: number | null = null;
  for (const block of blocks) {
    const tableIndex = block.table?.outerTableIndex ?? null;
    const containerPathKey = block.table ? null : containerPathKeyOf(block);
    const current = segments.at(-1);
    if (
      current !== undefined &&
      currentTableIndex === tableIndex &&
      current.containerPathKey === containerPathKey
    ) {
      current.blocks.push(block);
      continue;
    }
    if (block.table && segments.at(-1)?.kind !== "body") {
      segments.push({ kind: "body", blocks: [], containerPathKey: null });
    }
    segments.push(
      block.table
        ? { kind: "table", blocks: [block], containerPathKey: null }
        : { kind: "body", blocks: [block], containerPathKey },
    );
    currentTableIndex = tableIndex;
  }
  if (segments.length === 0 || segments.at(-1)?.kind === "table") {
    segments.push({ kind: "body", blocks: [], containerPathKey: null });
  }
  return segments;
};

export const groupFolioContentTableRows = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): Block[][] => {
  const rows = new Map<string, Block[]>();
  for (const block of blocks) {
    if (!block.table) {
      continue;
    }
    const key = `${String(block.table.tableIndex)}:${String(block.table.rowIndex)}`;
    const row = rows.get(key);
    if (row) {
      row.push(block);
    } else {
      rows.set(key, [block]);
    }
  }
  return [...rows.values()];
};

const groupTables = <Block extends FolioContentBlock>(blocks: readonly Block[]): Block[][] => {
  const tables = new Map<number, Block[]>();
  for (const block of blocks) {
    const tableIndex = block.table?.tableIndex;
    if (tableIndex === undefined) {
      continue;
    }
    const table = tables.get(tableIndex);
    if (table) {
      table.push(block);
    } else {
      tables.set(tableIndex, [block]);
    }
  }
  return [...tables.values()];
};

const tokenSimilarity = (left: string, right: string): number => {
  const leftTokens = left.split(/\s+/u).filter((token) => token.length > 0);
  const rightTokens = right.split(/\s+/u).filter((token) => token.length > 0);
  if (leftTokens.length === 0 || rightTokens.length === 0) {
    return 0;
  }
  const remaining = new Map<string, number>();
  for (const token of leftTokens) {
    remaining.set(token, (remaining.get(token) ?? 0) + 1);
  }
  let shared = 0;
  for (const token of rightTokens) {
    const count = remaining.get(token) ?? 0;
    if (count > 0) {
      remaining.set(token, count - 1);
      shared += 1;
    }
  }
  return (2 * shared) / (leftTokens.length + rightTokens.length);
};

const rowText = <Block extends FolioContentBlock>(row: readonly Block[]): string =>
  row.map(({ text }) => text).join(" ");

const rowLocation = <Block extends FolioContentBlock>(
  row: readonly Block[],
): FolioContentTableLocation | null => row.at(0)?.table ?? null;

const rowPhysicalCellCount = <Block extends FolioContentBlock>(row: readonly Block[]): number => {
  let count = 0;
  for (const block of row) {
    count = Math.max(count, (block.table?.cellIndex ?? -1) + 1);
  }
  return count;
};

const alignRowCells = <Block extends FolioContentBlock>(
  baseRow: readonly Block[],
  revisedRow: readonly Block[],
  baseColumnKeys?: ReadonlyMap<number, number>,
  revisedColumnKeys?: ReadonlyMap<number, number>,
): FolioContentAlignmentStep<Block>[] => {
  const byCell = (
    row: readonly Block[],
    columnKeys: ReadonlyMap<number, number> | undefined,
  ): Map<number, Block[]> => {
    const cells = new Map<number, Block[]>();
    for (const block of row) {
      const table = block.table;
      let cellIndex = table?.cellIndex ?? 0;
      if (table && columnKeys) {
        const alignedColumn = columnKeys.get(table.gridColumnIndex);
        if (alignedColumn === undefined) {
          return panic("A paired table cell has no aligned grid column", {
            gridColumnIndex: table.gridColumnIndex,
          });
        }
        cellIndex = alignedColumn;
      }
      const blocks = cells.get(cellIndex);
      if (blocks) {
        blocks.push(block);
      } else {
        cells.set(cellIndex, [block]);
      }
    }
    return cells;
  };

  const baseCells = byCell(baseRow, baseColumnKeys);
  const revisedCells = byCell(revisedRow, revisedColumnKeys);
  const cellIndexes = [...new Set([...baseCells.keys(), ...revisedCells.keys()])].toSorted(
    (left, right) => left - right,
  );
  const steps: FolioContentAlignmentStep<Block>[] = [];
  for (const cellIndex of cellIndexes) {
    const baseBlocks = baseCells.get(cellIndex) ?? [];
    const revisedBlocks = revisedCells.get(cellIndex) ?? [];
    const paired = Math.min(baseBlocks.length, revisedBlocks.length);
    for (let index = 0; index < paired; index++) {
      const baseBlock = baseBlocks[index];
      const revisedBlock = revisedBlocks[index];
      if (baseBlock && revisedBlock) {
        if (
          (baseBlock.containerPath === undefined && revisedBlock.containerPath === undefined) ||
          contentBlocksShareContainer(baseBlock, revisedBlock)
        ) {
          steps.push({ type: "pair", baseBlock, revisedBlock });
        } else {
          steps.push({ type: "baseOnly", block: baseBlock });
          steps.push({ type: "revisedOnly", block: revisedBlock });
        }
      }
    }
    for (const block of baseBlocks.slice(paired)) {
      steps.push({ type: "baseOnly", block });
    }
    for (const block of revisedBlocks.slice(paired)) {
      steps.push({ type: "revisedOnly", block });
    }
  }
  return steps;
};

type TableRowAlignment<Block extends FolioContentBlock> =
  | { type: "pair"; baseRow: readonly Block[]; revisedRow: readonly Block[] }
  | { type: "baseOnly"; row: readonly Block[] }
  | { type: "revisedOnly"; row: readonly Block[] };

const rowSimilarity = <Block extends FolioContentBlock>(
  base: readonly Block[] | undefined,
  revised: readonly Block[] | undefined,
): number => {
  if (!base || !revised) {
    return 0;
  }
  const text = tokenSimilarity(rowText(base), rowText(revised));
  return rowPhysicalCellCount(base) === rowPhysicalCellCount(revised) ? text : text / 2;
};

const pairTableRows = <Block extends FolioContentBlock>(
  baseRows: readonly Block[][],
  revisedRows: readonly Block[][],
): TableRowAlignment<Block>[] => {
  const ROW_PAIR_SIMILARITY = 0.5;
  const aligned: TableRowAlignment<Block>[] = [];
  let baseCursor = 0;
  let revisedCursor = 0;
  while (baseCursor < baseRows.length && revisedCursor < revisedRows.length) {
    const baseRow = baseRows[baseCursor];
    const revisedRow = revisedRows[revisedCursor];
    if (!baseRow || !revisedRow) {
      break;
    }
    const here = rowSimilarity(baseRow, revisedRow);
    if (here < 1) {
      const baseAhead = rowSimilarity(baseRows[baseCursor + 1], revisedRow);
      const revisedAhead = rowSimilarity(baseRow, revisedRows[revisedCursor + 1]);
      if (baseAhead >= ROW_PAIR_SIMILARITY && baseAhead > here && baseAhead >= revisedAhead) {
        aligned.push({ type: "baseOnly", row: baseRow });
        baseCursor += 1;
        continue;
      }
      if (revisedAhead >= ROW_PAIR_SIMILARITY && revisedAhead > here) {
        aligned.push({ type: "revisedOnly", row: revisedRow });
        revisedCursor += 1;
        continue;
      }
    }
    aligned.push({ type: "pair", baseRow, revisedRow });
    baseCursor += 1;
    revisedCursor += 1;
  }
  for (const row of baseRows.slice(baseCursor)) {
    aligned.push({ type: "baseOnly", row });
  }
  for (const row of revisedRows.slice(revisedCursor)) {
    aligned.push({ type: "revisedOnly", row });
  }
  return aligned;
};

const alignTableRows = <Block extends FolioContentBlock>(
  rows: readonly TableRowAlignment<Block>[],
  baseColumnKeys?: ReadonlyMap<number, number>,
  revisedColumnKeys?: ReadonlyMap<number, number>,
): FolioContentAlignmentStep<Block>[] => {
  const steps: FolioContentAlignmentStep<Block>[] = [];
  const pushRow = (row: readonly Block[], side: "base" | "revised"): void => {
    const location = rowLocation(row);
    if (location) {
      steps.push({ type: side === "base" ? "baseRow" : "revisedRow", blocks: row, location });
    }
  };
  for (const alignment of rows) {
    switch (alignment.type) {
      case "baseOnly":
        pushRow(alignment.row, "base");
        break;
      case "revisedOnly":
        pushRow(alignment.row, "revised");
        break;
      case "pair":
        steps.push(
          ...alignRowCells(
            alignment.baseRow,
            alignment.revisedRow,
            baseColumnKeys,
            revisedColumnKeys,
          ),
        );
        break;
      default: {
        const unreachable: never = alignment;
        panic("Unhandled table row alignment", { alignment: unreachable });
      }
    }
  }
  return steps;
};

const rowCellSpansEqual = <Block extends FolioContentBlock>(
  baseRow: readonly Block[],
  revisedRow: readonly Block[],
): boolean => {
  const cells = (row: readonly Block[]): FolioContentTableLocation[] => {
    const byPhysicalIndex = new Map<number, FolioContentTableLocation>();
    for (const block of row) {
      const table = block.table ?? panic("A table row contains a body block");
      if (!byPhysicalIndex.has(table.cellIndex)) {
        byPhysicalIndex.set(table.cellIndex, table);
      }
    }
    return [...byPhysicalIndex.values()];
  };
  const baseCells = cells(baseRow);
  const revisedCells = cells(revisedRow);
  if (baseCells.length !== revisedCells.length) {
    return false;
  }
  return baseCells.every((base, index) => {
    const revised = revisedCells[index];
    return (
      revised !== undefined &&
      base.gridColumnIndex === revised.gridColumnIndex &&
      base.columnSpan === revised.columnSpan &&
      base.rowSpan === revised.rowSpan
    );
  });
};

const rowHasVerticalSpan = <Block extends FolioContentBlock>(row: readonly Block[]): boolean =>
  row.some(({ table }) => table !== undefined && table.rowSpan > 1);

type TableStructurePlan<Block extends FolioContentBlock> = {
  steps: FolioContentAlignmentStep<Block>[];
  representable: boolean;
};

const buildTablePlan = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
): TableStructurePlan<Block> => {
  const columns = alignTableColumns(baseBlocks, revisedBlocks);
  const rows = pairTableRows(
    groupFolioContentTableRows(columns?.baseBlocks ?? baseBlocks),
    groupFolioContentTableRows(columns?.revisedBlocks ?? revisedBlocks),
  );
  const representable = rows.every((row) => {
    if (row.type === "pair") {
      return columns !== null || rowCellSpansEqual(row.baseRow, row.revisedRow);
    }
    return !rowHasVerticalSpan(row.row);
  });
  return {
    representable,
    steps: [
      ...(columns?.steps ?? []),
      ...alignTableRows(rows, columns?.baseColumnKeys, columns?.revisedColumnKeys),
    ],
  };
};

const buildTableSegmentPlan = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
): TableStructurePlan<Block> => {
  const baseTables = groupTables(baseBlocks);
  const revisedTables = groupTables(revisedBlocks);
  const steps: FolioContentAlignmentStep<Block>[] = [];
  const paired = Math.min(baseTables.length, revisedTables.length);
  let representable = baseTables.length === revisedTables.length;
  for (let index = 0; index < paired; index++) {
    const table = buildTablePlan(baseTables[index] ?? [], revisedTables[index] ?? []);
    steps.push(...table.steps);
    representable &&= table.representable;
  }
  for (const blocks of baseTables.slice(paired)) {
    const location = blocks.at(0)?.table;
    if (location) {
      steps.push({ type: "baseTable", blocks, location });
    }
  }
  for (const blocks of revisedTables.slice(paired)) {
    const location = blocks.at(0)?.table;
    if (location) {
      steps.push({ type: "revisedTable", blocks, location });
    }
  }
  return { steps, representable };
};

const segmentText = <Block extends FolioContentBlock>(segment: DocumentSegment<Block>): string =>
  segment.blocks.map(({ text }) => text).join(" ");

const segmentStructuralKey = <Block extends FolioContentBlock>(
  segment: DocumentSegment<Block>,
): string => JSON.stringify([segment.kind, segment.containerPathKey]);

const segmentIndexesByKey = <Block extends FolioContentBlock>(
  segments: readonly DocumentSegment<Block>[],
): ReadonlyMap<string, readonly number[]> => {
  const indexes = new Map<string, number[]>();
  segments.forEach((segment, index) => {
    const key = segmentStructuralKey(segment);
    const existing = indexes.get(key);
    if (existing) {
      existing.push(index);
    } else {
      indexes.set(key, [index]);
    }
  });
  return indexes;
};

const firstIndexAfter = (indexes: readonly number[] | undefined, cursor: number): number | null => {
  if (!indexes) {
    return null;
  }
  let low = 0;
  let high = indexes.length;
  while (low < high) {
    const middle = low + Math.floor((high - low) / 2);
    if ((indexes[middle] ?? -1) <= cursor) {
      low = middle + 1;
    } else {
      high = middle;
    }
  }
  return indexes[low] ?? null;
};

const segmentsCanPair = <Block extends FolioContentBlock>(
  base: DocumentSegment<Block>,
  revised: DocumentSegment<Block>,
): boolean => {
  if (base.kind !== revised.kind) {
    return false;
  }
  if (base.blocks.length === 0 || revised.blocks.length === 0) {
    return true;
  }
  if (base.containerPathKey === null && revised.containerPathKey === null) {
    return true;
  }
  return base.containerPathKey === revised.containerPathKey;
};

const alignSegments = <Block extends FolioContentBlock>(
  baseSegments: readonly DocumentSegment<Block>[],
  revisedSegments: readonly DocumentSegment<Block>[],
): { baseSegment: DocumentSegment<Block> | null; revisedSegment: DocumentSegment<Block> | null }[] => {
  const TABLE_PAIR_SIMILARITY = 0.5;
  const baseIndexesByKey = segmentIndexesByKey(baseSegments);
  const revisedIndexesByKey = segmentIndexesByKey(revisedSegments);
  const paired: {
    baseSegment: DocumentSegment<Block> | null;
    revisedSegment: DocumentSegment<Block> | null;
  }[] = [];
  const similarity = (
    left: DocumentSegment<Block> | undefined,
    right: DocumentSegment<Block> | undefined,
  ): number =>
    left === undefined || right === undefined || left.kind !== right.kind
      ? 0
      : tokenSimilarity(segmentText(left), segmentText(right));
  let baseCursor = 0;
  let revisedCursor = 0;
  while (baseCursor < baseSegments.length && revisedCursor < revisedSegments.length) {
    const baseSegment = baseSegments[baseCursor];
    const revisedSegment = revisedSegments[revisedCursor];
    if (!baseSegment || !revisedSegment) {
      break;
    }
    if (!segmentsCanPair(baseSegment, revisedSegment)) {
      const nextBaseIndex = firstIndexAfter(
        baseIndexesByKey.get(segmentStructuralKey(revisedSegment)),
        baseCursor,
      );
      const nextRevisedIndex = firstIndexAfter(
        revisedIndexesByKey.get(segmentStructuralKey(baseSegment)),
        revisedCursor,
      );
      const baseDistance =
        nextBaseIndex === null ? Number.POSITIVE_INFINITY : nextBaseIndex - baseCursor;
      const revisedDistance =
        nextRevisedIndex === null ? Number.POSITIVE_INFINITY : nextRevisedIndex - revisedCursor;
      if (revisedDistance < baseDistance) {
        paired.push({ baseSegment: null, revisedSegment });
        revisedCursor += 1;
      } else {
        paired.push({ baseSegment, revisedSegment: null });
        baseCursor += 1;
      }
      continue;
    }
    if (baseSegment.kind === "body") {
      paired.push({ baseSegment, revisedSegment });
      baseCursor += 1;
      revisedCursor += 1;
      continue;
    }
    const here = similarity(baseSegment, revisedSegment);
    const baseAhead = similarity(baseSegments[baseCursor + 2], revisedSegment);
    const revisedAhead = similarity(baseSegment, revisedSegments[revisedCursor + 2]);
    if (baseAhead >= TABLE_PAIR_SIMILARITY && baseAhead > here && baseAhead >= revisedAhead) {
      paired.push({ baseSegment, revisedSegment: null });
      baseCursor += 1;
      continue;
    }
    if (revisedAhead >= TABLE_PAIR_SIMILARITY && revisedAhead > here) {
      paired.push({ baseSegment: null, revisedSegment });
      revisedCursor += 1;
      continue;
    }
    paired.push({ baseSegment, revisedSegment });
    baseCursor += 1;
    revisedCursor += 1;
  }
  for (const segment of baseSegments.slice(baseCursor)) {
    paired.push({ baseSegment: segment, revisedSegment: null });
  }
  for (const segment of revisedSegments.slice(revisedCursor)) {
    paired.push({ baseSegment: null, revisedSegment: segment });
  }
  return paired;
};

const unpairedSegmentSteps = <Block extends FolioContentBlock>(
  segment: DocumentSegment<Block>,
  side: "base" | "revised",
): FolioContentAlignmentStep<Block>[] => {
  if (segment.kind !== "table") {
    return segment.blocks.map((block) =>
      side === "base" ? { type: "baseOnly", block } : { type: "revisedOnly", block },
    );
  }
  const location = segment.blocks.at(0)?.table;
  if (!location) {
    return [];
  }
  return [
    { type: side === "base" ? "baseTable" : "revisedTable", blocks: segment.blocks, location },
  ];
};

export type AlignFolioContentStructureOptions<Block extends FolioContentBlock> = {
  baseBlocks: readonly Block[];
  revisedBlocks: readonly Block[];
  workSession?: FolioContentAlignmentWorkSession;
  wholeTableReplacement?: "allow" | "avoid";
  /** @internal Compatibility hook for adapters whose source model predates `idStability`. */
  idStability?: ((block: Block) => FolioContentIdStability) | undefined;
};

export const alignFolioContentStructure = <Block extends FolioContentBlock>({
  baseBlocks,
  revisedBlocks,
  workSession = createFolioContentAlignmentWorkSession(),
  wholeTableReplacement = "allow",
  idStability,
}: AlignFolioContentStructureOptions<Block>): FolioContentAlignmentStep<Block>[] => {
  const canReplaceWholeTable =
    wholeTableReplacement === "allow" && baseBlocks.some(({ table }) => table === undefined);
  const steps: FolioContentAlignmentStep<Block>[] = [];
  for (const { baseSegment, revisedSegment } of alignSegments(
    splitSegments(baseBlocks),
    splitSegments(revisedBlocks),
  )) {
    if (baseSegment && revisedSegment) {
      if (baseSegment.kind !== "table") {
        steps.push(
          ...alignFolioContentBlocks(baseSegment.blocks, revisedSegment.blocks, {
            workSession,
            idStability,
          }),
        );
        continue;
      }
      const table = buildTableSegmentPlan(baseSegment.blocks, revisedSegment.blocks);
      if (canReplaceWholeTable && !table.representable) {
        steps.push(...unpairedSegmentSteps(baseSegment, "base"));
        steps.push(...unpairedSegmentSteps(revisedSegment, "revised"));
        continue;
      }
      steps.push(...table.steps);
      continue;
    }
    if (baseSegment) {
      steps.push(...unpairedSegmentSteps(baseSegment, "base"));
      continue;
    }
    if (revisedSegment) {
      steps.push(...unpairedSegmentSteps(revisedSegment, "revised"));
    }
  }
  return steps;
};

export const contentBlocksShareContainer = (
  left: FolioContentBlock,
  right: FolioContentBlock,
): boolean => {
  if (left.containerPath !== undefined || right.containerPath !== undefined) {
    if (left.containerPath === undefined || right.containerPath === undefined) {
      return false;
    }
    return (
      left.containerPath.length === right.containerPath.length &&
      left.containerPath.every((entry, index) => {
        const candidate = right.containerPath?.[index];
        return candidate?.kind === entry.kind && candidate.id === entry.id;
      })
    );
  }
  if (!left.table || !right.table) {
    return left.table === undefined && right.table === undefined;
  }
  return (
    left.table.tableIndex === right.table.tableIndex &&
    left.table.rowIndex === right.table.rowIndex &&
    left.table.cellIndex === right.table.cellIndex
  );
};
