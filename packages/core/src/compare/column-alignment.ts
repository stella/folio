import { panic } from "better-result";

import type { FolioContentBlock, FolioContentTableLocation } from "./content-types";

const MAX_TABLE_GRID_COLUMNS = 63;
const MAX_TABLE_GRID_AREA = 1_000_000;

type GridCell<Block extends FolioContentBlock> = {
  blocks: Block[];
  rowIndex: number;
  gridColumnIndex: number;
  columnSpan: number;
  rowSpan: number;
};

type GridColumn<Block extends FolioContentBlock> = {
  index: number;
  structureSignature: string;
  cellTextCounts: readonly ReadonlyMap<number, number>[];
  /** Shared references to per-cell identity sets; spanning cells are not copied per column. */
  stableIdentityKeyGroups: readonly (readonly number[])[];
  ownedCells: readonly GridCell<Block>[];
};

type TableGrid<Block extends FolioContentBlock> = {
  cells: readonly GridCell<Block>[];
  width: number;
  height: number;
};

export type TableColumnAlignmentStep<Block extends FolioContentBlock = FolioContentBlock> =
  | {
      type: "baseColumn";
      blocks: readonly Block[];
      location: FolioContentTableLocation;
      columnIndex: number;
    }
  | {
      type: "revisedColumn";
      blocks: readonly Block[];
      location: FolioContentTableLocation;
      columnIndex: number;
      anchor: { blockId: string; position: "after" | "before" };
    };

export type TableColumnAlignment<Block extends FolioContentBlock = FolioContentBlock> = {
  steps: TableColumnAlignmentStep<Block>[];
  baseBlocks: Block[];
  revisedBlocks: Block[];
  baseColumnKeys: ReadonlyMap<number, number>;
  revisedColumnKeys: ReadonlyMap<number, number>;
};

const extractTableGrid = <Block extends FolioContentBlock>(
  blocks: readonly Block[],
): TableGrid<Block> | null => {
  const cellsByPhysicalLocation = new Map<string, GridCell<Block>>();
  let width = 0;
  let height = 0;
  for (const block of blocks) {
    const location = block.table;
    if (!location) {
      return null;
    }
    const key = `${String(location.rowIndex)}:${String(location.cellIndex)}`;
    const existing = cellsByPhysicalLocation.get(key);
    if (existing) {
      existing.blocks.push(block);
      continue;
    }
    const right = location.gridColumnIndex + location.columnSpan;
    const bottom = location.rowIndex + location.rowSpan;
    if (
      !Number.isSafeInteger(location.gridColumnIndex) ||
      !Number.isSafeInteger(location.rowIndex) ||
      !Number.isSafeInteger(location.columnSpan) ||
      !Number.isSafeInteger(location.rowSpan) ||
      !Number.isSafeInteger(right) ||
      !Number.isSafeInteger(bottom) ||
      location.gridColumnIndex < 0 ||
      location.rowIndex < 0 ||
      location.columnSpan < 1 ||
      location.rowSpan < 1 ||
      right > MAX_TABLE_GRID_COLUMNS
    ) {
      return null;
    }
    cellsByPhysicalLocation.set(key, {
      blocks: [block],
      rowIndex: location.rowIndex,
      gridColumnIndex: location.gridColumnIndex,
      columnSpan: location.columnSpan,
      rowSpan: location.rowSpan,
    });
    width = Math.max(width, right);
    height = Math.max(height, bottom);
  }
  if (width === 0 || height === 0 || width * height > MAX_TABLE_GRID_AREA) {
    return null;
  }

  // The 63-column ceiling fits one bigint mask. A two-event sweep validates
  // rectangle overlap in O(cells log cells) storage and never materializes
  // empty rows in a sparse grid.
  const occupancyEvents = [...cellsByPhysicalLocation.values()].flatMap((cell) => {
    const mask = ((1n << BigInt(cell.columnSpan)) - 1n) << BigInt(cell.gridColumnIndex);
    return [
      { rowIndex: cell.rowIndex, type: "enter" as const, mask },
      { rowIndex: cell.rowIndex + cell.rowSpan, type: "leave" as const, mask },
    ];
  });
  occupancyEvents.sort((left, right) => {
    const rowOrder = left.rowIndex - right.rowIndex;
    if (rowOrder !== 0 || left.type === right.type) {
      return rowOrder;
    }
    return left.type === "leave" ? -1 : 1;
  });
  let occupiedColumns = 0n;
  for (const event of occupancyEvents) {
    if (event.type === "leave") {
      occupiedColumns &= ~event.mask;
      continue;
    }
    if ((occupiedColumns & event.mask) !== 0n) {
      return null;
    }
    occupiedColumns |= event.mask;
  }

  return { cells: [...cellsByPhysicalLocation.values()], width, height };
};

const tableGridColumns = <Block extends FolioContentBlock>(
  grid: TableGrid<Block>,
  internText: (text: string) => number,
  internStableIdentity: (identity: string) => number,
): GridColumn<Block>[] => {
  const coveringCellsByColumn: GridCell<Block>[][] = Array.from({ length: grid.width }, () => []);
  const ownedCellsByColumn: GridCell<Block>[][] = Array.from({ length: grid.width }, () => []);
  const textCountsByCell = new Map<GridCell<Block>, ReadonlyMap<number, number>>();
  const stableIdentityKeysByCell = new Map<GridCell<Block>, readonly number[]>();
  for (const cell of grid.cells) {
    const counts = new Map<number, number>();
    for (const { text } of cell.blocks) {
      const key = internText(text);
      counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    textCountsByCell.set(cell, counts);
    stableIdentityKeysByCell.set(
      cell,
      cell.blocks.flatMap(({ identity }) =>
        identity.type === "positional"
          ? []
          : [internStableIdentity(`${identity.type}\u0000${identity.id}`)],
      ),
    );
    ownedCellsByColumn[cell.gridColumnIndex]?.push(cell);
    for (
      let column = cell.gridColumnIndex;
      column < cell.gridColumnIndex + cell.columnSpan;
      column++
    ) {
      coveringCellsByColumn[column]?.push(cell);
    }
  }

  return Array.from({ length: grid.width }, (_unused, columnIndex) => {
    const coveringCells = [...(coveringCellsByColumn[columnIndex] ?? [])].toSorted(
      (left, right) => left.rowIndex - right.rowIndex,
    );
    const structuralCells = coveringCells.map((cell) => [
      columnIndex - cell.gridColumnIndex,
      cell.columnSpan,
      cell.rowIndex,
      cell.rowSpan,
    ]);
    return {
      index: columnIndex,
      structureSignature: JSON.stringify([grid.height, structuralCells]),
      cellTextCounts: coveringCells.map(
        (cell) =>
          textCountsByCell.get(cell) ?? panic("A table cell has no interned text signature"),
      ),
      stableIdentityKeyGroups: coveringCells.map(
        (cell) =>
          stableIdentityKeysByCell.get(cell) ??
          panic("A table cell has no interned identity signature"),
      ),
      ownedCells: ownedCellsByColumn[columnIndex] ?? [],
    };
  });
};

const exactCellContentEvidence = <Block extends FolioContentBlock>(
  left: GridColumn<Block>,
  right: GridColumn<Block>,
): number => {
  let evidence = 0;
  for (let cellIndex = 0; cellIndex < left.cellTextCounts.length; cellIndex++) {
    const leftCounts = left.cellTextCounts[cellIndex];
    const rightCounts = right.cellTextCounts[cellIndex];
    if (!leftCounts || !rightCounts) continue;
    const [smaller, larger] =
      leftCounts.size <= rightCounts.size ? [leftCounts, rightCounts] : [rightCounts, leftCounts];
    for (const [key, count] of smaller) {
      evidence += Math.min(count, larger.get(key) ?? 0);
    }
  }
  return evidence;
};

/**
 * The sole highest-text-evidence structural embedding that preserves every
 * stable identity visible on both sides, or `null` when the evidence ties.
 */
const uniqueColumnEmbedding = <Block extends FolioContentBlock>(
  shorter: readonly GridColumn<Block>[],
  wider: readonly GridColumn<Block>[],
  sharedStableIdentityKeys: ReadonlySet<number>,
): number[] | null => {
  const stableGroupKeyByGroup = new Map<readonly number[], number>();
  const stableGroupKeys = new Map<string, number>();
  let nextStableGroupKey = 0;
  const stableGroupKey = (group: readonly number[]): number => {
    const cached = stableGroupKeyByGroup.get(group);
    if (cached !== undefined) return cached;
    const signature = JSON.stringify(
      group
        .filter((key) => sharedStableIdentityKeys.has(key))
        .toSorted((left, right) => left - right),
    );
    const existing = stableGroupKeys.get(signature);
    const key = existing ?? nextStableGroupKey++;
    if (existing === undefined) stableGroupKeys.set(signature, key);
    stableGroupKeyByGroup.set(group, key);
    return key;
  };
  const sharedIdentitySignatures = (columns: readonly GridColumn<Block>[]): string[] =>
    columns.map(({ stableIdentityKeyGroups }) =>
      JSON.stringify(stableIdentityKeyGroups.map(stableGroupKey)),
    );
  const shorterIdentitySignatures = sharedIdentitySignatures(shorter);
  const widerIdentitySignatures = sharedIdentitySignatures(wider);
  const scores = Array.from({ length: shorter.length + 1 }, () =>
    Array.from({ length: wider.length + 1 }, () => -1),
  );
  const counts = Array.from({ length: shorter.length + 1 }, () =>
    Array.from({ length: wider.length + 1 }, () => 0),
  );
  for (let wideIndex = 0; wideIndex <= wider.length; wideIndex++) {
    const finalScores = scores[shorter.length];
    const finalRow = counts[shorter.length];
    if (finalScores && finalRow) {
      finalScores[wideIndex] = 0;
      finalRow[wideIndex] = 1;
    }
  }
  for (let shortIndex = shorter.length - 1; shortIndex >= 0; shortIndex--) {
    for (let wideIndex = wider.length - 1; wideIndex >= 0; wideIndex--) {
      const shortColumn = shorter[shortIndex];
      const wideColumn = wider[wideIndex];
      const scoreRow = scores[shortIndex];
      const countRow = counts[shortIndex];
      if (!shortColumn || !wideColumn || !scoreRow || !countRow) {
        continue;
      }
      const skipScore = scores[shortIndex]?.[wideIndex + 1] ?? -1;
      const skipCount = counts[shortIndex]?.[wideIndex + 1] ?? 0;
      const remainingScore = scores[shortIndex + 1]?.[wideIndex + 1] ?? -1;
      const remainingCount = counts[shortIndex + 1]?.[wideIndex + 1] ?? 0;
      const canMatch =
        shortColumn.structureSignature === wideColumn.structureSignature &&
        shorterIdentitySignatures[shortIndex] === widerIdentitySignatures[wideIndex] &&
        remainingScore >= 0 &&
        remainingCount > 0;
      const matchScore = canMatch
        ? remainingScore + exactCellContentEvidence(shortColumn, wideColumn)
        : -1;
      const bestScore = Math.max(skipScore, matchScore);
      scoreRow[wideIndex] = bestScore;
      countRow[wideIndex] = Math.min(
        2,
        (skipScore === bestScore ? skipCount : 0) + (matchScore === bestScore ? remainingCount : 0),
      );
    }
  }
  if (counts[0]?.[0] !== 1) {
    return null;
  }
  const mapping: number[] = [];
  let shortIndex = 0;
  let wideIndex = 0;
  while (shortIndex < shorter.length) {
    if (wideIndex >= wider.length) {
      return panic("A unique column embedding ended before every column was mapped");
    }
    const shortColumn = shorter[shortIndex];
    const wideColumn = wider[wideIndex];
    const remainingScore = scores[shortIndex + 1]?.[wideIndex + 1] ?? -1;
    const remainingCount = counts[shortIndex + 1]?.[wideIndex + 1] ?? 0;
    const matchScore =
      shortColumn &&
      wideColumn &&
      shortColumn.structureSignature === wideColumn.structureSignature &&
      shorterIdentitySignatures[shortIndex] === widerIdentitySignatures[wideIndex] &&
      remainingScore >= 0 &&
      remainingCount > 0
        ? remainingScore + exactCellContentEvidence(shortColumn, wideColumn)
        : -1;
    const bestScore = scores[shortIndex]?.[wideIndex] ?? -1;
    const skipScore = scores[shortIndex]?.[wideIndex + 1] ?? -1;
    if (matchScore === bestScore && skipScore !== bestScore) {
      mapping.push(wideIndex);
      shortIndex += 1;
    }
    wideIndex += 1;
  }
  return mapping;
};

const columnOwnedBlocks = <Block extends FolioContentBlock>({
  ownedCells,
}: GridColumn<Block>): Block[] | null => {
  if (ownedCells.length === 0 || ownedCells.some(({ columnSpan }) => columnSpan !== 1)) {
    return null;
  }
  return ownedCells.flatMap(({ blocks }) => blocks);
};

export const alignTableColumns = <Block extends FolioContentBlock>(
  baseBlocks: readonly Block[],
  revisedBlocks: readonly Block[],
): TableColumnAlignment<Block> | null => {
  const baseGrid = extractTableGrid(baseBlocks);
  const revisedGrid = extractTableGrid(revisedBlocks);
  if (!baseGrid || !revisedGrid || baseGrid.width === revisedGrid.width) {
    return null;
  }
  // Shared numeric keys preserve cross-side text equality without copying a
  // potentially long string into every column covered by a spanning cell.
  const textKeys = new Map<string, number>();
  let nextTextKey = 0;
  const internText = (text: string): number => {
    const existing = textKeys.get(text);
    if (existing !== undefined) {
      return existing;
    }
    const key = nextTextKey++;
    textKeys.set(text, key);
    return key;
  };
  const stableIdentityKeys = new Map<string, number>();
  let nextStableIdentityKey = 0;
  const internStableIdentity = (identity: string): number => {
    const existing = stableIdentityKeys.get(identity);
    if (existing !== undefined) return existing;
    const key = nextStableIdentityKey++;
    stableIdentityKeys.set(identity, key);
    return key;
  };
  const baseColumns = tableGridColumns(baseGrid, internText, internStableIdentity);
  const revisedColumns = tableGridColumns(revisedGrid, internText, internStableIdentity);
  const stableIdentityKeysOf = (columns: readonly GridColumn<Block>[]): Set<number> => {
    const keys = new Set<number>();
    const seenGroups = new Set<readonly number[]>();
    for (const { stableIdentityKeyGroups } of columns) {
      for (const group of stableIdentityKeyGroups) {
        if (seenGroups.has(group)) continue;
        seenGroups.add(group);
        for (const key of group) keys.add(key);
      }
    }
    return keys;
  };
  const baseStableIdentityKeys = stableIdentityKeysOf(baseColumns);
  const revisedStableIdentityKeys = stableIdentityKeysOf(revisedColumns);
  const sharedStableIdentityKeys = new Set(
    [...baseStableIdentityKeys].filter((key) => revisedStableIdentityKeys.has(key)),
  );
  const revisedIsWider = revisedColumns.length > baseColumns.length;
  const mapping = uniqueColumnEmbedding(
    revisedIsWider ? baseColumns : revisedColumns,
    revisedIsWider ? revisedColumns : baseColumns,
    sharedStableIdentityKeys,
  );
  if (!mapping) {
    return null;
  }
  const widerColumns = revisedIsWider ? revisedColumns : baseColumns;
  const mapped = new Set(mapping);
  const unmatchedColumns = widerColumns.filter((_column, index) => !mapped.has(index));
  const ownedColumns: { column: GridColumn<Block>; blocks: Block[] }[] = [];
  for (const column of unmatchedColumns) {
    const blocks = columnOwnedBlocks(column);
    if (!blocks) {
      return null;
    }
    ownedColumns.push({ column, blocks });
  }
  if (
    ownedColumns.length === 0 ||
    (revisedIsWider &&
      unmatchedColumns.some(({ ownedCells }) => ownedCells.some(({ rowSpan }) => rowSpan !== 1)))
  ) {
    return null;
  }

  const unmatchedIds = new Set(
    ownedColumns.flatMap(({ blocks }) => blocks.map(({ identity }) => identity.id)),
  );
  const baseColumnKeys = new Map<number, number>();
  const revisedColumnKeys = new Map<number, number>();
  mapping.forEach((wideIndex, shortIndex) => {
    if (revisedIsWider) {
      baseColumnKeys.set(shortIndex, shortIndex);
      revisedColumnKeys.set(wideIndex, shortIndex);
    } else {
      baseColumnKeys.set(wideIndex, shortIndex);
      revisedColumnKeys.set(shortIndex, shortIndex);
    }
  });

  const steps: TableColumnAlignmentStep<Block>[] = [];
  if (!revisedIsWider) {
    for (const { column, blocks } of ownedColumns) {
      const location = blocks.at(0)?.table;
      if (!location) {
        return null;
      }
      steps.push({ type: "baseColumn", blocks, location, columnIndex: column.index });
    }
  } else {
    for (const { column, blocks } of ownedColumns) {
      const location = blocks.at(0)?.table;
      const rightBaseIndex = mapping.findIndex((targetIndex) => targetIndex > column.index);
      const leftBaseIndex = mapping.findLastIndex((targetIndex) => targetIndex < column.index);
      const anchorColumn =
        rightBaseIndex >= 0 ? baseColumns[rightBaseIndex] : baseColumns[leftBaseIndex];
      const anchorBlock = anchorColumn?.ownedCells.at(0)?.blocks.at(0);
      if (!location || !anchorBlock) {
        return null;
      }
      steps.push({
        type: "revisedColumn",
        blocks,
        location,
        columnIndex: column.index,
        anchor: {
          blockId: anchorBlock.identity.id,
          position: rightBaseIndex >= 0 ? "before" : "after",
        },
      });
    }
  }
  return {
    steps,
    baseBlocks: baseBlocks.filter(
      ({ identity }) => revisedIsWider || !unmatchedIds.has(identity.id),
    ),
    revisedBlocks: revisedBlocks.filter(
      ({ identity }) => !revisedIsWider || !unmatchedIds.has(identity.id),
    ),
    baseColumnKeys,
    revisedColumnKeys,
  };
};
