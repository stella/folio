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
  signature: string;
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
  occupancyEvents.sort(
    (left, right) =>
      left.rowIndex - right.rowIndex ||
      (left.type === right.type ? 0 : left.type === "leave" ? -1 : 1),
  );
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
): GridColumn<Block>[] => {
  const coveringCellsByColumn: GridCell<Block>[][] = Array.from(
    { length: grid.width },
    () => [],
  );
  const ownedCellsByColumn: GridCell<Block>[][] = Array.from(
    { length: grid.width },
    () => [],
  );
  const textKeysByCell = new Map<GridCell<Block>, readonly number[]>();
  for (const cell of grid.cells) {
    textKeysByCell.set(cell, cell.blocks.map(({ text }) => internText(text)));
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
    const structuralCells = [...(coveringCellsByColumn[columnIndex] ?? [])]
      .toSorted((left, right) => left.rowIndex - right.rowIndex)
      .map((cell) => [
        columnIndex - cell.gridColumnIndex,
        cell.columnSpan,
        cell.rowIndex,
        cell.rowSpan,
        textKeysByCell.get(cell) ?? panic("A table cell has no interned text signature"),
      ]);
    return {
      index: columnIndex,
      signature: JSON.stringify([grid.height, structuralCells]),
      ownedCells: ownedCellsByColumn[columnIndex] ?? [],
    };
  });
};

/** The sole exact ordered embedding of `shorter` in `wider`, or null when ambiguous. */
const uniqueColumnEmbedding = <Block extends FolioContentBlock>(
  shorter: readonly GridColumn<Block>[],
  wider: readonly GridColumn<Block>[],
): number[] | null => {
  const counts = Array.from({ length: shorter.length + 1 }, () =>
    Array.from({ length: wider.length + 1 }, () => 0),
  );
  for (let wideIndex = 0; wideIndex <= wider.length; wideIndex++) {
    const finalRow = counts[shorter.length];
    if (finalRow) {
      finalRow[wideIndex] = 1;
    }
  }
  for (let shortIndex = shorter.length - 1; shortIndex >= 0; shortIndex--) {
    for (let wideIndex = wider.length - 1; wideIndex >= 0; wideIndex--) {
      const skip = counts[shortIndex]?.[wideIndex + 1] ?? 0;
      const match =
        shorter[shortIndex]?.signature === wider[wideIndex]?.signature
          ? (counts[shortIndex + 1]?.[wideIndex + 1] ?? 0)
          : 0;
      const row = counts[shortIndex];
      if (row) {
        row[wideIndex] = Math.min(2, skip + match);
      }
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
    const canMatch =
      shorter[shortIndex]?.signature === wider[wideIndex]?.signature &&
      (counts[shortIndex + 1]?.[wideIndex + 1] ?? 0) > 0;
    const canSkip = (counts[shortIndex]?.[wideIndex + 1] ?? 0) > 0;
    if (canMatch && !canSkip) {
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
  const baseColumns = tableGridColumns(baseGrid, internText);
  const revisedColumns = tableGridColumns(revisedGrid, internText);
  const revisedIsWider = revisedColumns.length > baseColumns.length;
  const mapping = uniqueColumnEmbedding(
    revisedIsWider ? baseColumns : revisedColumns,
    revisedIsWider ? revisedColumns : baseColumns,
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

  const unmatchedIds = new Set(ownedColumns.flatMap(({ blocks }) => blocks.map(({ id }) => id)));
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
        anchor: { blockId: anchorBlock.id, position: rightBaseIndex >= 0 ? "before" : "after" },
      });
    }
  }
  return {
    steps,
    baseBlocks: baseBlocks.filter(({ id }) => revisedIsWider || !unmatchedIds.has(id)),
    revisedBlocks: revisedBlocks.filter(({ id }) => !revisedIsWider || !unmatchedIds.has(id)),
    baseColumnKeys,
    revisedColumnKeys,
  };
};
