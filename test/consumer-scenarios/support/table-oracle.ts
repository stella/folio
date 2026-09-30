/** A small table-grid oracle built from the reader's pre-operation snapshot. */

import type { Operation } from "./operations.ts";
import type { Row } from "./oracle.ts";

/** The scenarios run as a consumer of the published packages, so no extra dependencies. */
function panic(message: string): never {
  throw new Error(message);
}

type Location = NonNullable<Row["table"]>;
type Cell = {
  row: number;
  column: number;
  columnSpan: number;
  rowSpan: number;
  /** Only new cells have known text; pre-existing text may be edited elsewhere in a batch. */
  paragraphs?: string[];
};
type CellGeometry = Readonly<Pick<Cell, "row" | "column" | "columnSpan" | "rowSpan">>;
type Table = {
  cells: Cell[];
  height: number;
  width: number;
  /** The source grid is immutable while `cells` composes this batch. */
  sourceCells: readonly CellGeometry[];
  sourceWidth: number;
  /** Original columns; null marks one added during this batch. */
  sourceColumns: (number | null)[];
};
type Source = { table: Table; cell: Cell; row: number };

export type TableModel = {
  tables: Table[];
  sources: Map<string, Source>;
  /** Pre-state rows locate an inserted table among prose and existing tables. */
  rows: readonly Row[];
};

export class UnsupportedTableExpectation extends Error {}

const unsupported = (reason: string): never => {
  throw new UnsupportedTableExpectation(reason);
};

const lines = (text: string): string[] => text.split(/\r\n|\r|\n/u);

const cellKey = ({ tableIndex, rowIndex, cellIndex }: Location): string =>
  `${tableIndex}:${rowIndex}:${cellIndex}`;

/** Snapshot locations are repeated on each paragraph of a cell, including blanks. */
export const modelFromRows = (rows: readonly Row[]): TableModel => {
  const tables: Table[] = [];
  const sources = new Map<string, Source>();
  const cells = new Map<string, Cell>();
  for (const row of rows) {
    const location = row.table;
    if (!location) continue;
    const table = (tables[location.tableIndex] ??= {
      cells: [],
      height: 0,
      width: 0,
      sourceCells: [],
      sourceWidth: 0,
      sourceColumns: [],
    });
    const key = cellKey(location);
    let cell = cells.get(key);
    if (!cell) {
      cell = {
        row: location.rowIndex,
        column: location.gridColumnIndex,
        columnSpan: location.columnSpan,
        rowSpan: location.rowSpan,
      };
      cells.set(key, cell);
      table.cells.push(cell);
    }
    table.height = Math.max(table.height, cell.row + cell.rowSpan);
    table.width = Math.max(table.width, cell.column + cell.columnSpan);
    while (table.sourceColumns.length < table.width) {
      table.sourceColumns.push(table.sourceColumns.length);
    }
    sources.set(row.id, { table, cell, row: location.rowIndex });
  }
  if (tables.some((table) => !table)) unsupported("a table has no readable cells");
  for (const table of tables) {
    if (!table) unsupported("a table has no readable cells");
    table.sourceWidth = table.width;
    table.sourceCells = table.cells.map(({ row, column, columnSpan, rowSpan }) => ({
      row,
      column,
      columnSpan,
      rowSpan,
    }));
  }
  return { tables, sources, rows };
};

const sourceOf = (model: TableModel, blockId: unknown): Source => {
  const source = model.sources.get(String(blockId));
  if (
    !source ||
    !model.tables.includes(source.table) ||
    !source.table.cells.includes(source.cell)
  ) {
    return unsupported(`the table anchor ${String(blockId)} no longer names a cell`);
  }
  if (
    source.table.sourceColumns.length !== source.table.width ||
    source.table.sourceWidth < 1 ||
    source.table.sourceCells.length === 0
  ) {
    panic(
      `Table oracle source coordinates diverged: width ${source.table.width}, source columns ${source.table.sourceColumns.length}, source width ${source.table.sourceWidth}, source cells ${source.table.sourceCells.length}`,
    );
  }
  return source;
};

const sourceColumnAt = (table: Table, column: number): number | null => {
  const sourceColumn = table.sourceColumns.at(column);
  if (sourceColumn === undefined) {
    panic(`Table oracle has no source coordinate for column ${column}.`);
  }
  return sourceColumn;
};

const tableInsertIndex = (model: TableModel, blockId: unknown, position: "before" | "after") => {
  const rowIndex = model.rows.findIndex((row) => row.id === blockId);
  if (rowIndex < 0) return unsupported(`the insertion anchor ${String(blockId)} is absent`);
  const anchor = model.rows[rowIndex];
  if (!anchor) return unsupported("the insertion anchor is absent");
  if (anchor.table) {
    // A table insertion anchored in a nested cell is adjacent to its outermost table.
    const outerRow = model.rows.find(
      (row) => row.table?.tableIndex === anchor.table?.outerTableIndex,
    );
    if (!outerRow) return unsupported("the outer table is absent");
    const index = model.tables.indexOf(sourceOf(model, outerRow.id).table);
    return index + (position === "after" ? 1 : 0);
  }
  const priorOuter = model.rows
    .slice(0, rowIndex)
    .findLast(
      (row) => row.table !== undefined && row.table.tableIndex === row.table.outerTableIndex,
    );
  return priorOuter ? model.tables.indexOf(sourceOf(model, priorOuter.id).table) + 1 : 0;
};

const rectangularTable = (rows: readonly (readonly string[])[]): Table => {
  const width = rows[0]?.length ?? 0;
  if (width === 0 || rows.some((row) => row.length !== width)) {
    unsupported("the inserted table is not a nonempty rectangle");
  }
  const cells = rows.flatMap((row, rowIndex) =>
    row.map((text, column) => ({
      row: rowIndex,
      column,
      columnSpan: 1,
      rowSpan: 1,
      paragraphs: lines(text),
    })),
  );
  return {
    width,
    height: rows.length,
    sourceWidth: width,
    sourceCells: cells.map(({ row, column, columnSpan, rowSpan }) => ({
      row,
      column,
      columnSpan,
      rowSpan,
    })),
    sourceColumns: Array.from({ length: width }, (_, column) => column),
    cells,
  };
};

const sorted = (table: Table): Cell[] =>
  [...table.cells].sort((left, right) => left.row - right.row || left.column - right.column);

const merge = (model: TableModel, operation: Operation): void => {
  const start = sourceOf(model, operation["blockId"]);
  const end =
    operation["endBlockId"] === undefined ? undefined : sourceOf(model, operation["endBlockId"]);
  if (end && end.table !== start.table) unsupported("the merge spans two tables");
  const top = end ? Math.min(start.cell.row, end.cell.row) : start.cell.row;
  const left = end ? Math.min(start.cell.column, end.cell.column) : start.cell.column;
  const bottom = end
    ? Math.max(start.cell.row + start.cell.rowSpan, end.cell.row + end.cell.rowSpan)
    : start.cell.row + Number(operation["rowCount"]);
  const right = end
    ? Math.max(start.cell.column + start.cell.columnSpan, end.cell.column + end.cell.columnSpan)
    : start.cell.column + start.cell.columnSpan;
  if (bottom > start.table.height || right > start.table.width || bottom <= top || right <= left) {
    unsupported("the merge rectangle is outside the table");
  }
  const covered = sorted(start.table).filter(
    (cell) =>
      cell.row < bottom &&
      cell.row + cell.rowSpan > top &&
      cell.column < right &&
      cell.column + cell.columnSpan > left,
  );
  if (
    covered.some(
      (cell) =>
        cell.row < top ||
        cell.column < left ||
        cell.row + cell.rowSpan > bottom ||
        cell.column + cell.columnSpan > right,
    )
  )
    unsupported("the merge cuts an existing merged cell");
  const origin = covered[0];
  if (!origin || origin.row !== top || origin.column !== left || covered.length < 2) {
    return unsupported("the merge has no distinct top-left cell");
  }
  // The direct mutator drops a row when the merge consumes every cell in it.
  // Such a row changes row numbering and needs a larger model.
  for (let row = top + 1; row < bottom; row++) {
    if (!start.table.cells.some((cell) => cell.row === row && !covered.includes(cell))) {
      unsupported("the merge removes an entire row");
    }
  }
  origin.columnSpan = right - left;
  origin.rowSpan = bottom - top;
  const known = covered.every((cell) => cell.paragraphs !== undefined);
  if (known) {
    const first = origin.paragraphs ?? [""];
    const appended = covered
      .slice(1)
      .flatMap((cell) =>
        cell.paragraphs?.some((text) => text.length > 0) ? (cell.paragraphs ?? []) : [],
      );
    if (appended.length === 0) origin.paragraphs = first;
    else if (first.some((text) => text.length > 0)) origin.paragraphs = first.concat(appended);
    else origin.paragraphs = appended;
  } else {
    origin.paragraphs = undefined;
  }
  start.table.cells = start.table.cells.filter(
    (cell) => cell === origin || !covered.includes(cell),
  );
};

const split = (model: TableModel, operation: Operation): void => {
  const { table, cell } = sourceOf(model, operation["blockId"]);
  if (cell.rowSpan === 1 && cell.columnSpan === 1) unsupported("a 1x1 cell cannot split");
  const { row, column, rowSpan, columnSpan } = cell;
  cell.rowSpan = 1;
  cell.columnSpan = 1;
  for (let r = row; r < row + rowSpan; r++) {
    for (let c = column; c < column + columnSpan; c++) {
      if (r === row && c === column) continue;
      table.cells.push({ row: r, column: c, rowSpan: 1, columnSpan: 1, paragraphs: [""] });
    }
  }
};

/** Apply a receipt-confirmed table operation. Returns false for other operation types. */
export const applyTableOperation = (model: TableModel, operation: Operation): boolean => {
  switch (operation.type) {
    case "insertTable": {
      const position = operation["position"] === "before" ? "before" : "after";
      const index = tableInsertIndex(model, operation["blockId"], position);
      model.tables.splice(index, 0, rectangularTable(operation["rows"] as string[][]));
      return true;
    }
    case "insertSignatureTable": {
      const parties = operation["parties"] as {
        name: string;
        signatory?: string;
        title?: string;
      }[];
      const cells = parties.map((party, column) => ({
        row: 0,
        column,
        columnSpan: 1,
        rowSpan: 1,
        paragraphs: [
          party.name,
          "",
          "",
          "_".repeat(28),
          ...(party.signatory ? [party.signatory] : []),
          ...(party.title ? [party.title] : []),
        ],
      }));
      if (cells.length === 0) unsupported("the signature table has no parties");
      const position = operation["position"] === "before" ? "before" : "after";
      model.tables.splice(tableInsertIndex(model, operation["blockId"], position), 0, {
        cells,
        height: 1,
        width: cells.length,
        sourceWidth: cells.length,
        sourceCells: cells.map(({ row, column, columnSpan, rowSpan }) => ({
          row,
          column,
          columnSpan,
          rowSpan,
        })),
        sourceColumns: cells.map((_, column) => column),
      });
      return true;
    }
    case "deleteTable": {
      const { table } = sourceOf(model, operation["blockId"]);
      model.tables.splice(model.tables.indexOf(table), 1);
      return true;
    }
    case "insertTableRow": {
      const { table, cell, row: sourceRow } = sourceOf(model, operation["blockId"]);
      const boundary = cell.row + (operation["position"] === "before" ? 0 : 1);
      const crossing = table.cells.filter(
        (candidate) => candidate.row < boundary && candidate.row + candidate.rowSpan > boundary,
      );
      for (const candidate of table.cells) {
        if (crossing.includes(candidate)) candidate.rowSpan++;
        else if (candidate.row >= boundary) candidate.row++;
      }
      const free = Array.from({ length: table.width }, (_, column) => column).filter(
        (column) =>
          !crossing.some(
            (candidate) =>
              candidate.column <= column && candidate.column + candidate.columnSpan > column,
          ),
      );
      const texts = (operation["cellTexts"] as string[] | undefined) ?? [];
      const sourceBoundary = sourceRow + (operation["position"] === "before" ? 0 : 1);
      const sourceCrossing = table.sourceCells.filter(
        (candidate) =>
          candidate.row < sourceBoundary && candidate.row + candidate.rowSpan > sourceBoundary,
      );
      const sourceFree = Array.from({ length: table.sourceWidth }, (_, column) => column).filter(
        (column) =>
          !sourceCrossing.some(
            (candidate) =>
              candidate.column <= column && candidate.column + candidate.columnSpan > column,
          ),
      );
      if (texts.length > sourceFree.length) unsupported("more row texts than new cells");
      const textBySourceColumn = new Map(
        sourceFree.map((column, index) => [column, texts[index] ?? ""]),
      );
      free.forEach((column) => {
        const sourceColumn = sourceColumnAt(table, column);
        let text = "";
        if (sourceColumn !== null) {
          const sourceText = textBySourceColumn.get(sourceColumn);
          if (sourceText === undefined) {
            panic(`Table oracle source column ${sourceColumn} has no row payload slot.`);
          }
          text = sourceText;
        }
        table.cells.push({
          row: boundary,
          column,
          rowSpan: 1,
          columnSpan: 1,
          paragraphs: lines(text),
        });
      });
      table.height++;
      return true;
    }
    case "deleteTableRow": {
      const { table, cell } = sourceOf(model, operation["blockId"]);
      if (table.height <= 1 || table.cells.some((candidate) => candidate.rowSpan > 1)) {
        unsupported("deleting the last row or a row in a merged table");
      }
      table.cells = table.cells.filter((candidate) => candidate.row !== cell.row);
      for (const candidate of table.cells) if (candidate.row > cell.row) candidate.row--;
      table.height--;
      return true;
    }
    case "insertTableColumn": {
      const { table, cell } = sourceOf(model, operation["blockId"]);
      const boundary =
        operation["position"] === "before" ? cell.column : cell.column + cell.columnSpan;
      const crossing = table.cells.filter(
        (candidate) =>
          candidate.column < boundary && candidate.column + candidate.columnSpan > boundary,
      );
      for (const candidate of table.cells) {
        if (crossing.includes(candidate)) candidate.columnSpan++;
        else if (candidate.column >= boundary) candidate.column++;
      }
      const free = Array.from({ length: table.height }, (_, row) => row).filter(
        (row) =>
          !crossing.some(
            (candidate) => candidate.row <= row && candidate.row + candidate.rowSpan > row,
          ),
      );
      const texts = (operation["cellTexts"] as string[] | undefined) ?? [];
      if (texts.length > free.length) unsupported("more column texts than new cells");
      free.forEach((row, index) =>
        table.cells.push({
          row,
          column: boundary,
          rowSpan: 1,
          columnSpan: 1,
          paragraphs: lines(texts[index] ?? ""),
        }),
      );
      table.sourceColumns.splice(boundary, 0, null);
      table.width++;
      return true;
    }
    case "deleteTableColumn": {
      const { table, cell } = sourceOf(model, operation["blockId"]);
      if (table.width <= 1) unsupported("deleting the last column");
      const column = cell.column;
      table.cells = table.cells.filter((candidate) => {
        if (candidate.column > column) candidate.column--;
        else if (candidate.column + candidate.columnSpan > column) {
          if (candidate.columnSpan === 1) return false;
          candidate.columnSpan--;
        }
        return true;
      });
      table.sourceColumns.splice(column, 1);
      table.width--;
      return true;
    }
    case "mergeTableCells":
      merge(model, operation);
      return true;
    case "splitTableCell":
      split(model, operation);
      return true;
    default:
      return false;
  }
};

const geometry = (table: Table) =>
  sorted(table).map((cell) => ({
    row: cell.row,
    column: cell.column,
    columnSpan: cell.columnSpan,
    rowSpan: cell.rowSpan,
  }));

/** Compare every physical cell; blanks count even though the text oracle hides them. */
export const compareTableGeometry = (model: TableModel, actual: readonly Row[]): string[] => {
  const got = modelFromRows(actual).tables;
  const problems: string[] = [];
  if (got.length !== model.tables.length) {
    problems.push(`table count is ${got.length}, expected ${model.tables.length}`);
    return problems;
  }
  model.tables.forEach((table, index) => {
    const actualTable = got[index];
    if (!actualTable) return;
    if (
      table.height !== actualTable.height ||
      table.width !== actualTable.width ||
      JSON.stringify(geometry(table)) !== JSON.stringify(geometry(actualTable))
    ) {
      problems.push(
        `table ${index} geometry is ${JSON.stringify({ height: actualTable.height, width: actualTable.width, cells: geometry(actualTable) })}, expected ${JSON.stringify({ height: table.height, width: table.width, cells: geometry(table) })}`,
      );
      return;
    }
    const actualCells = new Map<string, string[]>();
    for (const row of actual) {
      if (row.table?.tableIndex !== index) continue;
      const key = cellKey(row.table);
      const texts = actualCells.get(key) ?? [];
      texts.push(row.text);
      actualCells.set(key, texts);
    }
    for (const cell of table.cells) {
      if (!cell.paragraphs) continue;
      const matching = actualTable.cells.find(
        (candidate) => candidate.row === cell.row && candidate.column === cell.column,
      );
      if (!matching) continue;
      const location = actual.find(
        (row) =>
          row.table?.tableIndex === index &&
          row.table.rowIndex === cell.row &&
          row.table.gridColumnIndex === cell.column,
      )?.table;
      const texts = location ? actualCells.get(cellKey(location)) : undefined;
      if (JSON.stringify(texts) !== JSON.stringify(cell.paragraphs)) {
        problems.push(
          `table ${index} cell ${cell.row}:${cell.column} paragraphs are ${JSON.stringify(texts)}, expected ${JSON.stringify(cell.paragraphs)}`,
        );
      }
    }
  });
  return problems;
};
