/**
 * Tables with merged cells, built as packages and read back three ways, for
 * the table-operation tests.
 *
 * A table operation is only right when every reader agrees on the table it
 * leaves: the block snapshot a caller addresses cells through, the Markdown a
 * caller reads, and the package itself. Each reader here reports the same
 * projection — which cell starts where and how far it spans — so a test can
 * compare them directly, and {@link checkTableGrid} states what a well-formed
 * package table is on its own.
 */

import type { Paragraph, Table, TableCell, TableCellBlock, TableRow } from "@stll/docx-core/model";

import { parseDocx } from "../docx/parser";
import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { fromMarkdown, toMarkdown } from "../markdown";
import { FolioDocxReviewer } from "../ai-edits/headless";
import type { FolioAIBlock } from "../ai-edits/types";

/** One merged (or single) cell of a spec: where it starts and how far it spans. */
export type TableSpecCell = {
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
  /** The cell's only paragraph, or the paragraph before its nested table. */
  text: string;
  nested?: TableSpec;
};

export type TableSpec = {
  rows: number;
  columns: number;
  /** A partition of the grid: every slot is covered by exactly one cell. */
  cells: TableSpecCell[];
};

const paragraph = (text: string): Paragraph => ({
  type: "paragraph",
  formatting: {},
  content:
    text.length === 0 ? [] : [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

const COLUMN_WIDTH_TWIPS = 1400;

/** The package model of a spec, a restart/continue pair per vertical merge. */
export const tableFromSpec = (spec: TableSpec): Table => {
  const origins = new Map<string, TableSpecCell>();
  for (const cell of spec.cells) {
    origins.set(`${cell.row}:${cell.column}`, cell);
  }
  const coveringCell = (row: number, column: number): TableSpecCell | undefined =>
    spec.cells.find(
      (cell) =>
        row >= cell.row &&
        row < cell.row + cell.rowSpan &&
        column >= cell.column &&
        column < cell.column + cell.columnSpan,
    );
  const rows: TableRow[] = [];
  for (let row = 0; row < spec.rows; row++) {
    const cells: TableCell[] = [];
    for (let column = 0; column < spec.columns; column++) {
      const cell = coveringCell(row, column);
      if (!cell || cell.column !== column) {
        continue;
      }
      const starts = cell.row === row;
      const content: TableCellBlock[] = [paragraph(starts ? cell.text : "")];
      if (starts && cell.nested) {
        content.push(tableFromSpec(cell.nested), paragraph(""));
      }
      cells.push({
        type: "tableCell",
        formatting: {
          ...(cell.columnSpan > 1 && { gridSpan: cell.columnSpan }),
          ...(cell.rowSpan > 1 && { vMerge: starts ? "restart" : "continue" }),
        },
        content,
      });
    }
    rows.push({ type: "tableRow", cells });
  }
  return {
    type: "table",
    columnWidths: Array.from({ length: spec.columns }, () => COLUMN_WIDTH_TWIPS),
    rows,
  };
};

/** A package holding `Before.`, the table, and `After.`, every paragraph with an id. */
export const buildTableDocx = async (spec: TableSpec): Promise<ArrayBuffer> => {
  const document = fromMarkdown("Before.\n\nAfter.");
  const body = document.package.document;
  body.content = [body.content[0]!, tableFromSpec(spec), ...body.content.slice(1)];
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(document)))).docx;
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

/** A cell as every reader reports it: its origin slot, its spans, its text. */
export type ProjectedCell = {
  row: number;
  column: number;
  rowSpan: number;
  columnSpan: number;
  text: string;
};

/** One table: its column count and its cells, row-major by origin. */
export type ProjectedTable = {
  /** Whether the table sits in the body rather than inside another table's cell. */
  topLevel: boolean;
  columns: number;
  rows: number;
  cells: ProjectedCell[];
};

const byOrigin = (left: ProjectedCell, right: ProjectedCell): number =>
  left.row - right.row || left.column - right.column;

/**
 * The tables as the block snapshot describes them, by `tableIndex`. A cell's
 * text is its own paragraphs joined by a newline — a nested table's text
 * belongs to that table, not the cell around it.
 */
export const projectSnapshotTables = (blocks: readonly FolioAIBlock[]): ProjectedTable[] => {
  const tables = new Map<number, Map<string, ProjectedCell>>();
  const topLevel = new Set<number>();
  for (const block of blocks) {
    const table = block.table;
    if (!table) {
      continue;
    }
    if (table.outerTableIndex === table.tableIndex) {
      topLevel.add(table.tableIndex);
    }
    const cells = tables.get(table.tableIndex) ?? new Map<string, ProjectedCell>();
    tables.set(table.tableIndex, cells);
    const key = `${table.rowIndex}:${table.gridColumnIndex}`;
    const existing = cells.get(key);
    if (existing) {
      existing.text = `${existing.text}\n${block.text}`;
      continue;
    }
    cells.set(key, {
      row: table.rowIndex,
      column: table.gridColumnIndex,
      rowSpan: table.rowSpan,
      columnSpan: table.columnSpan,
      text: block.text,
    });
  }
  return [...tables.entries()]
    .sort(([left], [right]) => left - right)
    .map(([tableIndex, cells]) => {
      const list = [...cells.values()].sort(byOrigin);
      return {
        topLevel: topLevel.has(tableIndex),
        columns: Math.max(0, ...list.map((cell) => cell.column + cell.columnSpan)),
        rows: Math.max(0, ...list.map((cell) => cell.row + cell.rowSpan)),
        cells: list,
      };
    });
};

const cellParagraphText = (cell: TableCell): string =>
  cell.content
    .filter((block): block is Paragraph => block.type === "paragraph")
    .map((block) =>
      block.content
        .flatMap((child) => (child.type === "run" ? child.content : []))
        .map((content) => (content.type === "text" ? content.text : ""))
        .join(""),
    )
    .join("\n");

/** Tables of a parsed body in document order, nested tables after their parent's cell. */
const collectTables = (
  blocks: readonly unknown[],
  into: { table: Table; topLevel: boolean }[],
  topLevel = true,
): { table: Table; topLevel: boolean }[] => {
  for (const block of blocks) {
    if (typeof block !== "object" || block === null || !("type" in block)) {
      continue;
    }
    if (block.type !== "table") {
      continue;
    }
    const table = block as Table;
    into.push({ table, topLevel });
    for (const row of table.rows) {
      for (const cell of row.cells) {
        collectTables(cell.content, into, false);
      }
    }
  }
  return into;
};

export type TableGridProblem = { table: number; row: number; problem: string };

/**
 * What makes a package table well formed, and the cells it describes.
 *
 * Every row's cells (and `w:gridBefore` / `w:gridAfter`) span exactly the
 * grid; a `w:vMerge="continue"` sits under a cell of the same column and
 * width; and every row starts at least one cell of its own, because a row of
 * nothing but continuations is not one an editor can hold (the reader splits
 * the merge apart to give the row a cell, and the readers disagree).
 */
export const checkTableGrid = (
  table: Table,
  tableIndex: number,
  topLevel: boolean,
): { projection: ProjectedTable; problems: TableGridProblem[] } => {
  const problems: TableGridProblem[] = [];
  const columns = table.columnWidths?.length ?? 0;
  const cells: ProjectedCell[] = [];
  const openByColumn = new Map<number, ProjectedCell>();
  table.rows.forEach((row, rowIndex) => {
    let column = row.formatting?.gridBefore ?? 0;
    let origins = 0;
    const continued = new Set<number>();
    for (const cell of row.cells) {
      const span = cell.formatting?.gridSpan ?? 1;
      const vMerge = cell.formatting?.vMerge;
      if (vMerge === "continue") {
        const open = openByColumn.get(column);
        if (!open || open.columnSpan !== span) {
          problems.push({
            table: tableIndex,
            row: rowIndex,
            problem: `continuation at column ${column} has no merge of width ${span} above it`,
          });
        } else {
          open.rowSpan += 1;
          continued.add(column);
        }
      } else {
        const projected: ProjectedCell = {
          row: rowIndex,
          column,
          rowSpan: 1,
          columnSpan: span,
          text: cellParagraphText(cell),
        };
        cells.push(projected);
        origins += 1;
        if (vMerge === "restart") {
          openByColumn.set(column, projected);
          continued.add(column);
        }
      }
      column += span;
    }
    for (const openColumn of [...openByColumn.keys()]) {
      if (!continued.has(openColumn)) {
        openByColumn.delete(openColumn);
      }
    }
    column += row.formatting?.gridAfter ?? 0;
    if (column !== columns) {
      problems.push({
        table: tableIndex,
        row: rowIndex,
        problem: `row spans ${column} grid columns, the grid has ${columns}`,
      });
    }
    if (origins === 0) {
      problems.push({
        table: tableIndex,
        row: rowIndex,
        problem: "row holds nothing but continuations of merges from above",
      });
    }
  });
  return {
    projection: { topLevel, columns, rows: table.rows.length, cells: cells.sort(byOrigin) },
    problems,
  };
};

export type SavedTables = {
  tables: ProjectedTable[];
  problems: TableGridProblem[];
  markdown: string;
};

/** The package's own tables, checked, and the Markdown `docxToMarkdown` makes of it. */
export const readSavedTables = async (bytes: ArrayBuffer): Promise<SavedTables> => {
  const document = await parseDocx(bytes, { preloadFonts: false });
  const tables = collectTables(document.package.document.content, []);
  const checked = tables.map(({ table, topLevel }, index) =>
    checkTableGrid(table, index, topLevel),
  );
  return {
    tables: checked.map(({ projection }) => projection),
    problems: checked.flatMap(({ problems }) => problems),
    markdown: toMarkdown(document, {
      annotations: "strip",
      trackedChanges: "clean",
      comments: "strip",
      footnotes: "keep",
    }),
  };
};

/**
 * The spans of every outermost HTML table in a Markdown document, placed the
 * way HTML places cells: each one in the next slot of its row no `rowspan`
 * from above has taken. A GFM pipe table has no spans, so it projects to
 * single cells. Text is compared through the snapshot and the package; here
 * only the geometry is read.
 */
export const projectMarkdownTables = (markdown: string): Omit<ProjectedCell, "text">[][] => {
  const tables: Omit<ProjectedCell, "text">[][] = [];
  const lines = markdown.split("\n");
  let index = 0;
  while (index < lines.length) {
    const line = lines[index]!;
    if (line.startsWith("<table>")) {
      let depth = 0;
      const html: string[] = [];
      for (; index < lines.length; index++) {
        const current = lines[index]!;
        html.push(current);
        depth += (current.match(/<table>/g) ?? []).length;
        depth -= (current.match(/<\/table>/g) ?? []).length;
        if (depth === 0) {
          break;
        }
      }
      index += 1;
      tables.push(placeHtmlCells(html.join("\n")));
      continue;
    }
    if (line.startsWith("|") && lines[index + 1]?.startsWith("| ---")) {
      const rows: string[] = [line];
      index += 2;
      while (index < lines.length && lines[index]!.startsWith("|")) {
        rows.push(lines[index]!);
        index += 1;
      }
      tables.push(
        rows.flatMap((row, rowIndex) =>
          row
            .slice(1, -1)
            .split(/(?<!\\)\|/)
            .map((_cell, column) => ({ row: rowIndex, column, rowSpan: 1, columnSpan: 1 })),
        ),
      );
      continue;
    }
    index += 1;
  }
  return tables;
};

/** Top-level `<tr>` / `<td>` of one HTML table, nested tables skipped over. */
const placeHtmlCells = (html: string): Omit<ProjectedCell, "text">[] => {
  const tokens = html.match(/<\/?(table|tr|td|th)\b[^>]*>/g) ?? [];
  const cells: Omit<ProjectedCell, "text">[] = [];
  const taken = new Set<string>();
  let depth = 0;
  let row = -1;
  let column = 0;
  for (const token of tokens) {
    if (token.startsWith("<table")) {
      depth += 1;
      continue;
    }
    if (token.startsWith("</table")) {
      depth -= 1;
      continue;
    }
    if (depth !== 1) {
      continue;
    }
    if (token.startsWith("<tr")) {
      row += 1;
      column = 0;
      continue;
    }
    if (token.startsWith("<td") || token.startsWith("<th")) {
      while (taken.has(`${row}:${column}`)) {
        column += 1;
      }
      const rowSpan = Number(/rowspan="(\d+)"/.exec(token)?.[1] ?? 1);
      const columnSpan = Number(/colspan="(\d+)"/.exec(token)?.[1] ?? 1);
      cells.push({ row, column, rowSpan, columnSpan });
      for (let r = row; r < row + rowSpan; r++) {
        for (let c = column; c < column + columnSpan; c++) {
          taken.add(`${r}:${c}`);
        }
      }
      column += columnSpan;
    }
  }
  return cells;
};

const geometry = (cells: readonly Omit<ProjectedCell, "text">[]): string =>
  cells
    .map(({ row, column, rowSpan, columnSpan }) => `${row}:${column}+${rowSpan}x${columnSpan}`)
    .join(" ");

/** A reviewer's tables as the package, its reopened snapshot and its Markdown read them. */
export type TableReading = {
  bytes: ArrayBuffer;
  /** The reviewer's own snapshot, before the package was written. */
  live: ProjectedTable[];
  /** The snapshot of the package reopened. */
  snapshot: ProjectedTable[];
  saved: SavedTables;
  /** Every block's text of the reopened package, in order. */
  texts: string[];
};

/** Save a reviewer and read its tables back every way there is. */
export const readReviewerTables = async (reviewer: FolioDocxReviewer): Promise<TableReading> => {
  const live = projectSnapshotTables(reviewer.getContent());
  const bytes = await reviewer.toBuffer();
  const reopened = await FolioDocxReviewer.fromBuffer(bytes);
  const content = reopened.getContent();
  return {
    bytes,
    live,
    snapshot: projectSnapshotTables(content),
    saved: await readSavedTables(bytes),
    texts: content.map(({ text }) => text),
  };
};

/**
 * Where the readers disagree, or the package table is malformed: one line per
 * problem, empty when the table is coherent.
 */
export const tableReadingProblems = (reading: TableReading): string[] => {
  const problems = reading.saved.problems.map(
    ({ table, row, problem }) => `package table ${table} row ${row}: ${problem}`,
  );
  const snapshot = JSON.stringify(reading.snapshot);
  if (JSON.stringify(reading.saved.tables) !== snapshot) {
    problems.push(
      `snapshot ${snapshot} disagrees with package ${JSON.stringify(reading.saved.tables)}`,
    );
  }
  if (JSON.stringify(reading.live) !== snapshot) {
    problems.push(
      `live snapshot ${JSON.stringify(reading.live)} disagrees with reopened ${snapshot}`,
    );
  }
  const markdown = projectMarkdownTables(reading.saved.markdown).map(geometry);
  const topLevel = reading.saved.tables
    .filter((table) => table.topLevel)
    .map(({ cells }) => geometry(cells));
  if (JSON.stringify(markdown) !== JSON.stringify(topLevel)) {
    problems.push(
      `markdown spans ${JSON.stringify(markdown)} disagree with package ${JSON.stringify(topLevel)}`,
    );
  }
  return problems;
};
