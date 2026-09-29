/**
 * Pasting table cells into a table: the pasted cells overwrite the cells they
 * land on, merges the pasted block cuts through are split first, and the table
 * grows when the block reaches past its edge. A paste into a cell selection
 * fills every selected cell instead and leaves the table's shape alone.
 *
 * `prosemirror-tables` ships the same operation as its `handlePaste`, but its
 * merge splitting inserts the lower half of a vertical merge at a position
 * relative to the table rather than to the document, so in any table that does
 * not open the document the new cell lands inside an earlier cell and the
 * paste either throws or leaves rows that no longer tile the grid. Its split
 * halves also copy every attribute of the cell they came from, so both halves
 * claim the same source cell and the same stored merge continuation.
 *
 * With suggestions on, the paste is tracked: each overwritten cell keeps its
 * content as a deletion before the pasted content, a vertical merge the block
 * cuts through is split as a tracked cell merge, and the rows and cells the
 * table grows by are tracked insertions, so rejecting every change gives the
 * table back. A horizontal merge has no tracked split, so one the block cuts
 * through is split directly, as the structural table commands do.
 */

import { Fragment, Slice, type Node as PMNode, type Schema } from "prosemirror-model";
import type { EditorState, Transaction } from "prosemirror-state";
import { CellSelection, TableMap, cellAround, tableNodeTypes, type Rect } from "prosemirror-tables";
import { Transform } from "prosemirror-transform";

import {
  decodeTableCellParagraphSourcePayload,
  transportTableCellsWithParagraphPropertySources,
} from "../docx/paragraphPropertySource";
import type { TableCell } from "../types/document";
import { expectTableAttrs, mergeTableAttrs } from "./attrs";

/** A rectangular block of cells, one fragment of cells per row. */
export type PastedCells = {
  width: number;
  height: number;
  rows: Fragment[];
};

/** What a tracked paste records its changes as. */
export type TableCellPasteRevision = {
  revisionId: number;
  author: string;
  date: string;
};

export type TableCellPasteHooks = {
  /**
   * The revision a split merge, a grown row or a grown column is tracked
   * under; untracked when absent.
   */
  revision?: TableCellPasteRevision;
  /**
   * Replace the content of the cell at `cellPos` with `slice`. Untracked
   * pastes replace it outright; a tracked paste keeps the old content as a
   * deletion.
   */
  replaceCellContent?: CellContentReplacer;
};

export type CellContentReplacer = (tr: Transaction, cellPos: number, slice: Slice) => void;

const fitSlice = (nodeType: PMNode["type"], slice: Slice): PMNode => {
  const node = nodeType.createAndFill();
  if (!node) {
    throw new RangeError(`Cannot create an empty ${nodeType.name}`);
  }
  return new Transform(node).replace(0, node.content.size, slice).doc;
};

/**
 * A cell's attributes with `count` of its columns, from the `from`th on, taken
 * out of its span and its column widths.
 */
const withoutColumns = (attrs: PMNode["attrs"], from: number, count: number): PMNode["attrs"] => {
  const colwidth: unknown = attrs["colwidth"];
  const widths = Array.isArray(colwidth) ? colwidth.toSpliced(from, count) : null;
  return {
    ...attrs,
    colspan: Math.max(1, Number(attrs["colspan"]) || 1) - count,
    colwidth: widths?.some((width) => Number(width) > 0) ? widths : null,
  };
};

const spanOf = (cell: PMNode, name: "colspan" | "rowspan"): number =>
  Math.max(1, Number(cell.attrs[name]) || 1);

/** Pad short rows with empty cells so the block is rectangular. */
const ensureRectangular = (schema: Schema, rows: Fragment[]): PastedCells => {
  const widths: number[] = [];
  for (const [index, row] of rows.entries()) {
    for (let child = row.childCount - 1; child >= 0; child--) {
      const cell = row.child(child);
      for (let covered = index; covered < index + spanOf(cell, "rowspan"); covered++) {
        widths[covered] = (widths[covered] ?? 0) + spanOf(cell, "colspan");
      }
    }
  }
  const width = Math.max(0, ...widths.map((value) => value ?? 0));
  for (let index = 0; index < widths.length; index++) {
    if (index >= rows.length) {
      rows.push(Fragment.empty);
    }
    const rowWidth = widths[index] ?? 0;
    if (rowWidth < width) {
      const cells: PMNode[] = [];
      for (let column = rowWidth; column < width; column++) {
        const empty = tableNodeTypes(schema).cell.createAndFill();
        if (empty) {
          cells.push(empty);
        }
      }
      rows[index] = (rows[index] ?? Fragment.empty).append(Fragment.from(cells));
    }
  }
  return { height: rows.length, width, rows };
};

/**
 * The attributes of a cell made from another one, split off it or pasted from
 * a copy of it: the same formatting, but not the other cell's identity, its
 * revisions, or the continuation cells its merge stored, whose paragraphs
 * belong to the other cell.
 */
const newCellAttrs = (attrs: PMNode["attrs"]): PMNode["attrs"] => ({
  ...attrs,
  _docxCellId: null,
  cellMarker: null,
  tcPrChange: null,
  _preserveVMergeRestart: null,
  _docxVMergeContinuationCells: null,
});

const withNewTableCells = (node: PMNode): PMNode => {
  if (node.isTextblock || node.isLeaf) {
    return node;
  }
  const content = withNewTableCellsIn(node.content);
  const role = node.type.spec["tableRole"];
  const isCell = role === "cell" || role === "header_cell";
  return content === node.content && !isCell
    ? node
    : node.type.create(isCell ? newCellAttrs(node.attrs) : node.attrs, content, node.marks);
};

const withNewTableCellsIn = (fragment: Fragment): Fragment => {
  let changed = false;
  const nodes: PMNode[] = [];
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Fragment.forEach
  fragment.forEach((node) => {
    const next = withNewTableCells(node);
    changed ||= next !== node;
    nodes.push(next);
  });
  return changed ? Fragment.from(nodes) : fragment;
};

/**
 * A pasted slice whose table cells are new cells: a copy of a cell must not
 * claim the source cell's identity, revisions, or the continuation cells its
 * merge stored, whatever the paste lands on.
 */
export const pastedSliceWithNewTableCells = (slice: Slice): Slice => {
  const content = withNewTableCellsIn(slice.content);
  return content === slice.content ? slice : new Slice(content, slice.openStart, slice.openEnd);
};

/** The cells a slice holds, or null when it holds none (it is not table content). */
export const pastedCells = (slice: Slice): PastedCells | null => {
  if (slice.size === 0) {
    return null;
  }
  let { content, openStart, openEnd } = slice;
  while (
    content.childCount === 1 &&
    ((openStart > 0 && openEnd > 0) || content.child(0).type.spec["tableRole"] === "table")
  ) {
    openStart--;
    openEnd--;
    content = content.child(0).content;
  }
  const first = content.firstChild;
  if (!first) {
    return null;
  }
  const role = first.type.spec["tableRole"];
  const { schema } = first.type;
  const rows: Fragment[] = [];
  if (role === "row") {
    for (let index = 0; index < content.childCount; index++) {
      let cells = content.child(index).content;
      const left = index ? 0 : Math.max(0, openStart - 1);
      const right = index < content.childCount - 1 ? 0 : Math.max(0, openEnd - 1);
      if (left || right) {
        cells = fitSlice(tableNodeTypes(schema).row, new Slice(cells, left, right)).content;
      }
      rows.push(cells);
    }
  } else if (role === "cell" || role === "header_cell") {
    rows.push(
      openStart || openEnd
        ? fitSlice(tableNodeTypes(schema).row, new Slice(content, openStart, openEnd)).content
        : content,
    );
  } else {
    return null;
  }
  const newCells = rows.map((row) => {
    const cells: PMNode[] = [];
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Fragment.forEach
    row.forEach((cell) => {
      cells.push(cell.type.create(newCellAttrs(cell.attrs), cell.content, cell.marks));
    });
    return Fragment.from(cells);
  });
  return ensureRectangular(schema, newCells);
};

type TableContext = {
  table: PMNode;
  map: TableMap;
  /** The position just inside the table node. */
  tableStart: number;
  /** Maps pre-recompute positions forward. */
  mapFrom: number;
};

const readTable = (tr: Transaction, tableStart: number): TableContext => {
  const table = tr.doc.nodeAt(tableStart - 1);
  if (!table || table.type.spec["tableRole"] !== "table") {
    throw new RangeError("The pasted-into table is gone");
  }
  return { table, map: TableMap.get(table), tableStart, mapFrom: tr.mapping.maps.length };
};

const mapPosition = (tr: Transaction, context: TableContext, pos: number): number =>
  tr.mapping.slice(context.mapFrom).map(pos);

/** A new, empty cell of the same kind as `like`, with nothing that names another cell. */
const emptyCellLike = (like: PMNode | null, schema: Schema): PMNode => {
  const types = tableNodeTypes(schema);
  const type = like?.type === types.header_cell ? types.header_cell : types.cell;
  const cell = type.createAndFill();
  if (!cell) {
    throw new RangeError("Cannot create an empty table cell");
  }
  return cell;
};

const isOmittedGridSlot = (cell: PMNode): boolean => cell.attrs["_omittedGridSlot"] != null;

const hasOmittedGridSlots = (table: PMNode): boolean => {
  let found = false;
  // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
  table.forEach((row) => {
    // oxlint-disable-next-line unicorn/no-array-for-each -- ProseMirror Node.forEach
    row.forEach((cell) => {
      found ||= isOmittedGridSlot(cell);
    });
  });
  return found;
};

/**
 * The `w:vMerge` continuation cells a merged cell stored for the rows below its
 * first, one per row, empty ones standing in for any it did not store.
 */
const continuationCells = (cell: PMNode, rowspan: number): TableCell[] => {
  const stored = cell.attrs["_docxVMergeContinuationCells"];
  const cells =
    stored === undefined || stored === null
      ? []
      : [...decodeTableCellParagraphSourcePayload(stored).cells];
  while (cells.length < rowspan - 1) {
    cells.push({
      type: "tableCell",
      formatting: { vMerge: "continue" },
      content: [{ type: "paragraph", content: [] }],
    });
  }
  return cells;
};

/** Stored continuation cells to carry, or null for none. */
const continuationAttr = (cells: TableCell[]) =>
  cells.length > 0 ? transportTableCellsWithParagraphPropertySources(cells) : null;

/** Add columns and rows so the table reaches `width` × `height`. */
const growTable = (
  tr: Transaction,
  context: TableContext,
  width: number,
  height: number,
  revision: TableCellPasteRevision | undefined,
): boolean => {
  const { table, map, tableStart } = context;
  const { schema } = tr.doc.type;
  let grown = false;
  if (width > map.width) {
    let rowEnd = 0;
    for (let row = 0; row < map.height; row++) {
      const rowNode = table.child(row);
      rowEnd += rowNode.nodeSize;
      const cells: PMNode[] = [];
      for (let column = map.width; column < width; column++) {
        const cell = emptyCellLike(rowNode.lastChild, schema);
        cells.push(
          revision
            ? cell.type.create(
                { ...cell.attrs, cellMarker: { kind: "ins", info: { ...revision } } },
                cell.content,
              )
            : cell,
        );
      }
      tr.insert(mapPosition(tr, context, rowEnd - 1 + tableStart), cells);
    }
    const widths = expectTableAttrs(table).columnWidths;
    const lastWidth = widths?.at(-1);
    tr.setNodeMarkup(
      mapPosition(tr, context, tableStart - 1),
      undefined,
      mergeTableAttrs(table, {
        columnWidths:
          widths?.length === map.width && lastWidth !== undefined
            ? [...widths, ...Array.from({ length: width - map.width }, () => lastWidth)]
            : undefined,
      }),
    );
    grown = true;
  }
  if (height > map.height) {
    const lastRowStart = (map.height - 1) * map.width;
    const cells: PMNode[] = [];
    for (let column = 0; column < Math.max(map.width, width); column++) {
      const above = column < map.width ? table.nodeAt(map.map[lastRowStart + column] ?? -1) : null;
      cells.push(emptyCellLike(above ?? null, schema));
    }
    const rowType = tableNodeTypes(schema).row;
    const rows: PMNode[] = [];
    for (let row = map.height; row < height; row++) {
      rows.push(rowType.create(revision ? { trIns: { ...revision } } : null, cells));
    }
    tr.insert(mapPosition(tr, context, tableStart + table.nodeSize - 2), rows);
    grown = true;
  }
  return grown;
};

/** Split every vertical merge that crosses the line above row `top`, within [left, right). */
const isolateHorizontal = (
  tr: Transaction,
  context: TableContext,
  left: number,
  right: number,
  top: number,
  revision: TableCellPasteRevision | undefined,
): boolean => {
  const { table, map, tableStart } = context;
  if (top === 0 || top === map.height) {
    return false;
  }
  let found = false;
  for (let column = left; column < right; column++) {
    const index = top * map.width + column;
    const pos = map.map[index];
    if (pos === undefined || map.map[index - map.width] !== pos) {
      continue;
    }
    found = true;
    const cell = table.nodeAt(pos);
    if (!cell) {
      continue;
    }
    const { top: cellTop, left: cellLeft } = map.findCell(pos);
    const rowspan = spanOf(cell, "rowspan");
    const upperRows = top - cellTop;
    const lowerRows = rowspan - upperRows;
    const continuations = continuationCells(cell, rowspan);
    tr.setNodeMarkup(mapPosition(tr, context, tableStart + pos), null, {
      ...cell.attrs,
      rowspan: upperRows,
      _docxVMergeContinuationCells: continuationAttr(continuations.slice(0, upperRows - 1)),
    });
    const lower = cell.type.createAndFill({
      ...newCellAttrs(cell.attrs),
      rowspan: lowerRows,
      _docxVMergeContinuationCells: continuationAttr(continuations.slice(upperRows)),
      cellMarker: revision
        ? {
            kind: "merge",
            info: { ...revision },
            verticalMergeOriginal: "continue",
          }
        : null,
    });
    if (!lower) {
      throw new RangeError("Cannot split a merged table cell");
    }
    tr.insert(mapPosition(tr, context, tableStart + map.positionAt(top, cellLeft, table)), lower);
    column += spanOf(cell, "colspan") - 1;
  }
  return found;
};

/** Split every horizontal merge that crosses the line left of column `left`, within [top, bottom). */
const isolateVertical = (
  tr: Transaction,
  context: TableContext,
  top: number,
  bottom: number,
  left: number,
): boolean => {
  const { table, map, tableStart } = context;
  if (left === 0 || left === map.width) {
    return false;
  }
  let found = false;
  for (let row = top; row < bottom; row++) {
    const index = row * map.width + left;
    const pos = map.map[index];
    if (pos === undefined || map.map[index - 1] !== pos) {
      continue;
    }
    found = true;
    const cell = table.nodeAt(pos);
    if (!cell) {
      continue;
    }
    const cellLeft = map.colCount(pos);
    const colspan = spanOf(cell, "colspan");
    const updatePos = mapPosition(tr, context, tableStart + pos);
    tr.setNodeMarkup(
      updatePos,
      null,
      withoutColumns(cell.attrs, left - cellLeft, colspan - (left - cellLeft)),
    );
    const right = cell.type.createAndFill(
      withoutColumns(newCellAttrs(cell.attrs), 0, left - cellLeft),
    );
    if (!right) {
      throw new RangeError("Cannot split a merged table cell");
    }
    tr.insert(updatePos + cell.nodeSize, right);
    row += spanOf(cell, "rowspan") - 1;
  }
  return found;
};

/**
 * The range of a cell's content a paste replaces. A block of whole paragraphs
 * (a pasted cell's content) replaces all of it; open content (text, or
 * paragraphs cut from a longer run) goes from the start of the first paragraph
 * to the end of the last, so it takes on the cell's paragraph formatting the
 * way a paste into a paragraph does.
 */
export const cellPasteRange = (
  cell: PMNode,
  cellPos: number,
  slice: Slice,
): { from: number; to: number } => {
  const from = cellPos + 1;
  const to = cellPos + cell.nodeSize - 1;
  if (slice.openStart === 0 && slice.openEnd === 0 && slice.content.firstChild?.isBlock) {
    return { from, to };
  }
  return {
    from: cell.firstChild?.isTextblock ? from + 1 : from,
    to: cell.lastChild?.isTextblock ? to - 1 : to,
  };
};

/** Replace a cell's content outright. */
const replaceCellContentDirectly = (tr: Transaction, cellPos: number, slice: Slice): void => {
  const cell = tr.doc.nodeAt(cellPos);
  if (!cell) {
    throw new RangeError("The pasted-into table cell is gone");
  }
  const { from, to } = cellPasteRange(cell, cellPos, slice);
  tr.replace(from, to, slice);
};

type CellTarget = { pos: number; slice: Slice };

/**
 * The table cell each pasted cell lands on, when every one lands on a cell of
 * its own shape; null when the block's cells are shaped differently.
 */
const matchingCells = (
  { map, tableStart }: TableContext,
  cells: PastedCells,
  top: number,
  left: number,
): CellTarget[] | null => {
  const matches: CellTarget[] = [];
  const covered = new Set<number>();
  for (const [offset, fragment] of cells.rows.entries()) {
    const row = top + offset;
    let column = left;
    for (let child = 0; child < fragment.childCount; child++) {
      const pasted = fragment.child(child);
      while (covered.has(row * map.width + column)) {
        column++;
      }
      const pos = map.map[row * map.width + column];
      if (pos === undefined) {
        return null;
      }
      const rect = map.findCell(pos);
      if (
        rect.top !== row ||
        rect.left !== column ||
        rect.right - rect.left !== spanOf(pasted, "colspan") ||
        rect.bottom - rect.top !== spanOf(pasted, "rowspan")
      ) {
        return null;
      }
      for (let r = rect.top; r < rect.bottom; r++) {
        for (let c = rect.left; c < rect.right; c++) {
          covered.add(r * map.width + c);
        }
      }
      matches.push({ pos: tableStart + pos, slice: new Slice(pasted.content, 0, 0) });
      column = rect.right;
    }
  }
  return matches;
};

/** The pasted cell covering (`row`, `column`) of the block, repeating the block past its edge. */
const pastedCellAt = (cells: PastedCells, row: number, column: number): PMNode | null => {
  const wantedRow = row % cells.height;
  const wantedColumn = column % cells.width;
  const covered: number[][] = [];
  for (const [index, fragment] of cells.rows.entries()) {
    let at = 0;
    for (let child = 0; child < fragment.childCount; child++) {
      const cell = fragment.child(child);
      while (covered[index]?.[at]) {
        at++;
      }
      for (let r = index; r < index + spanOf(cell, "rowspan"); r++) {
        for (let c = at; c < at + spanOf(cell, "colspan"); c++) {
          (covered[r] ??= [])[c] = 1;
        }
      }
      if (
        wantedRow >= index &&
        wantedRow < index + spanOf(cell, "rowspan") &&
        wantedColumn >= at &&
        wantedColumn < at + spanOf(cell, "colspan")
      ) {
        return cell;
      }
      at += spanOf(cell, "colspan");
    }
  }
  return null;
};

/**
 * Every cell of the table touching `rect`, each taking the pasted cell at its
 * place in the block (the block repeating to cover the rectangle), or the
 * pasted content itself when it is not cells. The table's shape is untouched.
 */
const fillTargets = (
  { map, table, tableStart }: TableContext,
  rect: Rect,
  cells: PastedCells | null,
  slice: Slice,
): CellTarget[] => {
  const targets: CellTarget[] = [];
  for (const pos of map.cellsInRect(rect)) {
    const cellRect = map.findCell(pos);
    const row = Math.max(cellRect.top, rect.top) - rect.top;
    const column = Math.max(cellRect.left, rect.left) - rect.left;
    const pasted = cells ? pastedCellAt(cells, row, column) : null;
    const cell = table.nodeAt(pos);
    if ((cells && !pasted) || !cell || isOmittedGridSlot(cell)) {
      continue;
    }
    targets.push({
      pos: tableStart + pos,
      slice: pasted ? new Slice(pasted.content, 0, 0) : slice,
    });
  }
  return targets;
};

const replaceTargets = (
  tr: Transaction,
  targets: readonly CellTarget[],
  replaceCellContent: CellContentReplacer,
): void => {
  // Last cell first, so the earlier positions stay valid.
  for (const { pos, slice } of targets.toSorted((a, b) => b.pos - a.pos)) {
    replaceCellContent(tr, pos, slice);
  }
};

const selectBlock = (tr: Transaction, tableStart: number, rect: Rect): void => {
  const { map, table } = readTable(tr, tableStart);
  tr.setSelection(
    new CellSelection(
      tr.doc.resolve(tableStart + map.positionAt(rect.top, rect.left, table)),
      tr.doc.resolve(tableStart + map.positionAt(rect.bottom - 1, rect.right - 1, table)),
    ),
  );
};

/**
 * Paste a block of cells into the table at `tableStart` from (`top`, `left`):
 * grow the table to hold it, split the merges its edges cut through, then give
 * each cell the block covers its pasted cell. Where the block's cells have the
 * table's shapes the table keeps its cells, with their formatting and
 * identity, and takes the pasted content; otherwise the block's cells replace
 * them, except in a tracked paste, which keeps the table's shape. Selects the
 * pasted block.
 */
export const insertTableCells = (
  tr: Transaction,
  tableStart: number,
  { top, left }: Pick<Rect, "top" | "left">,
  cells: PastedCells,
  { revision, replaceCellContent = replaceCellContentDirectly }: TableCellPasteHooks = {},
): Transaction => {
  let rect = { top, left, right: left + cells.width, bottom: top + cells.height };
  let context = readTable(tr, tableStart);
  if (hasOmittedGridSlots(context.table)) {
    // Absent grid positions (`w:gridBefore`, `w:gridAfter`) hold no cell to
    // split, grow past or paste into: the block fills the real cells it
    // covers and the table keeps its shape.
    rect = {
      ...rect,
      right: Math.min(rect.right, context.map.width),
      bottom: Math.min(rect.bottom, context.map.height),
    };
    replaceTargets(tr, fillTargets(context, rect, cells, Slice.empty), replaceCellContent);
    selectBlock(tr, tableStart, rect);
    return tr;
  }
  const recompute = (changed: boolean) => {
    if (changed) {
      context = readTable(tr, tableStart);
    }
  };
  recompute(growTable(tr, context, rect.right, rect.bottom, revision));
  recompute(isolateHorizontal(tr, context, rect.left, rect.right, rect.top, revision));
  recompute(isolateHorizontal(tr, context, rect.left, rect.right, rect.bottom, revision));
  recompute(isolateVertical(tr, context, rect.top, rect.bottom, rect.left));
  recompute(isolateVertical(tr, context, rect.top, rect.bottom, rect.right));

  const matches = matchingCells(context, cells, rect.top, rect.left);
  if (matches) {
    replaceTargets(tr, matches, replaceCellContent);
  } else if (revision) {
    replaceTargets(tr, fillTargets(context, rect, cells, Slice.empty), replaceCellContent);
  } else {
    for (let row = rect.top; row < rect.bottom; row++) {
      const from = context.map.positionAt(row, rect.left, context.table);
      const to = context.map.positionAt(row, rect.right, context.table);
      tr.replace(
        mapPosition(tr, context, from + tableStart),
        mapPosition(tr, context, to + tableStart),
        new Slice(cells.rows[row - rect.top] ?? Fragment.empty, 0, 0),
      );
    }
  }
  selectBlock(tr, tableStart, rect);
  return tr;
};

/**
 * Paste into every cell of a cell selection: each takes the pasted cell at its
 * place in the block, the block repeating to cover the selection, or the
 * pasted content itself when it holds no cells. The table keeps its shape.
 */
export const fillSelectedCells = (
  tr: Transaction,
  selection: CellSelection,
  slice: Slice,
  { replaceCellContent = replaceCellContentDirectly }: TableCellPasteHooks = {},
): Transaction => {
  const tableStart = selection.$anchorCell.start(-1);
  const context = readTable(tr, tableStart);
  const rect = context.map.rectBetween(
    selection.$anchorCell.pos - tableStart,
    selection.$headCell.pos - tableStart,
  );
  replaceTargets(tr, fillTargets(context, rect, pastedCells(slice), slice), replaceCellContent);
  const after = readTable(tr, tableStart);
  tr.setSelection(
    new CellSelection(
      tr.doc.resolve(tableStart + after.map.positionAt(rect.top, rect.left, after.table)),
      tr.doc.resolve(
        tableStart + after.map.positionAt(rect.bottom - 1, rect.right - 1, after.table),
      ),
    ),
  );
  return tr;
};

/**
 * Paste `slice` into the table the selection is in: into every selected cell
 * of a cell selection, or, when the slice holds cells, as a block of cells
 * from the cell holding the cursor. False otherwise, so the paste falls
 * through to the ordinary one.
 */
export const pasteTableCells = (
  state: EditorState,
  slice: Slice,
  dispatch?: (tr: Transaction) => void,
  hooks?: TableCellPasteHooks,
): boolean => {
  const { selection } = state;
  let run: ((tr: Transaction) => Transaction) | null = null;
  if (selection instanceof CellSelection) {
    run = (tr) => fillSelectedCells(tr, selection, slice, hooks);
  } else {
    const cells = pastedCells(slice);
    const $cell = cells ? cellAround(selection.$head) : null;
    if (cells && $cell) {
      const tableStart = $cell.start(-1);
      const rect = TableMap.get($cell.node(-1)).findCell($cell.pos - tableStart);
      run = (tr) => insertTableCells(tr, tableStart, rect, cells, hooks);
    }
  }
  if (!run) {
    return false;
  }
  if (dispatch) {
    const tr = run(state.tr);
    dispatch(tr.scrollIntoView().setMeta("paste", true).setMeta("uiEvent", "paste"));
  }
  return true;
};
