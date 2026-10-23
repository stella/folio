/**
 * Table-resize commit + width-read helpers shared by the React and Vue
 * adapters. Pure `(view, …) → result` over the PM doc — the gesture state
 * machines (which DOM handle is grabbed, pixel deltas, listener wiring) stay
 * in each adapter.
 *
 * Widths and heights are stored in twips (1/20 of a point) on the PM doc,
 * matching OOXML's `w:tblGrid` and `w:tcW` units.
 */

import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

import {
  expectTableAttrs,
  expectTableCellAttrs,
  mergeTableAttrs,
  mergeTableCellAttrs,
} from "./attrs";

/** 1px ≈ 15 twips at 96dpi (20 twips/pt × 72pt/in ÷ 96px/in). */
export const TWIPS_PER_PIXEL = 15;
/** Minimum column width (~0.2"). */
export const MIN_CELL_WIDTH_TWIPS = 300;
/** Minimum row height (~0.14"). */
export const MIN_ROW_HEIGHT_TWIPS = 200;

type FoundTable = {
  node: PMNode;
  /** PM doc position of the table node (its `before(d)`) */
  tablePos: number;
};

/** Walk up from `pmStart + 1` to find the enclosing table node. */
function findTableAt(view: EditorView, pmStart: number): FoundTable | null {
  try {
    const $pos = view.state.doc.resolve(pmStart + 1);
    for (let d = $pos.depth; d >= 0; d--) {
      const node = $pos.node(d);
      if (node.type.name === "table") {
        return { node, tablePos: $pos.before(d) };
      }
    }
  } catch {
    // Resolution failed (stale pos after edit).
  }
  return null;
}

/** Read the [left, right] column widths at `colIndex` in the table starting at `pmStart`. */
export function readColumnWidths(
  view: EditorView,
  pmStart: number,
  colIndex: number,
): { left: number; right: number } | null {
  const found = findTableAt(view, pmStart);
  if (!found) return null;
  const widths = found.node.attrs["columnWidths"] as number[] | null;
  if (!widths) return null;
  const left = widths[colIndex];
  const right = widths[colIndex + 1];
  if (left === undefined || right === undefined) return null;
  return { left, right };
}

/**
 * Read the row height (in twips) for `rowIndex` in the table starting at
 * `pmStart`. Returns null if the row has no explicit height — the caller
 * can fall back to measuring the rendered DOM cell.
 */
export function readRowHeight(
  view: EditorView,
  pmStart: number,
  rowIndex: number,
): number | null {
  const found = findTableAt(view, pmStart);
  if (!found) return null;
  let rowHeight: number | null = null;
  let idx = 0;
  found.node.forEach((child) => {
    if (idx === rowIndex) {
      const h = child.attrs["height"] as number | null;
      if (h) rowHeight = h;
    }
    idx++;
  });
  return rowHeight;
}

/** Read the last-column width (the one being resized from the table's right edge). */
export function readColumnWidthAt(
  view: EditorView,
  pmStart: number,
  colIndex: number,
): number | null {
  const found = findTableAt(view, pmStart);
  if (!found) return null;
  const widths = found.node.attrs["columnWidths"] as number[] | null;
  if (!widths) return null;
  const width = widths[colIndex];
  if (width === undefined) return null;
  return width;
}

type CommitGridResizeOptions = {
  view: EditorView;
  table: FoundTable;
  widths: number[];
  firstColumn: number;
  lastColumn: number;
};

const commitGridResize = ({
  view,
  table: { node, tablePos },
  widths,
  firstColumn,
  lastColumn,
}: CommitGridResizeOptions): void => {
  const tr = view.state.tr;
  tr.setNodeMarkup(
    tablePos,
    undefined,
    mergeTableAttrs(node, {
      columnWidths: widths,
      _originalFormatting: {
        ...expectTableAttrs(node)._originalFormatting,
        layout: "fixed",
      },
    }),
  );

  let rowOffset = tablePos + 1;
  node.forEach((row) => {
    let cellOffset = rowOffset + 1;
    let cellColumn = 0;
    row.forEach((cell) => {
      const { colspan } = expectTableCellAttrs(cell);
      if (!Number.isInteger(colspan) || colspan < 1) {
        panic("A resized table cell must span a positive number of columns.");
      }
      const endColumn = cellColumn + colspan;
      if (cellColumn <= lastColumn && endColumn > firstColumn) {
        let width = 0;
        for (let column = cellColumn; column < endColumn; column++) {
          const coveredWidth = widths[column];
          if (
            coveredWidth === undefined ||
            !Number.isFinite(coveredWidth) ||
            coveredWidth <= 0
          ) {
            panic("A resized table cell must cover valid grid columns.");
          }
          width += coveredWidth;
        }
        tr.setNodeMarkup(
          cellOffset,
          undefined,
          mergeTableCellAttrs(cell, {
            width,
            widthType: "dxa",
            colwidth: null,
          }),
        );
      }
      cellOffset += cell.nodeSize;
      cellColumn = endColumn;
    });
    rowOffset += row.nodeSize;
  });
  view.dispatch(tr);
};

/** Commit a column-pair resize, preserving each spanning cell's full grid width. */
export function commitColumnResize(
  view: EditorView,
  opts: { pmStart: number; colIdx: number; newLeft: number; newRight: number },
): void {
  const { pmStart, colIdx, newLeft, newRight } = opts;
  const found = findTableAt(view, pmStart);
  if (!found) return;
  const sourceWidths = expectTableAttrs(found.node).columnWidths;
  if (
    !sourceWidths ||
    sourceWidths[colIdx] === undefined ||
    sourceWidths[colIdx + 1] === undefined
  )
    return;
  const widths = [...sourceWidths];
  widths[colIdx] = newLeft;
  widths[colIdx + 1] = newRight;
  commitGridResize({
    view,
    table: found,
    widths,
    firstColumn: colIdx,
    lastColumn: colIdx + 1,
  });
}

/** Bake a row-resize into the PM doc: update the target row's `height` + `heightRule`. */
export function commitRowResize(
  view: EditorView,
  opts: { pmStart: number; rowIdx: number; newHeight: number },
): void {
  const { pmStart, rowIdx, newHeight } = opts;
  const found = findTableAt(view, pmStart);
  if (!found) return;
  const { node, tablePos } = found;

  const tr = view.state.tr;
  let rowOffset = tablePos + 1;
  let idx = 0;
  node.forEach((row) => {
    if (idx === rowIdx) {
      tr.setNodeMarkup(tr.mapping.map(rowOffset), undefined, {
        ...row.attrs,
        height: newHeight,
        heightRule: "atLeast",
      });
    }
    rowOffset += row.nodeSize;
    idx++;
  });

  view.dispatch(tr);
}

/** Commit a right-edge resize, including cells that span the resized column. */
export function commitRightEdgeResize(
  view: EditorView,
  opts: { pmStart: number; colIdx: number; newWidth: number },
): void {
  const { pmStart, colIdx, newWidth } = opts;
  const found = findTableAt(view, pmStart);
  if (!found) return;
  const sourceWidths = expectTableAttrs(found.node).columnWidths;
  if (!sourceWidths || sourceWidths[colIdx] === undefined) return;
  const widths = [...sourceWidths];
  widths[colIdx] = newWidth;
  commitGridResize({
    view,
    table: found,
    widths,
    firstColumn: colIdx,
    lastColumn: colIdx,
  });
}
