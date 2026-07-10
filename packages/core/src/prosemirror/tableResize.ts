/**
 * Table-resize commit + width-read helpers shared by the React and Vue
 * adapters. Pure `(view, …) → result` over the PM doc — the gesture state
 * machines (which DOM handle is grabbed, pixel deltas, listener wiring) stay
 * in each adapter.
 *
 * Widths and heights are stored in twips (1/20 of a point) on the PM doc,
 * matching OOXML's `w:tblGrid` and `w:tcW` units.
 */

import type { Node as PMNode } from "prosemirror-model";
import type { EditorView } from "prosemirror-view";

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
export function readRowHeight(view: EditorView, pmStart: number, rowIndex: number): number | null {
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

/**
 * Bake a column-resize into the PM doc: update the table's `columnWidths`
 * attr and every cell at `colIdx` / `colIdx + 1` in every row.
 */
export function commitColumnResize(
  view: EditorView,
  opts: { pmStart: number; colIdx: number; newLeft: number; newRight: number },
): void {
  const { pmStart, colIdx, newLeft, newRight } = opts;
  const found = findTableAt(view, pmStart);
  if (!found) return;
  const { node, tablePos } = found;

  const tr = view.state.tr;
  const widths = [...((node.attrs["columnWidths"] as number[]) ?? [])];
  widths[colIdx] = newLeft;
  widths[colIdx + 1] = newRight;
  // Switch to fixed layout so Word honors the explicit widths — autofit would
  // recompute columns to fit content and discard the resize (issue #781).
  tr.setNodeMarkup(tablePos, undefined, {
    ...node.attrs,
    columnWidths: widths,
    tableLayout: "fixed",
  });

  let rowOffset = tablePos + 1;
  node.forEach((row) => {
    let cellOffset = rowOffset + 1;
    let cellColIdx = 0;
    row.forEach((cell) => {
      const colspan = (cell.attrs["colspan"] as number) || 1;
      if (cellColIdx === colIdx || cellColIdx === colIdx + 1) {
        const newWidth = cellColIdx === colIdx ? newLeft : newRight;
        tr.setNodeMarkup(tr.mapping.map(cellOffset), undefined, {
          ...cell.attrs,
          width: newWidth,
          widthType: "dxa",
          colwidth: null,
        });
      }
      cellOffset += cell.nodeSize;
      cellColIdx += colspan;
    });
    rowOffset += row.nodeSize;
  });

  view.dispatch(tr);
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

/**
 * Bake a right-edge resize into the PM doc: grow only the last column.
 * Updates `columnWidths[colIdx]` and the cell at `colIdx` in every row.
 */
export function commitRightEdgeResize(
  view: EditorView,
  opts: { pmStart: number; colIdx: number; newWidth: number },
): void {
  const { pmStart, colIdx, newWidth } = opts;
  const found = findTableAt(view, pmStart);
  if (!found) return;
  const { node, tablePos } = found;

  const tr = view.state.tr;
  const widths = [...((node.attrs["columnWidths"] as number[]) ?? [])];
  widths[colIdx] = newWidth;
  // Fixed layout so Word honors the explicit widths (issue #781).
  tr.setNodeMarkup(tablePos, undefined, {
    ...node.attrs,
    columnWidths: widths,
    tableLayout: "fixed",
  });

  let rowOffset = tablePos + 1;
  node.forEach((row) => {
    let cellOffset = rowOffset + 1;
    let cellColIdx = 0;
    row.forEach((cell) => {
      const colspan = (cell.attrs["colspan"] as number) || 1;
      if (cellColIdx === colIdx) {
        tr.setNodeMarkup(tr.mapping.map(cellOffset), undefined, {
          ...cell.attrs,
          width: newWidth,
          widthType: "dxa",
          colwidth: null,
        });
      }
      cellOffset += cell.nodeSize;
      cellColIdx += colspan;
    });
    rowOffset += row.nodeSize;
  });

  view.dispatch(tr);
}
