/**
 * Drag-to-select table cells, shared by the React and Vue pages-pointer
 * handlers. A pointer drag that starts inside a table cell and crosses into
 * another cell is promoted from a text selection to a prosemirror-tables
 * `CellSelection` (so multi-cell ops — delete row/column, fill, merge — become
 * reachable by dragging). Pure `(view, pmPos, clientX) → boolean` logic; the
 * adapters own the DOM event wiring and the active-view resolution (body vs
 * header/footer).
 */

import type { EditorView } from "prosemirror-view";
import { CellSelection } from "prosemirror-tables";

/**
 * @deprecated No longer used. The tracker ignores pointer overflow so same-cell
 * drags stay text selections. Kept only for backward-compatible imports; will be
 * removed in a future major.
 */
export const CELL_SELECT_OVERFLOW_PX = 5;

/**
 * Walk up from a PM position to the enclosing `tableCell`/`tableHeader` and
 * return its `before(d)` position (what `CellSelection` resolves against), or
 * null when the position isn't inside a table cell.
 */
export function findCellPosFromPmPos(view: EditorView, pmPos: number): number | null {
  try {
    const $pos = view.state.doc.resolve(pmPos);
    for (let d = $pos.depth; d > 0; d--) {
      const node = $pos.node(d);
      if (node.type.name === "tableCell" || node.type.name === "tableHeader") {
        return $pos.before(d);
      }
    }
  } catch {
    // Stale PM position after an edit.
  }
  return null;
}

/** Dispatch a `CellSelection` spanning the two cell positions. */
export function applyCellSelection(
  view: EditorView,
  anchorCellPos: number,
  headCellPos: number,
): boolean {
  try {
    const $a = view.state.doc.resolve(anchorCellPos);
    const $h = view.state.doc.resolve(headCellPos);
    view.dispatch(view.state.tr.setSelection(new CellSelection($a, $h)));
    return true;
  } catch {
    // Positions no longer resolve to cells (e.g. not inside a table).
    return false;
  }
}

/**
 * Per-gesture state machine for cell-drag selection. Create one per pointer
 * handler; drive it from mousedown/mousemove/mouseup.
 */
export type CellDragTracker = {
  /** mousedown: record the cell under the press (or null when outside a table). */
  begin(cellPos: number | null): void;
  /**
   * mousemove: returns true when it applied a `CellSelection` and the caller
   * should NOT also run its text-drag selection for this move.
   */
  update(view: EditorView, pmPos: number, clientX: number): boolean;
  /** mouseup / gesture end: clear the transient drag flags. */
  end(): void;
  /** True while a cell drag is actively producing CellSelections. */
  readonly isCellDragging: boolean;
}

export function createCellDragTracker(): CellDragTracker {
  let anchorCellPos: number | null = null;
  let cellDragging = false;

  return {
    begin(cellPos) {
      anchorCellPos = cellPos;
      cellDragging = false;
    },

    update(view, pmPos, _clientX) {
      if (anchorCellPos === null) return false;

      // Already cell-dragging: keep extending to the cell under the pointer.
      if (cellDragging) {
        const cur = findCellPosFromPmPos(view, pmPos);
        if (cur !== null) {
          applyCellSelection(view, anchorCellPos, cur);
          return true;
        }
      }

      // Crossed into a different cell → promote to a CellSelection.
      const cur = findCellPosFromPmPos(view, pmPos);
      if (cur !== null && cur !== anchorCellPos) {
        cellDragging = true;
        applyCellSelection(view, anchorCellPos, cur);
        return true;
      }

      return false;
    },

    end() {
      cellDragging = false;
    },

    get isCellDragging() {
      return cellDragging;
    },
  };
}
