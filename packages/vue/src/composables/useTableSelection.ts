// PORT-BLOCKED: divergent core API. Upstream's TableSelectionManager was a
// lightweight cell tracker — `selectCell(coords)` (no document) with a
// `{ selectedCell }` snapshot. Our fork's TableSelectionManager
// (@stll/folio-core/managers/TableSelectionManager) is document-aware:
// `selectCell(doc: Document, coords)` requires the live Document to derive a
// full TableContext, and its snapshot is
// `{ context, table, tableIndex, rowIndex, columnIndex }` with no
// `selectedCell`. This composable takes no document and its handlers accept
// only indices, so `handleCellClick` / `handleClickTarget` cannot call our
// `selectCell` without inventing a new contract (threading a Document getter
// through). The snapshot read and `isCellSelected` / `clearSelection` do map
// onto our API, but the mutating cell-selection surface does not.
//
// Resolution options (out of scope for this port):
//   - re-point this composable at the editor's Document (extend the signature
//     to accept a `() => Document`), OR
//   - add a doc-less cell-tracking manager to core for adapter parity.

/**
 * useTableSelection — Vue composable wrapping TableSelectionManager from core.
 *
 * Tracks selected table cell and provides table operations.
 */

import { ref, onScopeDispose, type Ref } from "vue";
import {
  TableSelectionManager,
  findTableFromClick,
} from "@stll/folio-core/managers/TableSelectionManager";
import type { CellCoordinates } from "@stll/folio-core/managers/types";

export interface UseTableSelectionReturn {
  selectedCell: Ref<CellCoordinates | null>;
  handleCellClick: (tableIndex: number, rowIndex: number, columnIndex: number) => void;
  handleClickTarget: (target: EventTarget | null, container?: HTMLElement | null) => void;
  clearSelection: () => void;
  isCellSelected: (tableIndex: number, rowIndex: number, columnIndex: number) => boolean;
}

export function useTableSelection(): UseTableSelectionReturn {
  const manager = new TableSelectionManager();
  const selectedCell: Ref<CellCoordinates | null> = ref(null);

  // Adapt our document-aware snapshot to the upstream `selectedCell` shape:
  // when all three indices are present the selection is a resolved cell.
  const sync = () => {
    const snap = manager.getSnapshot();
    selectedCell.value =
      snap.tableIndex !== null && snap.rowIndex !== null && snap.columnIndex !== null
        ? {
            tableIndex: snap.tableIndex,
            rowIndex: snap.rowIndex,
            columnIndex: snap.columnIndex,
          }
        : null;
  };
  const unsubscribe = manager.subscribe(sync);
  sync();

  onScopeDispose(unsubscribe);

  // PORT-BLOCKED: our `manager.selectCell` needs the live Document (see header).
  // These handlers have no document to pass, so cell selection cannot be
  // committed here. Left as documented no-ops pending the resolution above.
  function handleCellClick(_tableIndex: number, _rowIndex: number, _columnIndex: number) {
    // no-op: requires Document (divergent core API)
  }

  function handleClickTarget(target: EventTarget | null, container?: HTMLElement | null) {
    // `findTableFromClick` resolves cleanly; committing the selection does not.
    const coords = findTableFromClick(target, container);
    if (!coords) manager.clearSelection();
    // no-op on hit: requires Document (divergent core API)
  }

  function clearSelection() {
    manager.clearSelection();
  }

  function isCellSelected(tableIndex: number, rowIndex: number, columnIndex: number): boolean {
    return manager.isCellSelected(tableIndex, rowIndex, columnIndex);
  }

  return {
    selectedCell,
    handleCellClick,
    handleClickTarget,
    clearSelection,
    isCellSelected,
  };
}
