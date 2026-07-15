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

export type UseTableSelectionReturn = {
  selectedCell: Ref<CellCoordinates | null>;
  handleCellClick: (tableIndex: number, rowIndex: number, columnIndex: number) => void;
  handleClickTarget: (target: EventTarget | null, container?: HTMLElement | null) => void;
  clearSelection: () => void;
  isCellSelected: (tableIndex: number, rowIndex: number, columnIndex: number) => boolean;
};

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

  function handleCellClick(tableIndex: number, rowIndex: number, columnIndex: number) {
    manager.selectCellCoordinates({ tableIndex, rowIndex, columnIndex });
  }

  function handleClickTarget(target: EventTarget | null, container?: HTMLElement | null) {
    const coords = findTableFromClick(target, container);
    if (!coords) {
      manager.clearSelection();
      return;
    }
    manager.selectCellCoordinates(coords);
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
