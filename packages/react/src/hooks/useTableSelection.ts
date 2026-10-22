/**
 * useTableSelection Hook
 *
 * Thin React binding around the framework-agnostic TableSelectionManager. The
 * manager owns the selection state and the table-operation dispatch; this hook
 * subscribes to its snapshot via `useSyncExternalStore` and forwards the host
 * `onChange` / `onSelectionChange` callbacks.
 */

import { useCallback, useMemo, useSyncExternalStore } from "react";

import {
  TableSelectionManager,
  type TableSelectionState,
} from "@stll/folio-core/managers/TableSelectionManager";
import type { Document } from "@stll/folio-core/types/document";
import type {
  TableAction,
  TableContext,
} from "@stll/folio-core/utils/tableOperations";

// ============================================================================
// TYPES
// ============================================================================

export type { TableSelectionState } from "@stll/folio-core/managers/TableSelectionManager";

export type UseTableSelectionReturn = {
  state: TableSelectionState;
  handleCellClick: (
    tableIndex: number,
    rowIndex: number,
    columnIndex: number,
  ) => void;
  handleAction: (action: TableAction) => void;
  clearSelection: () => void;
  isCellSelected: (
    tableIndex: number,
    rowIndex: number,
    columnIndex: number,
  ) => boolean;
  tableContext: TableContext | null;
};

export type UseTableSelectionOptions = {
  document: Document | null;
  onChange?: (document: Document) => void;
  onSelectionChange?: (context: TableContext | null) => void;
};

// ============================================================================
// HOOK
// ============================================================================

export function useTableSelection({
  document: doc,
  onChange,
  onSelectionChange,
}: UseTableSelectionOptions): UseTableSelectionReturn {
  const manager = useMemo(() => new TableSelectionManager(), []);
  const state = useSyncExternalStore(manager.subscribe, manager.getSnapshot);

  const handleCellClick = useCallback(
    (tableIndex: number, rowIndex: number, columnIndex: number) => {
      if (!doc) {
        return;
      }
      const context = manager.selectCell(doc, {
        tableIndex,
        rowIndex,
        columnIndex,
      });
      if (context) {
        onSelectionChange?.(context);
      }
    },
    [doc, manager, onSelectionChange],
  );

  const clearSelection = useCallback(() => {
    manager.clearSelection();
    onSelectionChange?.(null);
  }, [manager, onSelectionChange]);

  const handleAction = useCallback(
    (action: TableAction) => {
      if (!doc) {
        return;
      }
      const result = manager.handleAction(doc, action);
      if (result.type === "noop") {
        return;
      }
      if (result.type === "deleted") {
        onSelectionChange?.(null);
        onChange?.(result.document);
        return;
      }
      onChange?.(result.document);
      onSelectionChange?.(result.context);
    },
    [doc, manager, onChange, onSelectionChange],
  );

  const isCellSelected = useCallback(
    (tableIndex: number, rowIndex: number, columnIndex: number): boolean =>
      manager.isCellSelected(tableIndex, rowIndex, columnIndex),
    [manager],
  );

  return {
    state,
    handleCellClick,
    handleAction,
    clearSelection,
    isCellSelected,
    tableContext: state.context,
  };
}
