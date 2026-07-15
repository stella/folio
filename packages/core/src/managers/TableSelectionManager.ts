/**
 * TableSelectionManager
 *
 * Framework-agnostic class for managing table cell selection state.
 * Extracted from the React `useTableSelection` hook.
 *
 * Handles:
 * - Cell selection via data-attribute queries on the DOM
 * - Table document operations (add/delete rows/columns, merge/split)
 */

import type { Document, Table } from "../types/document";
import {
  addColumn,
  addRow,
  createTableContext,
  deleteColumn,
  deleteRow,
  getColumnCount,
  mergeCells,
  splitCell,
  type TableAction,
  type TableContext,
  type TableSelection,
} from "../utils/tableOperations";
import { Subscribable } from "./Subscribable";
import type { CellCoordinates } from "./types";

// ============================================================================
// CONSTANTS
// ============================================================================

/** Data attributes for table elements in the rendered DOM */
export const TABLE_DATA_ATTRIBUTES = {
  TABLE_INDEX: "data-table-index",
  ROW_INDEX: "data-row",
  COLUMN_INDEX: "data-col",
  TABLE_CELL: "data-table-cell",
} as const;

// ============================================================================
// HELPER FUNCTIONS (framework-agnostic DOM queries)
// ============================================================================

/**
 * Find table cell coordinates from a click target by walking up the DOM
 * and reading data attributes.
 */
export function findTableFromClick(
  target: EventTarget | null,
  container?: HTMLElement | null,
): CellCoordinates | null {
  if (!(target instanceof Element)) {
    return null;
  }

  let current: Element | null = target;
  while (current && current !== container) {
    if (current.tagName === "TD" || current.tagName === "TH") {
      const rowAttr = current.getAttribute(TABLE_DATA_ATTRIBUTES.ROW_INDEX);
      const colAttr = current.getAttribute(TABLE_DATA_ATTRIBUTES.COLUMN_INDEX);

      if (rowAttr !== null && colAttr !== null) {
        let tableElement: Element | null = current;
        while (tableElement && tableElement !== container) {
          if (tableElement.tagName === "TABLE") {
            const tableIndexAttr = tableElement.getAttribute(TABLE_DATA_ATTRIBUTES.TABLE_INDEX);
            if (tableIndexAttr !== null) {
              return {
                tableIndex: Number.parseInt(tableIndexAttr, 10),
                rowIndex: Number.parseInt(rowAttr, 10),
                columnIndex: Number.parseInt(colAttr, 10),
              };
            }
            break;
          }
          tableElement = tableElement.parentElement;
        }
      }
      break;
    }
    current = current.parentElement;
  }

  return null;
}

/** Get a table from the document by index. */
export function getTableFromDocument(doc: Document, tableIndex: number): Table | null {
  let currentTableIndex = 0;
  for (const block of doc.package.document.content) {
    if (block.type === "table") {
      if (currentTableIndex === tableIndex) {
        return block;
      }
      currentTableIndex++;
    }
  }
  return null;
}

/** Update a table in the document immutably. */
export function updateTableInDocument(
  doc: Document,
  tableIndex: number,
  newTable: Table,
): Document {
  let currentTableIndex = 0;
  const newContent = doc.package.document.content.map((block) => {
    if (block.type === "table") {
      if (currentTableIndex === tableIndex) {
        currentTableIndex++;
        return newTable;
      }
      currentTableIndex++;
    }
    return block;
  });

  return {
    ...doc,
    package: {
      ...doc.package,
      document: {
        ...doc.package.document,
        content: newContent,
      },
    },
  };
}

/** Delete a table from the document immutably. */
export function deleteTableFromDocument(doc: Document, tableIndex: number): Document {
  let currentTableIndex = 0;
  const newContent = doc.package.document.content.filter((block) => {
    if (block.type === "table") {
      const shouldDelete = currentTableIndex === tableIndex;
      currentTableIndex++;
      return !shouldDelete;
    }
    return true;
  });

  return {
    ...doc,
    package: {
      ...doc.package,
      document: {
        ...doc.package.document,
        content: newContent,
      },
    },
  };
}

// ============================================================================
// MANAGER
// ============================================================================

/**
 * Tracked table-selection state: the selected cell coordinates plus the derived
 * {@link TableContext} (and the resolved table) for the current document.
 */
export type TableSelectionState = {
  context: TableContext | null;
  table: Table | null;
  tableIndex: number | null;
  rowIndex: number | null;
  columnIndex: number | null;
};

const EMPTY_SELECTION: TableSelectionState = {
  context: null,
  table: null,
  tableIndex: null,
  rowIndex: null,
  columnIndex: null,
};

/**
 * Outcome of dispatching a structural {@link TableAction} against a document.
 *  - `noop`: nothing to apply here (no live selection, or a border/selection
 *    action that the toolbar handles directly).
 *  - `deleted`: the table was removed and the selection cleared.
 *  - `updated`: the table was mutated; the selection now tracks the new table.
 */
export type TableActionResult =
  | { type: "noop" }
  | { type: "deleted"; document: Document }
  | { type: "updated"; document: Document; context: TableContext };

export class TableSelectionManager extends Subscribable<TableSelectionState> {
  constructor() {
    super(EMPTY_SELECTION);
  }

  /**
   * Track cell coordinates before a document model is available. This supports
   * adapter pointer state without fabricating a table context; callers that
   * need structural operations should use {@link selectCell} with a document.
   */
  selectCellCoordinates(coords: CellCoordinates): void {
    this.setSnapshot({
      context: null,
      table: null,
      tableIndex: coords.tableIndex,
      rowIndex: coords.rowIndex,
      columnIndex: coords.columnIndex,
    });
  }

  /**
   * Select a cell within `doc`. Resolves the table at `coords.tableIndex` and
   * derives the full {@link TableContext}. Returns the context, or `null` (and
   * clears the selection) when that table no longer exists.
   */
  selectCell(doc: Document, coords: CellCoordinates): TableContext | null {
    const table = getTableFromDocument(doc, coords.tableIndex);
    if (!table) {
      this.setSnapshot(EMPTY_SELECTION);
      return null;
    }
    const selection: TableSelection = {
      tableIndex: coords.tableIndex,
      rowIndex: coords.rowIndex,
      columnIndex: coords.columnIndex,
    };
    const context = createTableContext(table, selection);
    this.setSnapshot({
      context,
      table,
      tableIndex: coords.tableIndex,
      rowIndex: coords.rowIndex,
      columnIndex: coords.columnIndex,
    });
    return context;
  }

  /** Clear the current selection. */
  clearSelection(): void {
    this.setSnapshot(EMPTY_SELECTION);
  }

  /** Check if a specific cell is the current selection. */
  isCellSelected(tableIndex: number, rowIndex: number, columnIndex: number): boolean {
    const snapshot = this.getSnapshot();
    return (
      snapshot.tableIndex === tableIndex &&
      snapshot.rowIndex === rowIndex &&
      snapshot.columnIndex === columnIndex
    );
  }

  /**
   * Dispatch a structural table action against `doc`. Mutating actions
   * (add/delete row or column, merge, split, delete table) return the resulting
   * document and re-track the selection against it; border-style and selection
   * actions are routed through the toolbar's own handlers and resolve to `noop`.
   */
  handleAction(doc: Document, action: TableAction): TableActionResult {
    const { context, table, tableIndex, rowIndex, columnIndex } = this.getSnapshot();
    if (
      context === null ||
      table === null ||
      tableIndex === null ||
      rowIndex === null ||
      columnIndex === null
    ) {
      return { type: "noop" };
    }

    let newTable: Table | null = null;
    let newRowIndex = rowIndex;
    let newColumnIndex = columnIndex;

    switch (action) {
      case "addRowAbove":
        newTable = addRow(table, rowIndex, "before");
        newRowIndex = rowIndex + 1;
        break;

      case "addRowBelow":
        newTable = addRow(table, rowIndex, "after");
        break;

      case "addColumnLeft":
        newTable = addColumn(table, columnIndex, "before");
        newColumnIndex = columnIndex + 1;
        break;

      case "addColumnRight":
        newTable = addColumn(table, columnIndex, "after");
        break;

      case "deleteRow":
        if (table.rows.length > 1) {
          newTable = deleteRow(table, rowIndex);
          if (newRowIndex >= newTable.rows.length) {
            newRowIndex = newTable.rows.length - 1;
          }
        }
        break;

      case "deleteColumn": {
        const colCount = getColumnCount(table);
        if (colCount > 1) {
          newTable = deleteColumn(table, columnIndex);
          const newColCount = getColumnCount(newTable);
          if (newColumnIndex >= newColCount) {
            newColumnIndex = newColCount - 1;
          }
        }
        break;
      }

      case "mergeCells":
        if (context.selection.selectedCells) {
          newTable = mergeCells(table, context.selection);
        }
        break;

      case "splitCell":
        if (context.canSplitCell) {
          newTable = splitCell(table, rowIndex, columnIndex);
        }
        break;

      case "deleteTable": {
        const newDoc = deleteTableFromDocument(doc, tableIndex);
        this.clearSelection();
        return { type: "deleted", document: newDoc };
      }

      case "borderAll":
      case "borderBottom":
      case "borderInside":
      case "borderLeft":
      case "borderNone":
      case "borderOutside":
      case "borderRight":
      case "borderTop":
      case "selectColumn":
      case "selectRow":
      case "selectTable":
        // Border-style and selection actions are routed through the toolbar's
        // border/selection handlers — they don't modify table structure, so the
        // dispatcher above has nothing to do.
        break;
    }

    if (!newTable) {
      return { type: "noop" };
    }

    const newDoc = updateTableInDocument(doc, tableIndex, newTable);
    // Re-resolve the selection against the mutated document so the tracked
    // context reflects the new table shape rather than the pre-mutation one.
    const nextContext = this.selectCell(newDoc, {
      tableIndex,
      rowIndex: newRowIndex,
      columnIndex: newColumnIndex,
    });
    if (!nextContext) {
      return { type: "noop" };
    }
    return { type: "updated", document: newDoc, context: nextContext };
  }
}
