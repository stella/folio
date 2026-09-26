import type { ListNumberingMenuState } from "../prosemirror/listNumbering";

// ============================================================================
// TABLE SELECTION
// ============================================================================

/** Cell coordinates in a table */
export type CellCoordinates = {
  tableIndex: number;
  rowIndex: number;
  columnIndex: number;
};

// ============================================================================
// CONTEXT MENU
// ============================================================================

/** Viewport anchor (px) where the context menu opens. */
export type ContextMenuAnchor = { x: number; y: number };

/** ContextMenuManager snapshot. */
export type ContextMenuSnapshot = {
  isOpen: boolean;
  position: ContextMenuAnchor;
  hasSelection: boolean;
  selectionRange: { from: number; to: number };
  cursorInTable: boolean;
  cursorInTrackedChange: boolean;
  /** Restart / continue / set-value entries for the list item under the caret. */
  listNumbering: ListNumberingMenuState;
};

// ============================================================================
// ERROR MANAGER
// ============================================================================

/** Error severity levels */
export type ErrorSeverity = "error" | "warning" | "info";

/** Error notification */
export type ErrorNotification = {
  id: string;
  message: string;
  severity: ErrorSeverity;
  details?: string;
  timestamp: number;
  dismissed?: boolean;
};

/** ErrorManager snapshot */
export type ErrorManagerSnapshot = {
  notifications: ErrorNotification[];
};
