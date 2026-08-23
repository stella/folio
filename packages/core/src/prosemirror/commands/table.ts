/**
 * ProseMirror Table Commands — thin re-exports from extension system
 *
 * Table context detection, insert/delete operations, borders, cell styling.
 * All implementations live in extensions/nodes/TableExtension.ts; this file
 * re-exports for backward compatibility.
 */

import type { EditorState, Transaction } from "prosemirror-state";

import type { BorderPreset, TableBorderPreset } from "../extensions/nodes/TableExtension";
import type {
  TableBorderCommandSpec,
  TableCellBorderCommandSpec,
  TableCellMarginsCommand,
  TablePropertiesCommand,
  TableStyleCommand,
} from "../extensions/types";
import { singletonManager } from "../schema";

// Re-export types and query helpers from TableExtension
export type {
  TableContextInfo,
  BorderPreset,
  TableBorderPreset,
} from "../extensions/nodes/TableExtension";
export { getTableContext, isInTable } from "../extensions/nodes/TableExtension";

// ============================================================================
// COMMANDS — delegated to singleton extension manager
// ============================================================================

// The singleton validates the complete built-in registry during startup.
const cmds = singletonManager;

// Table creation
export function insertTable(
  rows: number,
  cols: number,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("insertTable")(rows, cols);
}

// Row operations
export function addRowAbove(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("addRowAbove")()(state, dispatch);
}
export function addRowBelow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("addRowBelow")()(state, dispatch);
}
export function deleteRow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("deleteRow")()(state, dispatch);
}

// Column operations
export function addColumnLeft(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("addColumnLeft")()(state, dispatch);
}
export function addColumnRight(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("addColumnRight")()(state, dispatch);
}
export function deleteColumn(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("deleteColumn")()(state, dispatch);
}

// Table deletion
export function deleteTable(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("deleteTable")()(state, dispatch);
}

// Table selection
export function selectTable(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("selectTable")()(state, dispatch);
}
export function selectRow(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("selectRow")()(state, dispatch);
}
export function selectColumn(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("selectColumn")()(state, dispatch);
}

// Merge/Split — delegated to prosemirror-tables via singleton extension manager
export function mergeCells(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("mergeCells")()(state, dispatch);
}
export function splitCell(state: EditorState, dispatch?: (tr: Transaction) => void): boolean {
  return cmds.requireCommand("splitCell")()(state, dispatch);
}

// Per-cell border editing
export function setCellBorder(
  side: "top" | "bottom" | "left" | "right" | "all",
  spec: TableCellBorderCommandSpec | null,
  clearOthers?: boolean,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setCellBorder")(side, spec, clearOthers);
}

// Whole-table border presets (applied to the table under the caret)
export function setTableBorderPreset(
  preset: TableBorderPreset,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setTableBorderPreset")(preset);
}

// Borders
export function setTableBorders(
  preset: BorderPreset,
  borderSpec?: TableBorderCommandSpec,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setTableBorders")(preset, borderSpec);
}
export function removeTableBorders(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
): boolean {
  return cmds.requireCommand("removeTableBorders")()(state, dispatch);
}
export function setAllTableBorders(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  borderSpec?: TableBorderCommandSpec,
): boolean {
  return cmds.requireCommand("setAllTableBorders")(borderSpec)(state, dispatch);
}
export function setOutsideTableBorders(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  borderSpec?: TableBorderCommandSpec,
): boolean {
  return cmds.requireCommand("setOutsideTableBorders")(borderSpec)(state, dispatch);
}
export function setInsideTableBorders(
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
  borderSpec?: TableBorderCommandSpec,
): boolean {
  return cmds.requireCommand("setInsideTableBorders")(borderSpec)(state, dispatch);
}

// Vertical alignment
export function setCellVerticalAlign(
  align: "top" | "center" | "bottom",
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setCellVerticalAlign")(align);
}

// Cell margins
export function setCellMargins(
  margins: TableCellMarginsCommand,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setCellMargins")(margins);
}

// Text direction
export function setCellTextDirection(
  direction: string | null,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setCellTextDirection")(direction);
}

// No-wrap toggle
export function toggleNoWrap(): (
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
) => boolean {
  return cmds.requireCommand("toggleNoWrap")();
}

// Row height
export function setRowHeight(
  height: number | null,
  rule?: "auto" | "atLeast" | "exact",
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setRowHeight")(height, rule);
}

// Header row
export function toggleHeaderRow(): (
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
) => boolean {
  return cmds.requireCommand("toggleHeaderRow")();
}

// Column distribution
export function distributeColumns(): (
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
) => boolean {
  return cmds.requireCommand("distributeColumns")();
}
export function autoFitContents(): (
  state: EditorState,
  dispatch?: (tr: Transaction) => void,
) => boolean {
  return cmds.requireCommand("autoFitContents")();
}

// Table properties
export function setTableProperties(
  props: TablePropertiesCommand,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setTableProperties")(props);
}

// Table style gallery
export function applyTableStyle(
  styleData: TableStyleCommand,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("applyTableStyle")(styleData);
}

// Cell styling
export function setCellFillColor(
  color: string | null,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setCellFillColor")(color);
}
export function setTableBorderColor(
  color: string,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setTableBorderColor")(color);
}
export function setTableBorderWidth(
  size: number,
): (state: EditorState, dispatch?: (tr: Transaction) => void) => boolean {
  return cmds.requireCommand("setTableBorderWidth")(size);
}
