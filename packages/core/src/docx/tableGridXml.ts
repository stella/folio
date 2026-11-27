/**
 * The `w:tblGrid` a table's model writes when no capture stands in for it.
 * The serializer writes it; the parser keeps a capture only when the source
 * grid differs from it, so a table without one reads back as it was written.
 */
import {
  gridColumnsInRows,
  MAX_TABLE_GRID_COLUMNS as MAX_GRID_COLUMNS,
} from "@stll/docx-core/model";

import type { Table, TableGridChange } from "../types/document";
import { serializeTableGeometryChange } from "./reviewResolutionProvenance";
import { intAttr } from "./serializer/xmlUtils";

/**
 * The most grid columns a table declares. Far above what any editor lays out,
 * it bounds a grid built from a source's skipped-column
 * and span counts, which a file can state at any size.
 */
export { MAX_GRID_COLUMNS };

/**
 * Columns the grid has to declare: the widest row's total of skipped columns
 * and spans, because a grid narrower than a row leaves cells with no column
 * to sit in. At most {@link MAX_GRID_COLUMNS}.
 */
export const gridColumnCount = (table: Table): number =>
  Math.min(gridColumnsInRows(table.rows), MAX_GRID_COLUMNS);

const currentGridColumnCount = (
  table: Table,
  change: TableGridChange | undefined,
): number => {
  const geometry = change?.geometryChange;
  if (change === undefined || geometry === undefined)
    return gridColumnCount(table);
  if (geometry.type === "mergeCells") return change.columnWidths.length;
  // The pending physical table still has the deleted cells. Its current
  // grid covers the surviving cells with their current spans and omissions.
  const ids = new Set(geometry.revisionIds);
  const rows = table.rows.map((row) => ({
    ...row,
    cells: row.cells.filter(
      ({ structuralChange }) =>
        structuralChange?.type !== "tableCellDeletion" ||
        !ids.has(structuralChange.info.id),
    ),
  }));
  return gridColumnCount({ ...table, rows });
};

/**
 * Serialize `w:tblGridChange`: the grid a reviewer replaced, from the change's
 * own record rather than from the bytes it arrived as.
 *
 * `w:w` is optional on a `w:gridCol`, so a column the snapshot recorded no
 * measure for is written back without one; spelling it as a zero would claim a
 * width the source never stated.
 */
const serializeTableGridChange = (change: TableGridChange): string => {
  const { id, columnWidths } = change;
  const columns = columnWidths
    .map((width) =>
      width === undefined
        ? "<w:gridCol/>"
        : `<w:gridCol w:w="${intAttr(width)}"/>`,
    )
    .join("");
  return `<w:tblGridChange w:id="${intAttr(id)}"${serializeTableGeometryChange(change)}><w:tblGrid>${columns}</w:tblGrid></w:tblGridChange>`;
};

/**
 * `w:tblGrid` from the model.
 *
 * `w:tblGrid` is required on every `w:tbl` and `w:w` is optional on a
 * `w:gridCol`, so a table whose column widths were never measured (one the
 * comparison creates, say) still declares a grid, just without widths.
 * `w:tblGridChange` records the grid a reviewer replaced. It is history, so
 * no width in the model derives it, and a rebuilt grid that leaves it out
 * accepts the reviewer's change without saying so. The schema declares it
 * after every `w:gridCol`, so it is appended.
 */
export const modelTableGrid = (table: Table): string => {
  const change = table.formatting?.gridChange;
  const gridChange =
    change === undefined ? "" : serializeTableGridChange(change);
  const columnWidths = table.columnWidths;
  if (columnWidths && columnWidths.length > 0) {
    const columns = columnWidths
      .map((w) => `<w:gridCol w:w="${intAttr(w)}"/>`)
      .join("");
    return `<w:tblGrid>${columns}${gridChange}</w:tblGrid>`;
  }
  const columns = currentGridColumnCount(table, change);
  return columns === 0 && gridChange === ""
    ? "<w:tblGrid/>"
    : `<w:tblGrid>${"<w:gridCol/>".repeat(columns)}${gridChange}</w:tblGrid>`;
};
