/**
 * The `w:tblGrid` a table's model writes when no capture stands in for it.
 * The serializer writes it; the parser keeps a capture only when the source
 * grid differs from it, so a table without one reads back as it was written.
 */
import type { Table, TableGridChange } from "../types/document";
import { intAttr } from "./serializer/xmlUtils";

/**
 * The most grid columns a table declares. Far above what any editor lays out,
 * it bounds a grid built from a source's skipped-column
 * and span counts, which a file can state at any size.
 */
export const MAX_GRID_COLUMNS = 4096;

/**
 * Columns the grid has to declare: the widest row's total of skipped columns
 * and spans, because a grid narrower than a row leaves cells with no column
 * to sit in. At most {@link MAX_GRID_COLUMNS}.
 */
export const gridColumnCount = (table: Table): number => {
  let widest = 0;
  for (const row of table.rows) {
    let columns =
      (row.formatting?.gridBefore ?? 0) + (row.formatting?.gridAfter ?? 0);
    for (const cell of row.cells) {
      columns += cell.formatting?.gridSpan ?? 1;
    }
    widest = Math.max(widest, columns);
  }
  return Math.min(widest, MAX_GRID_COLUMNS);
};

/**
 * Serialize `w:tblGridChange`: the grid a reviewer replaced, from the change's
 * own record rather than from the bytes it arrived as.
 *
 * `w:w` is optional on a `w:gridCol`, so a column the snapshot recorded no
 * measure for is written back without one; spelling it as a zero would claim a
 * width the source never stated.
 */
const serializeTableGridChange = ({
  id,
  columnWidths,
}: TableGridChange): string => {
  const columns = columnWidths
    .map((width) =>
      width === undefined
        ? "<w:gridCol/>"
        : `<w:gridCol w:w="${intAttr(width)}"/>`,
    )
    .join("");
  return `<w:tblGridChange w:id="${intAttr(id)}"><w:tblGrid>${columns}</w:tblGrid></w:tblGridChange>`;
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
  const columns = gridColumnCount(table);
  return columns === 0 && gridChange === ""
    ? "<w:tblGrid/>"
    : `<w:tblGrid>${"<w:gridCol/>".repeat(columns)}${gridChange}</w:tblGrid>`;
};
