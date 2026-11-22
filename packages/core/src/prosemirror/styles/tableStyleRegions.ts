/**
 * Which regions of a table style (`w:tblStylePr`) apply to a cell.
 *
 * The load path layers each cell's conditional formatting from these regions,
 * and the live editor resolves a paragraph an edit creates inside a cell from
 * the same list, so both read one cell the same way.
 */

import type { ResolvedTableLook } from "../../docx/tableLook";
import type { ConditionalFormatStyle } from "../../types";

export type TableStyleRegion =
  | "wholeTable"
  | "band1Horz"
  | "band2Horz"
  | "band1Vert"
  | "band2Vert"
  | "firstCol"
  | "lastCol"
  | "firstRow"
  | "lastRow"
  | "nwCell"
  | "neCell"
  | "swCell"
  | "seCell";

type TableStyleHeaderOptions = {
  look: ResolvedTableLook;
  totalRows: number;
  repeats: (rowIndex: number) => boolean;
};

/**
 * The first-row style covers row zero and any consecutive repeating rows
 * that follow it. This region does not control which rows repeat on pages.
 */
export const tableStyleHeaderRows = ({
  look,
  totalRows,
  repeats,
}: TableStyleHeaderOptions): number => {
  if (!look.firstRow || totalRows === 0) return 0;
  let headerRows = 1;
  while (headerRows < totalRows && repeats(headerRows)) headerRows += 1;
  return headerRows;
};

export type TableCellPlacement = {
  look: ResolvedTableLook;
  rowIndex: number;
  totalRows: number;
  /** Number of leading rows in the first-row style region. */
  headerRows: number;
  /** The first grid column the cell covers. */
  column: number;
  colspan: number;
  totalColumns: number;
  rowConditionalFormat?: ConditionalFormatStyle | undefined;
  cellConditionalFormat?: ConditionalFormatStyle | undefined;
};

/**
 * The regions that apply to one cell, lowest precedence first: the whole
 * table, banding, columns, rows, then corners (ECMA-376 §17.7.6). A cell's or
 * row's `w:cnfStyle` overrides what its position implies.
 */
export const tableCellStyleRegions = ({
  look,
  rowIndex,
  totalRows,
  headerRows,
  column,
  colspan,
  totalColumns,
  rowConditionalFormat: rowCnf,
  cellConditionalFormat: cellCnf,
}: TableCellPlacement): TableStyleRegion[] => {
  const lastRow = rowIndex === totalRows - 1;
  const rowIsFirstRow = rowCnf?.firstRow ?? rowIndex < headerRows;
  const rowIsLastRow = rowCnf?.lastRow ?? lastRow;
  const isFirstRow = cellCnf?.firstRow ?? rowIsFirstRow;
  const isLastRow = cellCnf?.lastRow ?? rowIsLastRow;
  const isFirstCol = cellCnf?.firstColumn ?? column === 0;
  const isLastCol = cellCnf?.lastColumn ?? column + colspan === totalColumns;

  let verticalBand: TableStyleRegion | undefined;
  if (!look.noVBand) {
    const bandColumn = column - (look.firstColumn ? 1 : 0);
    const eligible =
      bandColumn >= 0 &&
      !(look.lastColumn && isLastCol) &&
      !(look.firstColumn && isFirstCol);
    if (eligible) {
      verticalBand = bandColumn % 2 === 0 ? "band1Vert" : "band2Vert";
    }
  }
  if (cellCnf?.oddVBand) {
    verticalBand = "band1Vert";
  } else if (cellCnf?.evenVBand) {
    verticalBand = "band2Vert";
  }

  let horizontalBand: TableStyleRegion | undefined;
  if (!look.noHBand && rowIndex >= headerRows && !(look.lastRow && lastRow)) {
    horizontalBand =
      (rowIndex - headerRows) % 2 === 0 ? "band1Horz" : "band2Horz";
  }
  if (rowCnf?.oddHBand) {
    horizontalBand = "band1Horz";
  } else if (rowCnf?.evenHBand) {
    horizontalBand = "band2Horz";
  }
  if (cellCnf?.oddHBand) {
    horizontalBand = "band1Horz";
  } else if (cellCnf?.evenHBand) {
    horizontalBand = "band2Horz";
  }

  const firstRowOn =
    look.firstRow || rowCnf?.firstRow === true || cellCnf?.firstRow === true;
  const lastRowOn =
    look.lastRow || rowCnf?.lastRow === true || cellCnf?.lastRow === true;
  const firstColOn =
    look.firstColumn ||
    rowCnf?.firstColumn === true ||
    cellCnf?.firstColumn === true;
  const lastColOn =
    look.lastColumn ||
    rowCnf?.lastColumn === true ||
    cellCnf?.lastColumn === true;

  const regions: TableStyleRegion[] = ["wholeTable"];
  if (horizontalBand) {
    regions.push(horizontalBand);
  }
  if (verticalBand) {
    regions.push(verticalBand);
  }
  if (isFirstCol && firstColOn) {
    regions.push("firstCol");
  }
  if (isLastCol && lastColOn) {
    regions.push("lastCol");
  }
  if (isFirstRow && firstRowOn) {
    regions.push("firstRow");
  }
  if (isLastRow && lastRowOn) {
    regions.push("lastRow");
  }
  if (isFirstRow && isFirstCol && firstRowOn && firstColOn) {
    regions.push("nwCell");
  }
  if (isFirstRow && isLastCol && firstRowOn && lastColOn) {
    regions.push("neCell");
  }
  if (isLastRow && isFirstCol && lastRowOn && firstColOn) {
    regions.push("swCell");
  }
  if (isLastRow && isLastCol && lastRowOn && lastColOn) {
    regions.push("seCell");
  }
  return regions;
};
