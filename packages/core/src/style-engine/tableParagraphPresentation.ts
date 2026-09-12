import type {
  ParagraphFormatting,
  Table,
  TableCell,
  TableRow,
  TableLook,
  TextFormatting,
} from "../types/document";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import type { StyleEngine } from "./styleEngine";
import { resolveRunFormattingWithoutDefaults } from "./runPresentation";
import {
  mergeTableParagraphPresentations,
  projectTableParagraphPresentation,
  type TableParagraphPresentationProjection,
} from "./paragraphPresentation";

const TABLE_CONDITIONAL_REGIONS = Object.freeze([
  "wholeTable",
  "firstRow",
  "lastRow",
  "firstCol",
  "lastCol",
  "band1Horz",
  "band2Horz",
  "band1Vert",
  "band2Vert",
  "nwCell",
  "neCell",
  "swCell",
  "seCell",
] as const);

type TableConditionalRegion = (typeof TABLE_CONDITIONAL_REGIONS)[number];

export type TableCellPresentationProjection = {
  readonly paragraph?: TableParagraphPresentationProjection;
  readonly runFormatting?: TextFormatting;
};

type TableParagraphConditionalStyles = Partial<
  Record<TableConditionalRegion, TableCellPresentationProjection>
>;

const projectConditionalPresentation = (
  pPr: ParagraphFormatting | undefined,
  rPr: TextFormatting | undefined,
  styleResolver: Pick<StyleEngine, "getRunStyleOwnProperties">,
): TableCellPresentationProjection | undefined => {
  const paragraph = projectTableParagraphPresentation(pPr);
  const runFromParagraph = resolveRunFormattingWithoutDefaults(
    pPr?.runProperties,
    styleResolver,
  );
  const runFormatting = mergeTextFormatting(
    runFromParagraph,
    resolveRunFormattingWithoutDefaults(rPr, styleResolver),
  );
  return paragraph || runFormatting ? { paragraph, runFormatting } : undefined;
};

const mergeConditionalPresentation = (
  base: TableCellPresentationProjection | undefined,
  override: TableCellPresentationProjection | undefined,
): TableCellPresentationProjection | undefined => {
  if (!base) return override;
  if (!override) return base;
  const paragraph = mergeTableParagraphPresentations(base.paragraph, override.paragraph);
  const runFormatting = mergeTextFormatting(base.runFormatting, override.runFormatting);
  return paragraph || runFormatting ? { paragraph, runFormatting } : undefined;
};

const countTableColumns = (rows: readonly TableRow[]): number => {
  let maximum = 0;
  for (const row of rows) {
    let columns = row.formatting?.gridBefore ?? 0;
    for (const cell of row.cells) columns += cell.formatting?.gridSpan ?? 1;
    columns += row.formatting?.gridAfter ?? 0;
    maximum = Math.max(maximum, columns);
  }
  return maximum;
};

const resolveConditionalStyles = (
  table: Table,
  styleResolver: Pick<
    StyleEngine,
    "getDefaultTableStyle" | "getStyle" | "getRunStyleOwnProperties"
  > | null,
): TableParagraphConditionalStyles => {
  if (!styleResolver) return {};
  const explicitStyleId = table.formatting?.styleId;
  const style = explicitStyleId
    ? styleResolver.getStyle(explicitStyleId)
    : styleResolver.getDefaultTableStyle();
  if (!style) return {};
  const projections: TableParagraphConditionalStyles = {};
  const base = projectConditionalPresentation(style.pPr, style.rPr, styleResolver);
  const wholeTableStyle = style.tblStylePr?.find(({ type }) => type === "wholeTable");
  const wholeTable = projectConditionalPresentation(
    wholeTableStyle?.pPr,
    wholeTableStyle?.rPr,
    styleResolver,
  );
  const mergedWholeTable = mergeConditionalPresentation(base, wholeTable);
  if (mergedWholeTable) projections.wholeTable = mergedWholeTable;
  for (const region of TABLE_CONDITIONAL_REGIONS) {
    if (region === "wholeTable") continue;
    const conditional = style.tblStylePr?.find(({ type }) => type === region);
    const projection = projectConditionalPresentation(
      conditional?.pPr,
      conditional?.rPr,
      styleResolver,
    );
    if (projection) projections[region] = projection;
  }
  return projections;
};

const rowBandIndex = (table: Table, rowIndex: number, look: TableLook | undefined): number => {
  let index = 0;
  for (let candidate = 0; candidate < rowIndex; candidate++) {
    const isStyledFirst = candidate === 0 && look?.firstRow === true;
    const isStyledLast = candidate === table.rows.length - 1 && look?.lastRow === true;
    if (!isStyledFirst && !isStyledLast) index++;
  }
  return index;
};

const gridColumnStart = (row: TableRow, cellIndex: number): number => {
  let column = row.formatting?.gridBefore ?? 0;
  for (let index = 0; index < cellIndex; index++) {
    column += row.cells[index]?.formatting?.gridSpan ?? 1;
  }
  return column;
};

export type ResolveTableCellParagraphPresentationOptions = {
  readonly row: TableRow;
  readonly rowIndex: number;
  readonly cell: TableCell;
  readonly cellIndex: number;
};

/**
 * One immutable resolver for the table-style paragraph layer used by import
 * and by live comparison projection. The conditional precedence is computed
 * from the authored table model on every projection; no effective state is
 * cached on ProseMirror nodes.
 */
export const createTableCellPresentationResolver = ({
  table,
  styleResolver,
}: {
  readonly table: Table;
  readonly styleResolver: Pick<
    StyleEngine,
    "getDefaultTableStyle" | "getStyle" | "getRunStyleOwnProperties"
  > | null;
}): ((
  options: ResolveTableCellParagraphPresentationOptions,
) => TableCellPresentationProjection | undefined) => {
  const styles = resolveConditionalStyles(table, styleResolver);
  const look = table.formatting?.look;
  const totalColumns =
    (table.columnWidths?.length ?? 0) > 0
      ? (table.columnWidths?.length ?? 0)
      : Math.max(countTableColumns(table.rows), 1);
  const horizontalBanding = look?.noHBand !== true;
  const verticalBanding = look?.noVBand !== true;

  return ({ row, rowIndex, cell, cellIndex }) => {
    const rowCnf = row.formatting?.conditionalFormat;
    const cellCnf = cell.formatting?.conditionalFormat;
    const columnStart = gridColumnStart(row, cellIndex);
    const columnEnd = columnStart + (cell.formatting?.gridSpan ?? 1);
    const isFirstRow = rowIndex === 0;
    const isLastRow = rowIndex === table.rows.length - 1;
    const isFirstColumn = columnStart === 0;
    const isLastColumn = columnEnd === totalColumns;
    const cellIsFirstRow = cellCnf?.firstRow ?? rowCnf?.firstRow ?? isFirstRow;
    const cellIsLastRow = cellCnf?.lastRow ?? rowCnf?.lastRow ?? isLastRow;
    const cellIsFirstColumn = cellCnf?.firstColumn ?? isFirstColumn;
    const cellIsLastColumn = cellCnf?.lastColumn ?? isLastColumn;

    let projection = styles.wholeTable;
    let horizontalBand:
      | TableCellPresentationProjection
      | undefined;
    if (
      horizontalBanding &&
      !(isFirstRow && look?.firstRow === true) &&
      !(isLastRow && look?.lastRow === true)
    ) {
      horizontalBand =
        rowBandIndex(table, rowIndex, look) % 2 === 0 ? styles.band1Horz : styles.band2Horz;
    }
    if (rowCnf?.oddHBand || cellCnf?.oddHBand) horizontalBand = styles.band1Horz;
    if (rowCnf?.evenHBand || cellCnf?.evenHBand) horizontalBand = styles.band2Horz;
    projection = mergeConditionalPresentation(projection, horizontalBand);

    let verticalBand: TableCellPresentationProjection | undefined;
    if (verticalBanding) {
      const firstColumnOffset = look?.firstColumn ? 1 : 0;
      const bandColumn = columnStart - firstColumnOffset;
      if (
        bandColumn >= 0 &&
        !(look?.lastColumn && cellIsLastColumn) &&
        !(look?.firstColumn && cellIsFirstColumn)
      ) {
        verticalBand = bandColumn % 2 === 0 ? styles.band1Vert : styles.band2Vert;
      }
    }
    if (cellCnf?.oddVBand) verticalBand = styles.band1Vert;
    if (cellCnf?.evenVBand) verticalBand = styles.band2Vert;
    projection = mergeConditionalPresentation(projection, verticalBand);

    if (cellIsFirstColumn && (look?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)) {
      projection = mergeConditionalPresentation(projection, styles.firstCol);
    }
    if (cellIsLastColumn && (look?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)) {
      projection = mergeConditionalPresentation(projection, styles.lastCol);
    }
    if (cellIsFirstRow && (look?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow)) {
      projection = mergeConditionalPresentation(projection, styles.firstRow);
    }
    if (cellIsLastRow && (look?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow)) {
      projection = mergeConditionalPresentation(projection, styles.lastRow);
    }
    if (
      cellIsFirstRow &&
      cellIsFirstColumn &&
      (look?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow) &&
      (look?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)
    ) {
      projection = mergeConditionalPresentation(projection, styles.nwCell);
    }
    if (
      cellIsFirstRow &&
      cellIsLastColumn &&
      (look?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow) &&
      (look?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)
    ) {
      projection = mergeConditionalPresentation(projection, styles.neCell);
    }
    if (
      cellIsLastRow &&
      cellIsFirstColumn &&
      (look?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow) &&
      (look?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn)
    ) {
      projection = mergeConditionalPresentation(projection, styles.swCell);
    }
    if (
      cellIsLastRow &&
      cellIsLastColumn &&
      (look?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow) &&
      (look?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn)
    ) {
      projection = mergeConditionalPresentation(projection, styles.seCell);
    }
    return projection;
  };
};
