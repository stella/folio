import { panic } from "better-result";
import type { Node as PMNode } from "prosemirror-model";

import type { StyleEngine, TableCellParagraphSpacingOverlay } from "../style-engine";
import type {
  ParagraphFormatting,
  TableCellFormatting,
  TableLook,
  TableRowFormatting,
  TextFormatting,
} from "../types/document";
import { mergeParagraphFormatting } from "../utils/paragraphFormattingMerge";
import { mergeTextFormatting } from "../utils/textFormattingMerge";
import { expectTableAttrs, expectTableCellAttrs, expectTableRowAttrs } from "./attrs";
import { resolveRunFormattingWithoutDefaults } from "./runStyleFormatting";

type TableConditionalStyleResolver = Pick<
  StyleEngine,
  "getDefaultTableStyle" | "getRunStyleOwnProperties" | "getStyle"
>;

const TABLE_CONDITIONAL_STYLE_TYPES = [
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
] as const;

type TableConditionalStyleType = (typeof TABLE_CONDITIONAL_STYLE_TYPES)[number];

export type TableConditionalStyle = {
  tcPr?: TableCellFormatting;
  rPr?: TextFormatting;
  pPr?: TableCellParagraphSpacingOverlay;
};

export type TableConditionalStyles = Partial<
  Record<TableConditionalStyleType, TableConditionalStyle>
>;

const extractTableParagraphOverlay = (
  pPr: ParagraphFormatting | undefined,
): TableCellParagraphSpacingOverlay | undefined => {
  if (!pPr) {
    return undefined;
  }
  const overlay: TableCellParagraphSpacingOverlay = {};
  for (const key of [
    "spaceBefore",
    "spaceAfter",
    "lineSpacing",
    "lineSpacingRule",
    "contextualSpacing",
    "frame",
  ] as const) {
    const value = pPr[key];
    if (value !== undefined) {
      Object.assign(overlay, { [key]: value });
    }
  }
  return Object.keys(overlay).length > 0 ? overlay : undefined;
};

const resolveTableStyleConditional = (
  styleResolver: TableConditionalStyleResolver,
  tableStyleId: string,
  conditionType: TableConditionalStyleType,
): TableConditionalStyle | undefined => {
  const conditional = styleResolver
    .getStyle(tableStyleId)
    ?.tblStylePr?.find(({ type }) => type === conditionType);
  if (!conditional) {
    return undefined;
  }

  const runPropsFromPpr = resolveRunFormattingWithoutDefaults(
    conditional.pPr?.runProperties,
    styleResolver,
  );
  const resolvedRpr = resolveRunFormattingWithoutDefaults(conditional.rPr, styleResolver);
  const result: TableConditionalStyle = {};
  const rPr = mergeTextFormatting(runPropsFromPpr, resolvedRpr);
  const pPr = extractTableParagraphOverlay(conditional.pPr);
  if (conditional.tcPr) {
    result.tcPr = conditional.tcPr;
  }
  if (rPr) {
    result.rPr = rPr;
  }
  if (pPr) {
    result.pPr = pPr;
  }
  return result;
};

const resolveTableBaseStyle = (
  styleResolver: TableConditionalStyleResolver,
  tableStyleId: string,
): TableConditionalStyle | undefined => {
  const style = styleResolver.getStyle(tableStyleId);
  if (!style) {
    return undefined;
  }
  const runPropsFromPpr = resolveRunFormattingWithoutDefaults(
    style.pPr?.runProperties,
    styleResolver,
  );
  const resolvedRpr = resolveRunFormattingWithoutDefaults(style.rPr, styleResolver);
  const result: TableConditionalStyle = {};
  const rPr = mergeTextFormatting(runPropsFromPpr, resolvedRpr);
  const pPr = extractTableParagraphOverlay(style.pPr);
  if (style.tcPr) {
    result.tcPr = style.tcPr;
  }
  if (rPr) {
    result.rPr = rPr;
  }
  if (pPr) {
    result.pPr = pPr;
  }
  return result.tcPr || result.rPr || result.pPr ? result : undefined;
};

export const mergeTableConditionalStyles = (
  base?: TableConditionalStyle,
  override?: TableConditionalStyle,
): TableConditionalStyle | undefined => {
  if (!base) {
    return override;
  }
  if (!override) {
    return base;
  }

  const merged: TableConditionalStyle = {};
  if (base.tcPr || override.tcPr) {
    merged.tcPr = {
      ...base.tcPr,
      ...override.tcPr,
      ...(base.tcPr?.borders || override.tcPr?.borders
        ? { borders: { ...base.tcPr?.borders, ...override.tcPr?.borders } }
        : {}),
      ...(base.tcPr?.shading || override.tcPr?.shading
        ? { shading: { ...base.tcPr?.shading, ...override.tcPr?.shading } }
        : {}),
      ...(base.tcPr?.margins || override.tcPr?.margins
        ? { margins: { ...base.tcPr?.margins, ...override.tcPr?.margins } }
        : {}),
    };
  }
  const rPr = mergeTextFormatting(base.rPr, override.rPr);
  if (rPr) {
    merged.rPr = rPr;
  }
  const pPr = mergeParagraphFormatting(base.pPr, override.pPr);
  if (pPr) {
    merged.pPr = pPr;
  }
  return merged;
};

export const resolveTableConditionalStyles = (
  styleResolver: TableConditionalStyleResolver | null,
  tableStyleId: string | undefined,
): TableConditionalStyles => {
  if (!styleResolver) {
    return {};
  }
  const style = tableStyleId
    ? styleResolver.getStyle(tableStyleId)
    : styleResolver.getDefaultTableStyle();
  if (!style) {
    return {};
  }

  const styles: TableConditionalStyles = {};
  for (const type of TABLE_CONDITIONAL_STYLE_TYPES) {
    const conditional = resolveTableStyleConditional(styleResolver, style.styleId, type);
    if (conditional) {
      styles[type] = conditional;
    }
  }
  const wholeTable = mergeTableConditionalStyles(
    resolveTableBaseStyle(styleResolver, style.styleId),
    styles.wholeTable,
  );
  if (wholeTable) {
    styles.wholeTable = wholeTable;
  }
  return styles;
};

type ResolveTableCellConditionalStyleOptions = {
  styles: TableConditionalStyles;
  look: TableLook | undefined;
  rowIndex: number;
  totalRows: number;
  columnIndex: number;
  columnSpan: number;
  totalColumns: number;
  rowFormatting: Pick<TableRowFormatting, "conditionalFormat"> | undefined;
  cellFormatting: Pick<TableCellFormatting, "conditionalFormat"> | undefined;
};

export const resolveTableCellConditionalStyle = ({
  styles,
  look,
  rowIndex,
  totalRows,
  columnIndex,
  columnSpan,
  totalColumns,
  rowFormatting,
  cellFormatting,
}: ResolveTableCellConditionalStyleOptions): TableConditionalStyle | undefined => {
  const rowCnf = rowFormatting?.conditionalFormat;
  const cellCnf = cellFormatting?.conditionalFormat;
  const isFirstRow = rowIndex === 0;
  const isLastRow = rowIndex === totalRows - 1;
  const isFirstColumn = columnIndex === 0;
  const isLastColumn = columnIndex + columnSpan === totalColumns;
  const cellIsFirstRow = cellCnf?.firstRow ?? rowCnf?.firstRow ?? isFirstRow;
  const cellIsLastRow = cellCnf?.lastRow ?? rowCnf?.lastRow ?? isLastRow;
  const cellIsFirstColumn = cellCnf?.firstColumn ?? isFirstColumn;
  const cellIsLastColumn = cellCnf?.lastColumn ?? isLastColumn;

  let rowBand: TableConditionalStyle | undefined;
  const firstRowStyled = isFirstRow && look?.firstRow === true;
  const lastRowStyled = isLastRow && look?.lastRow === true;
  if (look?.noHBand !== true && !firstRowStyled && !lastRowStyled) {
    const bandIndex = rowIndex - (look?.firstRow === true ? 1 : 0);
    rowBand = bandIndex % 2 === 0 ? styles.band1Horz : styles.band2Horz;
  }
  if (rowCnf?.oddHBand) {
    rowBand = styles.band1Horz;
  } else if (rowCnf?.evenHBand) {
    rowBand = styles.band2Horz;
  }
  if (cellCnf?.oddHBand) {
    rowBand = styles.band1Horz;
  } else if (cellCnf?.evenHBand) {
    rowBand = styles.band2Horz;
  }

  let columnBand: TableConditionalStyle | undefined;
  if (look?.noVBand !== true) {
    const bandIndex = columnIndex - (look?.firstColumn === true ? 1 : 0);
    const isEligible =
      bandIndex >= 0 &&
      !(look?.firstColumn && cellIsFirstColumn) &&
      !(look?.lastColumn && cellIsLastColumn);
    if (isEligible) {
      columnBand = bandIndex % 2 === 0 ? styles.band1Vert : styles.band2Vert;
    }
  }
  if (cellCnf?.oddVBand) {
    columnBand = styles.band1Vert;
  } else if (cellCnf?.evenVBand) {
    columnBand = styles.band2Vert;
  }

  const firstRowActive =
    cellIsFirstRow && !!(look?.firstRow || rowCnf?.firstRow || cellCnf?.firstRow);
  const lastRowActive = cellIsLastRow && !!(look?.lastRow || rowCnf?.lastRow || cellCnf?.lastRow);
  const firstColumnActive =
    cellIsFirstColumn && !!(look?.firstColumn || rowCnf?.firstColumn || cellCnf?.firstColumn);
  const lastColumnActive =
    cellIsLastColumn && !!(look?.lastColumn || rowCnf?.lastColumn || cellCnf?.lastColumn);

  let result = styles.wholeTable;
  result = mergeTableConditionalStyles(result, rowBand);
  result = mergeTableConditionalStyles(result, columnBand);
  if (firstColumnActive) {
    result = mergeTableConditionalStyles(result, styles.firstCol);
  }
  if (lastColumnActive) {
    result = mergeTableConditionalStyles(result, styles.lastCol);
  }
  if (firstRowActive) {
    result = mergeTableConditionalStyles(result, styles.firstRow);
  }
  if (lastRowActive) {
    result = mergeTableConditionalStyles(result, styles.lastRow);
  }

  const cornerFallbacks = {
    nwCell: firstRowActive && firstColumnActive,
    neCell: firstRowActive && lastColumnActive,
    swCell: lastRowActive && firstColumnActive,
    seCell: lastRowActive && lastColumnActive,
  } as const;
  for (const corner of ["nwCell", "neCell", "swCell", "seCell"] as const) {
    const explicit = cellCnf?.[corner];
    if (explicit ?? cornerFallbacks[corner]) {
      result = mergeTableConditionalStyles(result, styles[corner]);
    }
  }
  return result;
};

export type TableParagraphStyleContext = {
  pPr?: TableCellParagraphSpacingOverlay;
  rPr?: TextFormatting;
};

export type TableParagraphStyleContextByParagraph = {
  take: (paragraph: PMNode) => TableParagraphStyleContext | null | undefined;
};

export type TableRunFormattingByParagraph = {
  take: (paragraph: PMNode) => TextFormatting | null | undefined;
};

export const emptyTableParagraphStyleContext = (): TableParagraphStyleContextByParagraph => ({
  take: () => undefined,
});

export const emptyTableRunFormatting = (): TableRunFormattingByParagraph => ({
  take: () => undefined,
});

type PlacedTableCell = {
  cell: PMNode;
  columnIndex: number;
  columnSpan: number;
};

type PlacedTableRow = {
  cells: PlacedTableCell[];
  row: PMNode;
};

type TableGridPlacement = {
  rows: PlacedTableRow[];
  totalColumns: number;
};

/**
 * Place cells on the OOXML grid, including columns occupied by a rowspan from
 * an earlier row. PM removes the continuation cells when it imports vMerge,
 * so a row-local colspan sum shifts every later conditional region left.
 */
const placeTableCells = (table: PMNode): TableGridPlacement => {
  const rows: PMNode[] = [];
  table.forEach((row) => {
    if (row.type.name === "tableRow") {
      rows.push(row);
    }
  });

  const occupiedUntilRow: number[] = [];
  const placedRows: PlacedTableRow[] = [];
  let totalColumns = Math.max(1, expectTableAttrs(table).columnWidths?.length ?? 0);
  for (const [rowIndex, row] of rows.entries()) {
    const rowFormatting = expectTableRowAttrs(row)._originalFormatting;
    let columnIndex = rowFormatting?.gridBefore ?? 0;
    const cells: PlacedTableCell[] = [];
    row.forEach((cell) => {
      if (cell.type.name !== "tableCell" && cell.type.name !== "tableHeader") {
        return;
      }
      const cellAttrs = expectTableCellAttrs(cell);
      const columnSpan = cellAttrs.colspan;
      const spanIsOccupied = (): boolean => {
        for (let column = columnIndex; column < columnIndex + columnSpan; column++) {
          if ((occupiedUntilRow[column] ?? 0) > rowIndex) {
            return true;
          }
        }
        return false;
      };
      while (spanIsOccupied()) {
        columnIndex += 1;
      }
      cells.push({ cell, columnIndex, columnSpan });
      const occupiedUntil = rowIndex + cellAttrs.rowspan;
      for (let column = columnIndex; column < columnIndex + columnSpan; column++) {
        occupiedUntilRow[column] = Math.max(occupiedUntilRow[column] ?? 0, occupiedUntil);
      }
      columnIndex += columnSpan;
    });
    totalColumns = Math.max(totalColumns, columnIndex + (rowFormatting?.gridAfter ?? 0));
    placedRows.push({ cells, row });
  }
  return { rows: placedRows, totalColumns };
};

/** Index the conditional table paragraph/run cascade for every paragraph occurrence. */
export const tableParagraphStyleContextByParagraph = (
  doc: PMNode,
  styleResolver: TableConditionalStyleResolver | null,
): TableParagraphStyleContextByParagraph => {
  const formattingByParagraph = new WeakMap<
    PMNode,
    Array<TableParagraphStyleContext | null>
  >();
  const append = (paragraph: PMNode, formatting: TableParagraphStyleContext | null): void => {
    const existing = formattingByParagraph.get(paragraph);
    if (existing) {
      existing.push(formatting);
    } else {
      formattingByParagraph.set(paragraph, [formatting]);
    }
  };

  const visit = (node: PMNode, tableFormatting: TableParagraphStyleContext | null = null): void => {
    if (node.type.name === "paragraph") {
      append(node, tableFormatting);
      return;
    }
    if (node.type.name === "table") {
      visitTable(node);
      return;
    }
    const childTableFormatting = node.type.name === "textBox" ? null : tableFormatting;
    node.forEach((child) => visit(child, childTableFormatting));
  };

  const visitTable = (table: PMNode): void => {
    const tableAttrs = expectTableAttrs(table);
    const styles = resolveTableConditionalStyles(styleResolver, tableAttrs.styleId);
    const { rows, totalColumns } = placeTableCells(table);
    for (const [rowIndex, { cells, row }] of rows.entries()) {
      const rowFormatting = expectTableRowAttrs(row)._originalFormatting;
      for (const { cell, columnIndex, columnSpan } of cells) {
        const cellAttrs = expectTableCellAttrs(cell);
        const conditional = resolveTableCellConditionalStyle({
          styles,
          look: tableAttrs.look,
          rowIndex,
          totalRows: rows.length,
          columnIndex,
          columnSpan,
          totalColumns,
          rowFormatting,
          cellFormatting: cellAttrs._originalFormatting,
        });
        cell.forEach((block) => visit(block, conditional ?? {}));
      }
    }
  };

  visit(doc);
  const offsets = new WeakMap<PMNode, number>();
  return {
    take: (paragraph) => {
      const values = formattingByParagraph.get(paragraph);
      const offset = offsets.get(paragraph) ?? 0;
      const formatting = values?.at(offset);
      if (formatting === undefined) {
        return panic("Paragraph traversal diverged from its structural table-format index");
      }
      offsets.set(paragraph, offset + 1);
      return formatting;
    },
  };
};

/** Project the occurrence index into stable ProseMirror paragraph positions. */
export const tableParagraphStyleContextAtPositions = (
  doc: PMNode,
  styleResolver: TableConditionalStyleResolver | null,
): ReadonlyMap<number, TableParagraphStyleContext | null> => {
  const byOccurrence = tableParagraphStyleContextByParagraph(doc, styleResolver);
  const byPosition = new Map<number, TableParagraphStyleContext | null>();
  doc.descendants((node, position) => {
    if (node.type.name === "paragraph") {
      byPosition.set(position, byOccurrence.take(node) ?? null);
      return false;
    }
    return true;
  });
  return byPosition;
};

/** Run-only view used by PM-to-document serialization. */
export const tableRunFormattingByParagraph = (
  doc: PMNode,
  styleResolver: TableConditionalStyleResolver | null,
): TableRunFormattingByParagraph => {
  const contexts = tableParagraphStyleContextByParagraph(doc, styleResolver);
  return {
    take: (paragraph) => {
      const context = contexts.take(paragraph);
      return context === undefined ? undefined : (context?.rPr ?? null);
    },
  };
};

/** Run-only paragraph-position view retained for run-formatting callers. */
export const tableRunFormattingAtParagraphPositions = (
  doc: PMNode,
  styleResolver: TableConditionalStyleResolver | null,
): ReadonlyMap<number, TextFormatting | null> => {
  const contexts = tableParagraphStyleContextAtPositions(doc, styleResolver);
  return new Map(
    [...contexts].map(([position, context]) => [position, context?.rPr ?? null] as const),
  );
};
