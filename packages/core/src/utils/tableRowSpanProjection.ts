import type { Paragraph, Table, TableCell } from "../types/document";

export type TableCellRowSpanProjection = {
  rowSpan: number;
  skip: boolean;
  preserveVMergeRestart?: boolean;
  continuationCells?: TableCell[];
};

const cellHasMeaningfulContent = (cell: TableCell): boolean =>
  cell.content.some(blockHasMeaningfulContent);

const blockHasMeaningfulContent = (block: Paragraph | Table): boolean => {
  if (block.type === "table") {
    return block.rows.some((row) => row.cells.some(cellHasMeaningfulContent));
  }
  return block.content.some(paragraphContentHasMeaningfulContent);
};

const paragraphContentHasMeaningfulContent = (
  content: Paragraph["content"][number],
): boolean => {
  if (content.type === "run") return content.content.length > 0;
  if (content.type === "hyperlink") {
    return content.children.some(paragraphContentHasMeaningfulContent);
  }
  if (
    content.type === "insertion" ||
    content.type === "deletion" ||
    content.type === "moveFrom" ||
    content.type === "moveTo"
  ) {
    return content.content.some(paragraphContentHasMeaningfulContent);
  }
  return true;
};

const clearActiveVerticalMerges = (
  activeMerges: Map<number, number>,
  result: Map<string, TableCellRowSpanProjection>,
): void => {
  for (const [colIndex, startRow] of activeMerges) {
    const restartCell = result.get(`${String(startRow)}-${String(colIndex)}`);
    if (restartCell) restartCell.preserveVMergeRestart = true;
  }
  activeMerges.clear();
};

/**
 * Project the exact vertical-merge cells imported into the editor. Both DOCX
 * import and live comparison consume this one model-level interpretation.
 */
export const projectTableCellRowSpans = (
  table: Table,
): ReadonlyMap<string, TableCellRowSpanProjection> => {
  const result = new Map<string, TableCellRowSpanProjection>();
  const activeMerges = new Map<number, number>();
  for (const [rowIndex, row] of table.rows.entries()) {
    if (row.cells.length === 0) {
      clearActiveVerticalMerges(activeMerges, result);
      continue;
    }
    let colIndex = row.formatting?.gridBefore ?? 0;
    const rowCells = row.cells.map((cell) => {
      const colspan = cell.formatting?.gridSpan ?? 1;
      const vMerge = cell.formatting?.vMerge;
      const startRow = vMerge === "continue" ? activeMerges.get(colIndex) : undefined;
      const projected = {
        cell,
        colIndex,
        vMerge,
        startRow,
        hasMeaningfulContent: cellHasMeaningfulContent(cell),
        shouldSkip: vMerge === "continue" && startRow !== undefined,
      };
      colIndex += colspan;
      return projected;
    });
    const rowWouldBeEmpty = rowCells.every(({ shouldSkip }) => shouldSkip);

    for (const cell of rowCells) {
      const key = `${String(rowIndex)}-${String(cell.colIndex)}`;
      if (cell.vMerge === "restart") {
        activeMerges.set(cell.colIndex, rowIndex);
        result.set(key, { rowSpan: 1, skip: false });
        continue;
      }
      if (cell.vMerge === "continue") {
        if (
          cell.startRow === undefined ||
          rowWouldBeEmpty ||
          cell.hasMeaningfulContent
        ) {
          result.set(key, { rowSpan: 1, skip: false });
          if (
            (rowWouldBeEmpty || cell.hasMeaningfulContent) &&
            cell.startRow !== undefined
          ) {
            const restartCell = result.get(
              `${String(cell.startRow)}-${String(cell.colIndex)}`,
            );
            if (restartCell) restartCell.preserveVMergeRestart = true;
            activeMerges.delete(cell.colIndex);
          }
          continue;
        }
        const startCell = result.get(
          `${String(cell.startRow)}-${String(cell.colIndex)}`,
        );
        if (startCell) {
          startCell.rowSpan++;
          startCell.continuationCells ??= [];
          startCell.continuationCells.push(cell.cell);
        }
        result.set(key, { rowSpan: 1, skip: true });
        continue;
      }
      activeMerges.delete(cell.colIndex);
      result.set(key, { rowSpan: 1, skip: false });
    }
  }
  return result;
};
