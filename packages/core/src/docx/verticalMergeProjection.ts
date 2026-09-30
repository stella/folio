import type { Table, TableCell, TableCellBlock, Paragraph } from "../types/document";
import { isCellMergeContinuation } from "./tableParser";

/** The reader and revision resolver share which continuation cells remain visible. */
export type RowSpanInfo = {
  rowSpan: number;
  skip: boolean;
  preserveVMergeRestart?: boolean;
  continuationCells?: TableCell[];
};

export function calculateRowSpans(table: Table): Map<string, RowSpanInfo> {
  const result = new Map<string, RowSpanInfo>();
  const numRows = table.rows.length;

  // Track active vertical merges per column (stores the row index where merge started)
  const activeMerges = new Map<number, number>();

  // Process each row
  for (let rowIndex = 0; rowIndex < numRows; rowIndex++) {
    // SAFETY: rowIndex < numRows <= table.rows.length
    const row = table.rows[rowIndex]!;
    if (row.cells.length === 0) {
      clearActiveVerticalMerges(activeMerges, result);
      continue;
    }
    let colIndex = row.formatting?.gridBefore ?? 0;
    const rowCells = row.cells.map((cell) => {
      const colspan = cell.formatting?.gridSpan ?? 1;
      const vMerge = cell.formatting?.vMerge;
      const isMergeContinuation = isCellMergeContinuation(cell);
      const startRow = isMergeContinuation ? activeMerges.get(colIndex) : undefined;
      const info = {
        cell,
        colIndex,
        colspan,
        vMerge,
        isMergeContinuation,
        startRow,
        hasMeaningfulContent: tableCellHasMeaningfulContent(cell),
        shouldSkip: isMergeContinuation && startRow !== undefined,
      };
      colIndex += colspan;
      return info;
    });
    const rowWouldBeEmpty = rowCells.length > 0 && rowCells.every((cell) => cell.shouldSkip);

    for (const cellInfo of rowCells) {
      const {
        colIndex: cellColIndex,
        vMerge,
        isMergeContinuation,
        startRow,
        hasMeaningfulContent,
      } = cellInfo;
      const key = `${rowIndex}-${cellColIndex}`;

      if (vMerge === "restart") {
        // Start of a new vertical merge. A restart directly under another one
        // ends that one, so close it before this row takes the column over.
        closeVerticalMerge(activeMerges, result, cellColIndex);
        activeMerges.set(cellColIndex, rowIndex);
        result.set(key, { rowSpan: 1, skip: false });
      } else if (isMergeContinuation) {
        // Continuation of a merge - only skip it when the parsed grid has a
        // matching restart in this exact column and the continuation is only a
        // structural placeholder. Real DOCX tables can be ragged, and some
        // continuation cells contain drawings or other payload that must not be
        // merged away.
        if (startRow === undefined || rowWouldBeEmpty || hasMeaningfulContent) {
          result.set(key, { rowSpan: 1, skip: false });
          if ((rowWouldBeEmpty || hasMeaningfulContent) && startRow !== undefined) {
            const restartCell = result.get(`${startRow}-${cellColIndex}`);
            if (restartCell) {
              restartCell.preserveVMergeRestart = true;
            }
            activeMerges.delete(cellColIndex);
          }
          continue;
        }

        // Increment rowSpan of the starting cell
        const startKey = `${startRow}-${cellColIndex}`;
        const startCell = result.get(startKey);
        if (startCell) {
          startCell.rowSpan++;
          startCell.continuationCells ??= [];
          startCell.continuationCells.push(cellInfo.cell);
        }
        result.set(key, { rowSpan: 1, skip: true });
      } else {
        // No vMerge - clear any active merge for this column
        closeVerticalMerge(activeMerges, result, cellColIndex);
        result.set(key, { rowSpan: 1, skip: false });
      }
    }
  }

  // A merge still open when the table ends never gained a continuation, so
  // nothing but the flag records that the cell said `w:vMerge="restart"`.
  clearActiveVerticalMerges(activeMerges, result);

  return result;
}

/**
 * End the vertical merge active in one column.
 *
 * A `w:vMerge="restart"` is carried through the editor by its cell's rowspan,
 * which only exists once a continuation joins it. A merge that closes with a
 * rowspan of 1 — a restart with no continuation, one interrupted by a plain
 * cell, one whose continuation is in a row a revision removed — has nothing
 * but this flag to say the cell was a merge origin, and dropping it changes
 * the table's visible structure.
 */
function closeVerticalMerge(
  activeMerges: Map<number, number>,
  result: Map<string, RowSpanInfo>,
  colIndex: number,
): void {
  const startRow = activeMerges.get(colIndex);
  if (startRow === undefined) {
    return;
  }
  const restartCell = result.get(`${startRow}-${colIndex}`);
  // A merge that absorbed a continuation is already spelled by the rowspan,
  // and flagging it would resurrect the restart when a revision later splits
  // the cell back apart. Only a merge closing at a rowspan of one needs it.
  if (restartCell && restartCell.rowSpan === 1) {
    restartCell.preserveVMergeRestart = true;
  }
  activeMerges.delete(colIndex);
}

function clearActiveVerticalMerges(
  activeMerges: Map<number, number>,
  result: Map<string, RowSpanInfo>,
): void {
  for (const colIndex of [...activeMerges.keys()]) {
    closeVerticalMerge(activeMerges, result, colIndex);
  }
}

/**
 * Whether a cell carries content a vertical-merge continuation must not hide.
 * The reader keeps such a continuation as a cell of its own.
 */
export function tableCellHasMeaningfulContent(cell: TableCell): boolean {
  return cell.content.some(blockHasMeaningfulContent);
}

function blockHasMeaningfulContent(block: TableCellBlock): boolean {
  // Markup the cell carries is content, even though folio cannot read it.
  if (block.type === "preservedBlock") {
    return true;
  }
  // A delimiter is not content, but a cell holding one is not empty either:
  // pruning it would take the bookmark with it.
  if (block.type === "bookmarkStart" || block.type === "bookmarkEnd") {
    return true;
  }
  if (block.type === "table") {
    return block.rows.some((row) => row.cells.some((cell) => tableCellHasMeaningfulContent(cell)));
  }
  if (block.type === "blockSdt" || block.type === "blockCustomXml") {
    // The wrapper carries authored structure even when its children are empty.
    return true;
  }

  return block.content.some(paragraphContentHasMeaningfulContent);
}

function paragraphContentHasMeaningfulContent(content: Paragraph["content"][number]): boolean {
  if (content.type === "run") {
    return content.content.length > 0;
  }
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
}
