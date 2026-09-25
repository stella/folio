/**
 * Places measured tables onto pages: row splitting across pages, repeated
 * header rows, rendered-break and tracked-change row handling, row footnote
 * collection, and floating table placement.
 */

import { FOOTNOTE_SEPARATOR_HEIGHT, createPaginator } from "./paginator";
import type { PageState } from "./paginator";
import { isEmptyParagraph } from "./paragraphSpacing";
import { buildTableRowBreakInfo, getRowContinuationSkip, snapRowBreak } from "./tableRowBreak";
import { resolveFloatingTablePageX } from "./measure/floatingTablePosition";
import { resolveTableInlineOffset } from "./measure/tableInlinePlacement";
import type { FlowBlock, TableBlock, TableCell, TableMeasure, TableFragment } from "./types";
import {
  RENDERED_BREAK_REFLOW_TOLERANCE_LINES,
  projectedFootnoteReserveGrowth,
} from "./layoutFlowShared";

/**
 * Count consecutive header rows at the start of a table.
 * Header rows are marked with isHeader: true in the block data.
 */
function countHeaderRows(block: TableBlock): number {
  let count = 0;
  for (const row of block.rows) {
    if (row.isHeader) {
      count++;
    } else {
      break;
    }
  }
  return count;
}

/**
 * Calculate total height of header rows from their measures.
 */
export function getHeaderRowsHeight(measure: TableMeasure, headerRowCount: number): number {
  let height = 0;
  for (let i = 0; i < headerRowCount && i < measure.rows.length; i++) {
    height += measure.rows[i]!.height; // SAFETY: i < measure.rows.length
  }
  return height;
}

const tableRowStartsWithRenderedPageBreak = (block: TableBlock, rowIndex: number): boolean => {
  const row = block.rows[rowIndex];
  const visibleCells = row?.cells.filter((cell) =>
    cell.blocks.some((cellBlock) => cellBlock.kind !== "paragraph" || !isEmptyParagraph(cellBlock)),
  );
  if (!visibleCells || visibleCells.length === 0) {
    return false;
  }
  const startsWithRenderedPageBreak = (cell: TableCell): boolean => {
    const firstVisibleBlock = cell.blocks.find(
      (cellBlock) => cellBlock.kind !== "paragraph" || !isEmptyParagraph(cellBlock),
    );
    return (
      firstVisibleBlock?.kind === "paragraph" &&
      firstVisibleBlock.attrs?.renderedPageBreakBefore === true
    );
  };

  // A fixed row's leading marker describes the row boundary once, in its
  // first visible cell. Flexible rows need agreement across visible cells
  // because unmarked siblings may carry content from the preceding page.
  if (row?.heightRule === "exact") {
    return startsWithRenderedPageBreak(visibleCells[0]!); // SAFETY: guarded by length check.
  }
  return visibleCells.every(startsWithRenderedPageBreak);
};

const getVerticallyMergedRows = (block: TableBlock): Set<number> => {
  const mergedRows = new Set<number>();
  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex += 1) {
    const row = block.rows[rowIndex];
    if (!row) {
      continue;
    }
    for (const cell of row.cells) {
      const rowSpan = cell.rowSpan ?? 1;
      if (rowSpan <= 1) {
        continue;
      }
      for (
        let mergedRowIndex = rowIndex;
        mergedRowIndex < Math.min(block.rows.length, rowIndex + rowSpan);
        mergedRowIndex += 1
      ) {
        mergedRows.add(mergedRowIndex);
      }
    }
  }
  return mergedRows;
};

const flowBlockHasTrackedChanges = (block: FlowBlock): boolean => {
  if (block.kind === "paragraph") {
    return block.runs.some((run) => {
      if (run.kind === "lineBreak") {
        return false;
      }
      return run.isInsertion || run.isDeletion;
    });
  }
  if (block.kind === "table") {
    return block.rows.some((row) =>
      row.cells.some((cell) => cell.blocks.some(flowBlockHasTrackedChanges)),
    );
  }
  if (block.kind === "textBox") {
    return block.content.some(flowBlockHasTrackedChanges);
  }
  return false;
};

const tableRowHasTrackedChanges = (block: TableBlock, rowIndex: number): boolean =>
  block.rows[rowIndex]?.cells.some((cell) => cell.blocks.some(flowBlockHasTrackedChanges)) ?? false;

/**
 * Layout a table block onto pages.
 */
export function layoutTable(
  block: TableBlock,
  measure: TableMeasure,
  paginator: ReturnType<typeof createPaginator>,
  footnoteHeightById?: Map<number, number>,
): void {
  const rows = measure.rows;
  if (rows.length === 0) {
    return;
  }
  const rowFootnoteIds = block.rows.map((row) =>
    collectTableRowFootnoteIds(row, footnoteHeightById),
  );

  // Detect header rows (consecutive rows at start with isHeader: true)
  const headerRowCount = countHeaderRows(block);
  const headerRowsHeight = getHeaderRowsHeight(measure, headerRowCount);

  let currentRowIndex = 0;

  const breakInfo = buildTableRowBreakInfo(block, measure);
  const verticallyMergedRows = getVerticallyMergedRows(block);
  // X position from justification / indent, recomputed per fragment because the
  // active column can change across section breaks.
  const computeTableX = ({
    columnIndex,
    rowIndex,
  }: {
    columnIndex: number;
    rowIndex: number;
  }): number => {
    const x = paginator.getColumnX(columnIndex);
    return (
      x +
      resolveTableInlineOffset({
        table: block,
        rowJustification: block.rows[rowIndex]?.justification,
        frameWidth: paginator.columnWidth,
        tableWidth: measure.totalWidth,
      })
    );
  };

  const getCurrentRowCapacity = (state = paginator.getCurrentState()): number =>
    state.rawContentBottom - state.topMargin;

  const hasAdjacentPriorTableRows = (
    rowIndex: number,
    state = paginator.getCurrentState(),
  ): boolean => {
    const previous = state.page.fragments.at(-1);
    return (
      previous?.kind === "table" &&
      previous.blockId === block.id &&
      previous.toRow === rowIndex &&
      previous.y + previous.height === state.cursorY
    );
  };

  const shouldRepeatHeaderRows = (
    rowIndex: number,
    consumed: number,
    state = paginator.getCurrentState(),
  ): boolean =>
    headerRowCount > 0 &&
    rowIndex >= headerRowCount &&
    !(consumed === 0 && hasAdjacentPriorTableRows(rowIndex, state));

  // True when nothing of this column precedes the row except, possibly, this
  // table's own leading header rows placed at the column top. Moving on from
  // there gains no room: the next region would repeat those same headers.
  const rowOpensFlowRegion = (rowIndex: number, state: PageState): boolean => {
    if (state.cursorY === state.topMargin) {
      return true;
    }
    const previous = state.page.fragments.at(-1);
    return (
      rowIndex > 0 &&
      rowIndex <= headerRowCount &&
      previous?.kind === "table" &&
      previous.blockId === block.id &&
      previous.fromRow === 0 &&
      previous.toRow === rowIndex &&
      previous.y === state.topMargin &&
      previous.y + previous.height === state.cursorY
    );
  };

  const canSplitRow = (rowIndex: number, state = paginator.getCurrentState()): boolean => {
    const row = rows[rowIndex];
    const sourceRow = block.rows[rowIndex];
    if (!row || !sourceRow || sourceRow.isHeader || verticallyMergedRows.has(rowIndex)) {
      return false;
    }
    if ((breakInfo.breakOffsets[rowIndex]?.length ?? 0) <= 1) {
      return false;
    }
    const freshHeaderOverhead =
      headerRowCount > 0 && rowIndex >= headerRowCount ? headerRowsHeight : 0;
    const requiredHeight = row.height + freshHeaderOverhead;
    const oversized = requiredHeight > getCurrentRowCapacity(state);
    if (sourceRow.cantSplit) {
      // w:cantSplit (§17.4.6) keeps the row on one page only while a page can
      // hold it. A taller row first moves to a fresh flow region and splits
      // there; an exact w:trHeight row keeps its fixed box and never splits.
      return oversized && sourceRow.heightRule !== "exact" && rowOpensFlowRegion(rowIndex, state);
    }
    if (!oversized && (state.footnoteHeight > 0 || (rowFootnoteIds[rowIndex]?.length ?? 0) > 0)) {
      return false;
    }

    const currentHeaderOverhead = shouldRepeatHeaderRows(rowIndex, 0, state) ? headerRowsHeight : 0;
    const currentAvailableHeight =
      paginator.getAvailableHeight() - currentHeaderOverhead - state.trailingSpacing;
    return oversized || row.height > currentAvailableHeight;
  };

  while (currentRowIndex < rows.length) {
    const rowState = paginator.getCurrentState();
    const rowHeaderOverhead = shouldRepeatHeaderRows(currentRowIndex, 0, rowState)
      ? headerRowsHeight
      : 0;
    const rowAvailableHeight =
      paginator.getAvailableHeight() - rowHeaderOverhead - rowState.trailingSpacing;
    const rowStartsFreshPage =
      rowState.cursorY === rowState.topMargin && rowState.page.fragments.length === 0;

    if (block.rows[currentRowIndex]?.breakBefore === "page" && !rowStartsFreshPage) {
      paginator.forcePageBreak();
      continue;
    }

    // A leading w:lastRenderedPageBreak is Word's cached boundary for this
    // row. Keep the hint advisory while the row fits, but when Folio would
    // otherwise split it in the remaining page space, snap the row to the
    // cached page boundary. Oversized rows still split after reaching the
    // fresh page, so the hint cannot create a retry loop.
    if (
      !rowStartsFreshPage &&
      tableRowStartsWithRenderedPageBreak(block, currentRowIndex) &&
      !tableRowHasTrackedChanges(block, currentRowIndex) &&
      rows[currentRowIndex]!.height > rowAvailableHeight
    ) {
      paginator.forcePageBreak();
      continue;
    }

    // Break permitted rows between whole text lines when they exceed the current
    // flow region; rows taller than a full region use the same path repeatedly.
    const splittableRow = rows[currentRowIndex]!; // SAFETY: currentRowIndex < rows.length
    if (canSplitRow(currentRowIndex)) {
      let consumed = 0;
      const renderedBreakHints = tableRowHasTrackedChanges(block, currentRowIndex)
        ? []
        : (breakInfo.renderedBreakHints[currentRowIndex] ?? []);
      let renderedBreakHintIndex = 0;
      const discardPassedRenderedBreakHints = (): void => {
        while ((renderedBreakHints[renderedBreakHintIndex]?.offset ?? Infinity) <= consumed) {
          renderedBreakHintIndex += 1;
        }
      };
      const advanceAfterNaturalFlowBreak = (previousPageNumber: number): void => {
        discardPassedRenderedBreakHints();
        if (
          paginator.getCurrentState().page.number > previousPageNumber &&
          renderedBreakHintIndex < renderedBreakHints.length
        ) {
          renderedBreakHintIndex += 1;
        }
      };
      while (consumed < splittableRow.height) {
        const sliceState = paginator.getCurrentState();
        const repeatHeaderRows = shouldRepeatHeaderRows(currentRowIndex, consumed, sliceState);
        const headerOverhead = repeatHeaderRows ? headerRowsHeight : 0;
        const sliceAvail =
          paginator.getAvailableHeight() -
          headerOverhead -
          (consumed === 0 ? sliceState.trailingSpacing : 0);
        let slice = snapRowBreak(breakInfo, currentRowIndex, consumed, sliceAvail);
        let forceRenderedPageBreak = false;
        if (slice <= 0) {
          const isFreshPage =
            sliceState.cursorY === sliceState.topMargin && sliceState.page.fragments.length === 0;
          if (!isFreshPage) {
            // Not even one line fits in the space left; continue in the next
            // column, or on a fresh page when this is the last column.
            const previousPageNumber = sliceState.page.number;
            paginator.forceColumnBreak();
            advanceAfterNaturalFlowBreak(previousPageNumber);
            continue;
          }
          // Fresh page and a single line still exceeds the page height: place
          // the next whole line anyway so the loop always makes progress.
          const from = consumed;
          const next = breakInfo.breakOffsets[currentRowIndex]?.find((o) => o > from);
          slice = (next ?? splittableRow.height) - consumed;
        }
        discardPassedRenderedBreakHints();
        const renderedBreakHint = renderedBreakHints[renderedBreakHintIndex];
        if (
          renderedBreakHint &&
          renderedBreakHint.offset <= consumed + slice &&
          sliceAvail - (renderedBreakHint.offset - consumed) <=
            renderedBreakHint.lineAdvance * RENDERED_BREAK_REFLOW_TOLERANCE_LINES
        ) {
          slice = renderedBreakHint.offset - consumed;
          renderedBreakHintIndex += 1;
          forceRenderedPageBreak = true;
        }
        const sliceBottom = consumed + slice;
        const continuationSkip =
          sliceBottom < splittableRow.height
            ? getRowContinuationSkip(breakInfo, currentRowIndex, sliceBottom)
            : 0;
        const nextConsumed = Math.min(splittableRow.height, sliceBottom + continuationSkip);
        const reachesRowEnd = nextConsumed >= splittableRow.height;
        const moreAfter = !reachesRowEnd || currentRowIndex + 1 < rows.length;
        const fragmentHeight = headerOverhead + slice;
        const sliceFragment: TableFragment = {
          kind: "table",
          blockId: block.id,
          x: computeTableX({
            columnIndex: sliceState.columnIndex,
            rowIndex: currentRowIndex,
          }),
          y: 0,
          width: measure.totalWidth,
          height: fragmentHeight,
          fromRow: currentRowIndex,
          toRow: currentRowIndex + 1,
          ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
          ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
          ...(consumed > 0 || currentRowIndex > 0 ? { continuesFromPrev: true } : {}),
          ...(moreAfter ? { continuesOnNext: true } : {}),
          ...(repeatHeaderRows ? { headerRowCount } : {}),
          ...(consumed > 0 ? { topClip: consumed } : {}),
          ...(sliceBottom >= splittableRow.height ? {} : { bottomClip: sliceBottom }),
          ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
        };
        const sliceResult = paginator.addFragment(sliceFragment, fragmentHeight, 0, 0);
        sliceFragment.y = sliceResult.y;
        sliceFragment.x = computeTableX({
          columnIndex: sliceResult.state.columnIndex,
          rowIndex: currentRowIndex,
        });
        consumed = nextConsumed;
        if (consumed < splittableRow.height) {
          if (forceRenderedPageBreak) {
            paginator.forcePageBreak();
          } else {
            const previousPageNumber = paginator.getCurrentState().page.number;
            paginator.forceColumnBreak();
            advanceAfterNaturalFlowBreak(previousPageNumber);
          }
        }
      }
      currentRowIndex += 1;
      continue;
    }

    const state = paginator.getCurrentState();
    const rawAvailableHeight = paginator.getAvailableHeight();
    const isFirstFragment = currentRowIndex === 0;

    // Leading skip past a page-pinned band, applied only to the table's first
    // fragment (a band sits on one page). eigenpal #694.
    const bandSkip = isFirstFragment ? (measure.bandSkipBefore ?? 0) : 0;

    // Account for the space addFragment will consume before the fragment, which
    // is max(spaceBefore, trailingSpacing). We pass bandSkip as spaceBefore, so
    // the overhead is the larger of that and the previous block's trailing space.
    const pendingSpacing = isFirstFragment ? Math.max(bandSkip, state.trailingSpacing) : 0;
    const availableHeight = rawAvailableHeight - pendingSpacing;

    const repeatHeaderRowsForNormalFragment = shouldRepeatHeaderRows(currentRowIndex, 0, state);

    // For continuation fragments, we need space for header rows + at least one content row.
    const normalHeaderOverhead = repeatHeaderRowsForNormalFragment ? headerRowsHeight : 0;

    // Calculate how many rows fit (excluding header rows which are prepended separately)
    let rowsHeight = 0;
    let fittingRows = 0;
    const pageFootnoteIds = new Set(state.page.footnoteIds ?? []);
    const fragmentFootnoteIds: number[] = [];
    let fragmentFootnoteHeight = 0;
    let retryOnNextFlowRegion = false;
    const fragmentX = computeTableX({
      columnIndex: state.columnIndex,
      rowIndex: currentRowIndex,
    });

    for (let j = currentRowIndex; j < rows.length; j++) {
      if (j > currentRowIndex && block.rows[j]?.breakBefore === "page") {
        break;
      }
      if (
        j > currentRowIndex &&
        computeTableX({ columnIndex: state.columnIndex, rowIndex: j }) !== fragmentX
      ) {
        break;
      }
      const rowHeight = rows[j]!.height; // SAFETY: j < rows.length
      const currentRowFootnoteIds = rowFootnoteIds[j] ?? [];
      const fragmentFootnoteCountBeforeRow = fragmentFootnoteIds.length;
      let rowFootnoteHeight = 0;
      for (const id of currentRowFootnoteIds) {
        if (pageFootnoteIds.has(id) || fragmentFootnoteIds.includes(id)) {
          continue;
        }
        fragmentFootnoteIds.push(id);
        rowFootnoteHeight += footnoteHeightById?.get(id) ?? 0;
      }
      const candidateFootnoteHeight = fragmentFootnoteHeight + rowFootnoteHeight;
      const separatorHeight =
        state.footnoteDemandHeight === 0 && candidateFootnoteHeight > 0
          ? FOOTNOTE_SEPARATOR_HEIGHT
          : 0;
      const footnoteGrowth = projectedFootnoteReserveGrowth(
        state,
        candidateFootnoteHeight + separatorHeight,
      );
      const totalWithRow = rowsHeight + rowHeight + normalHeaderOverhead + footnoteGrowth;

      if (totalWithRow <= availableHeight) {
        rowsHeight += rowHeight;
        fragmentFootnoteHeight += rowFootnoteHeight;
        fittingRows++;
      } else if (fittingRows === 0) {
        const isFreshFlowRegion =
          state.cursorY === state.topMargin && state.page.fragments.length === 0;
        if (isFreshFlowRegion) {
          rowsHeight += rowHeight;
          fragmentFootnoteHeight += rowFootnoteHeight;
          fittingRows++;
          break;
        }
        fragmentFootnoteIds.splice(fragmentFootnoteCountBeforeRow);
        paginator.forceColumnBreak();
        retryOnNextFlowRegion = true;
        break;
      } else {
        fragmentFootnoteIds.splice(fragmentFootnoteCountBeforeRow);
        break;
      }
    }

    if (retryOnNextFlowRegion) {
      continue;
    }

    // Total fragment height includes header rows for continuation fragments
    const fragmentHeight = rowsHeight + normalHeaderOverhead;

    // Create fragment for these rows
    const isLastFragment = currentRowIndex + fittingRows >= rows.length;

    // Calculate x position based on table justification and indent
    const desiredX = fragmentX;

    const fragment: TableFragment = {
      kind: "table",
      blockId: block.id,
      x: desiredX,
      y: 0, // Will be set by addFragment
      width: measure.totalWidth,
      height: fragmentHeight,
      fromRow: currentRowIndex,
      toRow: currentRowIndex + fittingRows,
      ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
      ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
      ...(!isFirstFragment ? { continuesFromPrev: true } : {}),
      ...(!isLastFragment ? { continuesOnNext: true } : {}),
      ...(repeatHeaderRowsForNormalFragment ? { headerRowCount } : {}),
      ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
    };

    if (fragmentFootnoteHeight > 0) {
      paginator.addFootnoteHeight(fragmentFootnoteHeight, fragmentFootnoteIds);
    }
    const result = paginator.addFragment(fragment, fragmentHeight, bandSkip, 0);
    fragment.y = result.y;
    fragment.x = desiredX;

    currentRowIndex += fittingRows;

    // If more rows remain, advance to next column/page
    if (currentRowIndex < rows.length) {
      if (canSplitRow(currentRowIndex)) {
        const nextState = paginator.getCurrentState();
        const nextHeaderOverhead = shouldRepeatHeaderRows(currentRowIndex, 0, nextState)
          ? headerRowsHeight
          : 0;
        const nextSliceAvail =
          paginator.getAvailableHeight() - nextHeaderOverhead - nextState.trailingSpacing;
        if (snapRowBreak(breakInfo, currentRowIndex, 0, nextSliceAvail) > 0) {
          continue;
        }
      }
      // Need space for at least one content row plus repeated header rows
      const nextState = paginator.getCurrentState();
      const nextRowHeight =
        rows[currentRowIndex]!.height + // SAFETY: guarded by length check
        (shouldRepeatHeaderRows(currentRowIndex, 0, nextState) ? headerRowsHeight : 0);
      paginator.ensureFits(nextRowHeight);
    }
  }
}

function collectTableRowFootnoteIds(
  row: TableBlock["rows"][number] | undefined,
  footnoteHeightById: Map<number, number> | undefined,
): number[] {
  // A footnote referenced only from a hidden row must not reserve or paint
  // its body — Word never renders a hidden row, so its footnote reference
  // never becomes visible either.
  if (!row || !footnoteHeightById || row.hidden) {
    return [];
  }

  const ids: number[] = [];
  const walk = (blocks: FlowBlock[]): void => {
    for (const block of blocks) {
      if (block.kind === "paragraph") {
        for (const run of block.runs) {
          if (
            run.kind === "text" &&
            run.footnoteRefId !== undefined &&
            footnoteHeightById.has(run.footnoteRefId) &&
            !ids.includes(run.footnoteRefId)
          ) {
            ids.push(run.footnoteRefId);
          }
        }
        continue;
      }
      if (block.kind === "table") {
        for (const nestedRow of block.rows) {
          if (nestedRow.hidden) {
            continue;
          }
          for (const cell of nestedRow.cells) {
            walk(cell.blocks);
          }
        }
        continue;
      }
      if (block.kind === "textBox") {
        walk(block.content);
      }
    }
  };

  for (const cell of row.cells) {
    walk(cell.blocks);
  }
  return ids;
}

/**
 * Layout a floating table (anchored) without advancing the cursor.
 */
export function layoutFloatingTable(
  block: TableBlock,
  measure: TableMeasure,
  paginator: ReturnType<typeof createPaginator>,
  contentWidth: number,
): void {
  const state = paginator.getCurrentState();
  const floating = block.floating;
  const page = state.page;
  const margins = page.margins;

  const tableWidth = measure.totalWidth;
  const tableHeight = measure.totalHeight;

  const contentHeight = page.size.h - margins.top - margins.bottom;

  // Default anchor base (content area)
  let baseY = margins.top;

  if (floating?.vertAnchor === "page") {
    baseY = 0;
  } else if (floating?.vertAnchor === "text") {
    // A text-relative table position is offset from the current text
    // anchor, not from the page's body margin. Using the margin here lifts
    // later floating tables to the top of the page and overlays earlier text.
    baseY = state.cursorY;
  }

  // Determine X position
  let x = paginator.getColumnX(state.columnIndex);
  if (floating) {
    x = resolveFloatingTablePageX({
      anchor: floating,
      justification: block.justification,
      tableWidth,
      marginWidth: contentWidth,
      pageWidth: page.size.w,
      marginLeft: margins.left,
      textFrameWidth: paginator.columnWidth,
      textFrameLeft: paginator.getColumnX(state.columnIndex),
    });
  }

  // Determine Y position
  let y = state.cursorY;
  let usedExplicitY = false;
  if (floating?.tblpY !== undefined) {
    y = baseY + floating.tblpY;
    usedExplicitY = true;
  } else if (floating?.tblpYSpec) {
    usedExplicitY = true;
    const spec = floating.tblpYSpec;
    if (spec === "top") {
      y = baseY;
    } else if (spec === "bottom") {
      y = baseY + contentHeight - tableHeight;
    } else if (spec === "center") {
      y = baseY + (contentHeight - tableHeight) / 2;
    }
  }

  // If not explicitly positioned, ensure it fits on the current page
  if (!usedExplicitY) {
    const fitState = paginator.ensureFits(tableHeight);
    y = fitState.cursorY;
  }

  const fragment: TableFragment = {
    kind: "table",
    blockId: block.id,
    x,
    y,
    width: tableWidth,
    height: tableHeight,
    fromRow: 0,
    toRow: block.rows.length,
    ...(block.pmStart !== undefined ? { pmStart: block.pmStart } : {}),
    ...(block.pmEnd !== undefined ? { pmEnd: block.pmEnd } : {}),
    isFloating: true,
    ...(block.sdtGroups ? { sdtGroups: block.sdtGroups } : {}),
  };

  paginator.addUnflowedFragment(fragment);
}
