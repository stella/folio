/**
 * Table row-break geometry. Ported from eigenpal/docx-editor#698 (folio subset).
 *
 * Table rows may break across flow regions unless `cantSplit` is set. When a
 * permitted row exceeds the current region, the portion that fits stays and
 * the rest continues in the next region, broken between whole text lines and
 * never through a glyph. This also prevents oversized rows from overflowing
 * one page and clipping content (upstream #570).
 *
 * This computes, per row, the safe break offsets (the y of every line bottom in
 * the row's cells) so the paginator can snap a break to the deepest whole line
 * that still fits. A nested table contributes the lines and row bottoms of its
 * own rows, so it continues across regions like the table that holds it.
 * Vertically-merged cells that span into a row of the top-level table are not
 * modeled here yet (a separate follow-up); each row uses its own cells' content.
 */

import { measuredLineAdvance } from "./lineFlow";
import { measureParagraph } from "./measure";
import {
  buildTableCellFloatingZones,
  getTableCellContentWidth,
  getTableCellFloatingImages,
} from "./measure/tableCellFloating";
import {
  createTableCellFlowState,
  placeTableCellBlock,
} from "./measure/tableCellFlow";
import { getTableCellSpacingInsetsY } from "./measure/tableCellGrid";
import { isEmptyParagraph } from "./paragraphSpacing";
import {
  resolveTableCellPadding,
  type TableBlock,
  type TableCell,
  type TableCellMeasure,
  type TableMeasure,
  type TableRow,
} from "./types";

type UnsafeBreakRange = {
  top: number;
  bottom: number;
};

type RenderedRowBreakHint = {
  offset: number;
  lineAdvance: number;
};

type CellBreakGeometry = {
  bottoms: number[];
  renderedBreakHints: RenderedRowBreakHint[];
  /** Text lines and atomic blocks; nothing ever breaks through one. */
  unsafeRanges: UnsafeBreakRange[];
  /**
   * Nested rows that stay whole while a flow region can hold them. A region
   * too short for one breaks it between its lines instead of losing the rest.
   */
  keepTogetherRanges: UnsafeBreakRange[];
  suppressibleLeadingRanges: UnsafeBreakRange[];
};

const BREAK_OFFSET_EPSILON = 0.01;

/**
 * Sorted, disjoint form of `ranges`. Ranges that only touch stay separate:
 * their shared edge is a clean break.
 */
const mergeBreakRanges = (ranges: UnsafeBreakRange[]): UnsafeBreakRange[] => {
  const merged: UnsafeBreakRange[] = [];
  for (const range of ranges.toSorted((a, b) => a.top - b.top)) {
    const previous = merged.at(-1);
    if (previous && range.top < previous.bottom) {
      previous.bottom = Math.max(previous.bottom, range.bottom);
      continue;
    }
    merged.push({ ...range });
  }
  return merged;
};

/**
 * Binary search over the output of `mergeBreakRanges`. An offset within rounding
 * of an edge is on that edge: nested positions are sums of more line heights
 * than the sibling line they have to meet.
 */
const isInsideMergedRanges = (
  offset: number,
  ranges: UnsafeBreakRange[],
): boolean => {
  let low = 0;
  let high = ranges.length - 1;
  while (low <= high) {
    const middle = (low + high) >> 1;
    const range = ranges[middle];
    if (!range) {
      return false;
    }
    if (offset <= range.top + BREAK_OFFSET_EPSILON) {
      high = middle - 1;
    } else if (offset >= range.bottom - BREAK_OFFSET_EPSILON) {
      low = middle + 1;
    } else {
      return true;
    }
  }
  return false;
};

function getVerticalAlignmentOffset(
  cell: TableCell | undefined,
  measure: TableCellMeasure,
  rowHeight: number,
): number {
  const spareHeight = Math.max(0, rowHeight - measure.height);
  if (cell?.verticalAlign === "bottom") {
    return spareHeight;
  }
  if (cell?.verticalAlign === "center") {
    return spareHeight / 2;
  }
  return 0;
}

/** A nested table can outgrow the argument count `push(...items)` allows. */
const appendAll = <T>(target: T[], items: readonly T[]): void => {
  for (const item of items) {
    target.push(item);
  }
};

const shiftRange = (
  range: UnsafeBreakRange,
  offset: number,
): UnsafeBreakRange => ({
  top: range.top + offset,
  bottom: range.bottom + offset,
});

function shiftCellGeometry(
  geometry: CellBreakGeometry,
  offset: number,
): CellBreakGeometry {
  if (offset <= 0) {
    return geometry;
  }
  return {
    bottoms: geometry.bottoms.map((bottom) => bottom + offset),
    renderedBreakHints: geometry.renderedBreakHints.map((hint) => ({
      offset: hint.offset + offset,
      lineAdvance: hint.lineAdvance,
    })),
    unsafeRanges: geometry.unsafeRanges.map((range) =>
      shiftRange(range, offset),
    ),
    keepTogetherRanges: geometry.keepTogetherRanges.map((range) =>
      shiftRange(range, offset),
    ),
    suppressibleLeadingRanges: geometry.suppressibleLeadingRanges.map((range) =>
      shiftRange(range, offset),
    ),
  };
}

/** Each measured cell's geometry in one table row, relative to the row top. */
function rowCellGeometries(
  block: TableBlock,
  measure: TableMeasure,
  rowIndex: number,
): CellBreakGeometry[] {
  const rowHeight = measure.rows[rowIndex]?.height ?? 0;
  const sourceCells = block.rows[rowIndex]?.cells ?? [];
  const measuredCells = measure.rows[rowIndex]?.cells ?? [];
  const geometries: CellBreakGeometry[] = [];
  for (let c = 0; c < measuredCells.length; c++) {
    const measuredCell = measuredCells[c];
    if (!measuredCell) {
      continue;
    }
    const sourceCell = sourceCells[c];
    const spacing = getTableCellSpacingInsetsY(
      block.cellSpacing,
      rowIndex,
      sourceCell?.rowSpan ?? 1,
      block.rows.length,
    );
    geometries.push(
      shiftCellGeometry(
        cellBreakGeometry(sourceCell, measuredCell),
        spacing.top +
          getVerticalAlignmentOffset(
            sourceCell,
            measuredCell,
            Math.max(0, rowHeight - spacing.top - spacing.bottom),
          ),
      ),
    );
  }
  return geometries;
}

/** `w:cantSplit` (§17.4.6) and header rows stay whole while a region can hold them. */
const nestedRowKeepsTogether = (row: TableRow | undefined): boolean =>
  row?.cantSplit === true || row?.isHeader === true;

/**
 * Break geometry of a table nested in a cell, relative to its top. The cells of
 * all its rows join one flat geometry: a break must clear every line of every
 * cell, including a merged cell that spans several rows. Whitespace that a
 * continuation may consume is left out, because it would have to be shared with
 * the sibling cells flattening no longer tells apart.
 */
function nestedTableBreakGeometry(
  block: TableBlock,
  measure: TableMeasure,
): CellBreakGeometry {
  const nested: CellBreakGeometry = {
    bottoms: [],
    renderedBreakHints: [],
    unsafeRanges: [],
    keepTogetherRanges: [],
    suppressibleLeadingRanges: [],
  };
  let rowTop = 0;
  for (let r = 0; r < measure.rows.length; r++) {
    const rowHeight = measure.rows[r]?.height ?? 0;
    const sourceRow = block.rows[r];
    if (sourceRow?.hidden !== true && rowHeight > 0) {
      const rowBand = { top: rowTop, bottom: rowTop + rowHeight };
      nested.bottoms.push(rowBand.bottom);
      if (nestedRowKeepsTogether(sourceRow)) {
        nested.keepTogetherRanges.push(rowBand);
      }
      // A fixed `w:trHeight` box never splits, and it clips its content, so
      // nothing past the box can constrain a break either.
      const fixedBox = sourceRow?.heightRule === "exact";
      if (fixedBox) {
        nested.unsafeRanges.push(rowBand);
      }
      const cells = fixedBox ? [] : rowCellGeometries(block, measure, r);
      for (const cell of cells) {
        const geometry = shiftCellGeometry(cell, rowTop);
        appendAll(nested.bottoms, geometry.bottoms);
        appendAll(nested.renderedBreakHints, geometry.renderedBreakHints);
        appendAll(nested.unsafeRanges, geometry.unsafeRanges);
        appendAll(nested.keepTogetherRanges, geometry.keepTogetherRanges);
      }
    }
    rowTop += rowHeight;
  }
  return nested;
}

/** Cumulative break geometry within a single cell's content. */
function cellBreakGeometry(
  cell: TableCell | undefined,
  measure: TableCellMeasure,
): CellBreakGeometry {
  // Mirror the painter's shared cell-flow placement so clean row breaks stay
  // aligned with both glyph boundaries and inter-paragraph whitespace.
  const bottoms: number[] = [];
  const renderedBreakHints: RenderedRowBreakHint[] = [];
  const unsafeRanges: UnsafeBreakRange[] = [];
  const keepTogetherRanges: UnsafeBreakRange[] = [];
  const suppressibleLeadingRanges: UnsafeBreakRange[] = [];
  const cellBlocks = cell?.blocks;
  const blockMeasures = measure.blocks;
  const { top: padTop } = resolveTableCellPadding(cell);
  const contentWidth = getTableCellContentWidth(cell, measure);
  const floatingImages =
    cell !== undefined
      ? getTableCellFloatingImages(cell, measure, contentWidth)
      : [];
  const floatingZones = buildTableCellFloatingZones(
    floatingImages,
    contentWidth,
  );
  const flowState = createTableCellFlowState();
  for (let i = 0; i < blockMeasures.length; i++) {
    let blockMeasure = blockMeasures[i];
    if (blockMeasure?.kind === "paragraph") {
      const block = cellBlocks?.[i];
      if (block?.kind !== "paragraph") {
        continue;
      }
      if (floatingZones.length > 0) {
        const previewState = { ...flowState };
        const preview = placeTableCellBlock(previewState, block, blockMeasure);
        blockMeasure = measureParagraph(block, contentWidth, {
          floatingZones,
          paragraphYOffset: preview.contentTop,
        });
      }
      const placement = placeTableCellBlock(flowState, block, blockMeasure);
      if (placement.leadingSpacing > 0 && isEmptyParagraph(block)) {
        suppressibleLeadingRanges.push({
          top: padTop + placement.top,
          bottom: padTop + placement.contentTop,
        });
      }
      // Paragraph lines paint after the resolved leading spacing, matching the
      // paragraph fragment's top padding in the cell painter.
      let y = padTop + placement.contentTop;
      for (const line of blockMeasure.lines) {
        y += line.floatSkipBefore ?? 0;
        const top = y;
        if (line.renderedPageBreakBefore === true) {
          renderedBreakHints.push({
            offset: top,
            lineAdvance: measuredLineAdvance(line),
          });
        }
        y += line.lineHeight;
        unsafeRanges.push({ top, bottom: y });
        bottoms.push(y);
      }
      // Inter-paragraph spacing is also a safe break band. Recording the box
      // boundary lets pagination use available whitespace instead of backing
      // up to the preceding glyph line.
      bottoms.push(padTop + flowState.height);
    } else if (blockMeasure) {
      const block = cellBlocks?.[i];
      if (!block) {
        continue;
      }
      const placement = placeTableCellBlock(flowState, block, blockMeasure);
      const blockHeight = placement.contentHeight;
      if (blockHeight > 0) {
        const top = padTop + placement.contentTop;
        const y = top + blockHeight;
        bottoms.push(y);
        if (block.kind === "table" && blockMeasure.kind === "table") {
          const nested = shiftCellGeometry(
            nestedTableBreakGeometry(block, blockMeasure),
            top,
          );
          appendAll(bottoms, nested.bottoms);
          appendAll(renderedBreakHints, nested.renderedBreakHints);
          appendAll(unsafeRanges, nested.unsafeRanges);
          appendAll(keepTogetherRanges, nested.keepTogetherRanges);
        } else {
          unsafeRanges.push({ top, bottom: y });
        }
      }
    }
  }
  return {
    bottoms,
    renderedBreakHints,
    unsafeRanges,
    keepTogetherRanges,
    suppressibleLeadingRanges,
  };
}

const mergeRenderedBreakHints = (
  hints: RenderedRowBreakHint[],
): RenderedRowBreakHint[] => {
  const merged: RenderedRowBreakHint[] = [];
  for (const hint of hints.sort((a, b) => a.offset - b.offset)) {
    const previous = merged.at(-1);
    if (
      previous &&
      Math.abs(previous.offset - hint.offset) <= BREAK_OFFSET_EPSILON
    ) {
      previous.lineAdvance = Math.max(previous.lineAdvance, hint.lineAdvance);
      continue;
    }
    merged.push({ ...hint });
  }
  return merged;
};

type ContinuationCell = {
  geometry: CellBreakGeometry;
  contentBottom: number;
};

const continuationSkipForCell = (
  { geometry, contentBottom }: ContinuationCell,
  offset: number,
): number | undefined => {
  if (contentBottom <= offset + BREAK_OFFSET_EPSILON) {
    return undefined;
  }

  const leadingRange = geometry.suppressibleLeadingRanges.find(
    ({ top, bottom }) =>
      top <= offset + BREAK_OFFSET_EPSILON &&
      bottom > offset + BREAK_OFFSET_EPSILON,
  );
  return leadingRange ? leadingRange.bottom - offset : 0;
};

/**
 * Leading spacing before an empty paragraph is consumed when a split row
 * resumes at the top of a fresh flow region. Only a whitespace band shared by
 * every cell with remaining content is suppressible; non-empty paragraphs,
 * line boxes, cell padding, and explicit row height retain authored space.
 */
const continuationSkipAfter = (
  cells: ContinuationCell[],
  offset: number,
): number => {
  const skips = cells
    .map((cell) => continuationSkipForCell(cell, offset))
    .filter((skip): skip is number => skip !== undefined);
  if (
    skips.length === 0 ||
    skips.some((skip) => skip <= BREAK_OFFSET_EPSILON)
  ) {
    return 0;
  }
  return Math.min(...skips);
};

/** Precomputed break geometry for a table. */
export type TableRowBreakInfo = {
  /** Cumulative y of the top of each row; `rowTops[rows.length]` is the table height. */
  rowTops: number[];
  /**
   * Per-row sorted, de-duplicated line-bottom offsets (relative to the row top)
   * at which a break is clean. Always includes the row's full height as the
   * final boundary.
   */
  breakOffsets: number[][];
  /**
   * `breakOffsets` plus the line bottoms inside nested rows that stay whole
   * while a flow region can hold them: where to break once none can.
   */
  oversizedBreakOffsets: number[][];
  /** Safe cached page boundaries within each row, relative to the row top. */
  renderedBreakHints: RenderedRowBreakHint[][];
  /** Suppressible leading paragraph whitespace after each matching break offset. */
  continuationSkips: number[][];
};

/**
 * Height of the start of a table that a preceding `w:keepNext` paragraph
 * (§17.3.1.15) must share a page with: the leading header rows (§17.4.49) and
 * the first body row. A row that may break across pages (no `w:cantSplit`,
 * §17.4.6) needs only its first line there; an unbreakable row needs all of it.
 */
export function tableKeepNextOpeningHeight(
  block: TableBlock,
  measure: TableMeasure,
): number {
  let height = 0;
  let rowIndex = 0;
  while (
    rowIndex < measure.rows.length &&
    block.rows[rowIndex]?.isHeader === true
  ) {
    height += measure.rows[rowIndex]?.height ?? 0;
    rowIndex += 1;
  }
  const row = measure.rows[rowIndex];
  const sourceRow = block.rows[rowIndex];
  if (!row || !sourceRow) {
    return height;
  }
  const breakOffsets =
    buildTableRowBreakInfo(block, measure).breakOffsets[rowIndex] ?? [];
  const firstBreak = breakOffsets.at(0);
  const rowMayBreak = sourceRow.cantSplit !== true && breakOffsets.length > 1;
  return (
    height + (rowMayBreak && firstBreak !== undefined ? firstBreak : row.height)
  );
}

/** Build break geometry for a table from its block + measure. */
export function buildTableRowBreakInfo(
  block: TableBlock,
  measure: TableMeasure,
): TableRowBreakInfo {
  const rowCount = measure.rows.length;
  const rowTops: number[] = [];
  let acc = 0;
  for (let r = 0; r < rowCount; r++) {
    rowTops.push(acc);
    acc += measure.rows[r]?.height ?? 0;
  }
  rowTops.push(acc);

  const breakOffsets: number[][] = [];
  const oversizedBreakOffsets: number[][] = [];
  const renderedBreakHints: RenderedRowBreakHint[][] = [];
  const continuationSkips: number[][] = [];
  for (let r = 0; r < rowCount; r++) {
    const rowHeight = measure.rows[r]?.height ?? 0;
    const cellGeometries = rowCellGeometries(block, measure, r);
    const offsets = new Set<number>();
    offsets.add(rowHeight); // a row boundary is always a clean break
    for (const geometry of cellGeometries) {
      for (const b of geometry.bottoms) {
        if (b > 0 && b < rowHeight) {
          offsets.add(b);
        }
      }
    }
    const unsafeRanges = mergeBreakRanges(
      cellGeometries.flatMap((geometry) => geometry.unsafeRanges),
    );
    const keepTogetherRanges = mergeBreakRanges(
      cellGeometries.flatMap((geometry) => geometry.keepTogetherRanges),
    );
    const lineSafeOffsets = [...offsets]
      .filter(
        (offset) =>
          offset === rowHeight || !isInsideMergedRanges(offset, unsafeRanges),
      )
      .sort((a, b) => a - b);
    const sortedOffsets = lineSafeOffsets.filter(
      (offset) =>
        offset === rowHeight ||
        !isInsideMergedRanges(offset, keepTogetherRanges),
    );
    breakOffsets.push(sortedOffsets);
    oversizedBreakOffsets.push(lineSafeOffsets);
    renderedBreakHints.push(
      mergeRenderedBreakHints(
        cellGeometries
          .flatMap((geometry) => geometry.renderedBreakHints)
          .filter(
            ({ offset }) =>
              offset > BREAK_OFFSET_EPSILON &&
              offset < rowHeight - BREAK_OFFSET_EPSILON &&
              !isInsideMergedRanges(offset, unsafeRanges) &&
              !isInsideMergedRanges(offset, keepTogetherRanges),
          ),
      ),
    );
    const continuationCells = cellGeometries.map((geometry) => {
      let contentBottom = Number.NEGATIVE_INFINITY;
      for (const range of geometry.unsafeRanges) {
        contentBottom = Math.max(contentBottom, range.bottom);
      }
      return { geometry, contentBottom };
    });
    continuationSkips.push(
      sortedOffsets.map((offset) =>
        continuationSkipAfter(continuationCells, offset),
      ),
    );
  }

  return {
    rowTops,
    breakOffsets,
    oversizedBreakOffsets,
    renderedBreakHints,
    continuationSkips,
  };
}

/** Each cell's whole lines and atomic blocks, relative to the row top. */
export function rowCellLineBands(
  block: TableBlock,
  measure: TableMeasure,
  rowIndex: number,
): { top: number; bottom: number }[][] {
  return rowCellGeometries(block, measure, rowIndex).map((geometry) =>
    mergeBreakRanges(geometry.unsafeRanges),
  );
}

type PerCellRowCutOptions = {
  cellLines: readonly (readonly { top: number; bottom: number }[])[];
  bandTop: number;
  shownThrough: number;
  maxSlice: number;
};

/**
 * Cut a row that no shared line boundary lets a page hold. The cut itself runs
 * through some cell's line, which then waits for the next band, so that band
 * starts at the highest such line. Null when no line would finish on this page,
 * or the next band would not start lower than this one, which is a line taller
 * than the page: only clipping can place that.
 */
export function perCellRowCut({
  cellLines,
  bandTop,
  shownThrough,
  maxSlice,
}: PerCellRowCutOptions): { cut: number; nextTop: number } | null {
  const cut = bandTop + maxSlice;
  let finishesLine = false;
  let nextTop = cut;
  for (const lines of cellLines) {
    for (const line of lines) {
      if (line.bottom <= shownThrough + BREAK_OFFSET_EPSILON) {
        continue;
      }
      if (line.bottom <= cut + BREAK_OFFSET_EPSILON) {
        finishesLine = true;
        continue;
      }
      nextTop = Math.min(nextTop, line.top);
      break;
    }
  }
  // A line starting at or above this band, and running past the cut, would send
  // the next band back here: only clipping can place it.
  return finishesLine && nextTop > bandTop + BREAK_OFFSET_EPSILON
    ? { cut, nextTop }
    : null;
}

/** The part of one cell a `cellLineBand` shows, in the row's own coordinates. */
export function cellBandFromLines(
  lines: readonly { top: number; bottom: number }[],
  { shownThrough, cut }: { shownThrough: number; cut: number },
): { top: number; bottom: number } {
  const shown = lines.filter(
    ({ bottom }) =>
      bottom > shownThrough + BREAK_OFFSET_EPSILON &&
      bottom <= cut + BREAK_OFFSET_EPSILON,
  );
  const first = shown.at(0);
  const last = shown.at(-1);
  return first && last
    ? { top: first.top, bottom: last.bottom }
    : { top: cut, bottom: cut };
}

/** Leading empty-paragraph whitespace to consume when a row resumes after `offset`. */
export function getRowContinuationSkip(
  info: TableRowBreakInfo,
  rowIndex: number,
  offset: number,
): number {
  const offsets = info.breakOffsets[rowIndex];
  const skips = info.continuationSkips[rowIndex];
  const index = offsets?.findIndex(
    (candidate) => Math.abs(candidate - offset) <= BREAK_OFFSET_EPSILON,
  );
  if (index === undefined || index < 0) {
    return 0;
  }
  return skips?.[index] ?? 0;
}

/**
 * Given a row and how much of it has already been placed (`fromOffset`), return
 * how many more px can be placed ending on a whole line, without exceeding
 * `maxSlice`. Returns 0 when not even the first line fits.
 */
export function snapRowBreak(
  info: TableRowBreakInfo,
  rowIndex: number,
  fromOffset: number,
  maxSlice: number,
): number {
  return snapToOffsets(info.breakOffsets[rowIndex], fromOffset, maxSlice);
}

/**
 * `snapRowBreak` for a region that starts empty and still cannot hold the next
 * unit: nested rows that stay whole give way so their content stays reachable.
 */
export function snapOversizedRowBreak(
  info: TableRowBreakInfo,
  rowIndex: number,
  fromOffset: number,
  maxSlice: number,
): number {
  return snapToOffsets(
    info.oversizedBreakOffsets[rowIndex],
    fromOffset,
    maxSlice,
  );
}

function snapToOffsets(
  offsets: number[] | undefined,
  fromOffset: number,
  maxSlice: number,
): number {
  if (!offsets || offsets.length === 0) {
    return 0;
  }
  const limit = fromOffset + maxSlice;
  let best = 0;
  for (const off of offsets) {
    if (off <= fromOffset) {
      continue;
    }
    if (off <= limit) {
      best = off - fromOffset;
    } else {
      break;
    }
  }
  return best;
}
