import { recordMeasureBlock, recordMeasureBlockError } from "../layoutInstrumentation";
import { isParagraphFrameTextBox } from "../paragraphFrame";
import {
  createTableCellFlowState,
  finishTableCellFlow,
  placeTableCellBlock,
} from "./tableCellFlow";
import { bandFragmentX, bandTopContentY, isPageFrameRelativeAnchor } from "../textBoxFlow";
import { getTextBoxGroupId } from "../textBoxGroup";
import {
  DEFAULT_TEXTBOX_MARGINS,
  floatingTextBoxReservesBand,
  floatingTextBoxWrapsText,
  tableColumnsArePinned,
} from "../types";
import type {
  FlowBlock,
  ImageBlock,
  ImageRun,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableCellMeasure,
  TableMeasure,
  TextBoxBlock,
  TextBoxMeasure,
} from "../types";
import { getCachedParagraphMeasure, setCachedParagraphMeasure } from "./cache";
import { findClearLineY, measureParagraph, MIN_WRAP_SEGMENT_WIDTH } from "./measureParagraph";
import type { FloatingImageZone } from "./measureParagraph";
import { resolveFloatingTableX } from "./floatingTablePosition";
import {
  buildTableCellGrid,
  getFirstAvailableColumn,
  getTableCellVerticalBorderHeight,
} from "./tableCellGrid";
import { layoutTextBoxContent } from "./textBoxParagraphLayout";

/**
 * Pseudo-infinite measurement width (px) used for `w:noWrap` table cells so
 * the paragraph line breaker never inserts a soft break. Large enough to
 * exceed any realistic single Word line; small enough to stay well clear of
 * floating-point precision concerns when summed downstream.
 */
const NO_WRAP_MEASURE_WIDTH = 1_000_000;

/**
 * Sanity cap on the table column count derived from summed cell colSpans
 * when a table has no `tblGrid`/explicit column widths. Mirrors the
 * `MAX_TABLE_COLUMNS` clamp applied to `w:gridSpan` at parse time
 * (`docx/tableParser.ts`); enforced again here since a hostile colSpan sum
 * would otherwise size the fallback column array and cell grid unbounded.
 */
const MAX_TABLE_COLUMNS = 63;

/**
 * Check if an image run is a *text-wrapping* floating image — it
 * occupies an exclusion zone the body text should flow around.
 *
 * `wrapType: "behind"` and `wrapType: "inFront"` are anchored
 * (out-of-flow) but Word's wrapNone semantics put them behind /
 * over the text without shrinking the body. They render as
 * `displayMode: "float"` in the prose model so the painter knows
 * they're out of normal flow, but `extractFloatingZones` must skip
 * them — including them here would make the line breaker wrap text
 * around a background letterhead or a foreground overlay (Codex
 * PR #258 review).
 */
function isSideWrappingImageRun(run: ImageRun): boolean {
  const wrapType = run.wrapType;
  const displayMode = run.displayMode;

  // wrapNone (behind / inFront): never an exclusion zone, regardless
  // of displayMode.
  if (wrapType === "behind" || wrapType === "inFront") {
    return false;
  }

  // Floating images have specific wrap types that allow text to flow around them
  if (wrapType && ["square", "tight", "through"].includes(wrapType)) {
    return true;
  }

  // Or explicit float display mode (only when no wrapNone semantics —
  // already filtered above).
  if (displayMode === "float") {
    return true;
  }

  return false;
}

/**
 * EMU to pixels conversion
 */
function emuToPixels(emu: number | undefined): number {
  if (emu === undefined) {
    return 0;
  }
  return Math.round((emu * 96) / 914_400);
}

function resolveTableWidthPx(
  width: number | undefined,
  widthType: string | undefined,
  contentWidth: number,
): number | undefined {
  if (!width) {
    return undefined;
  }
  if (widthType === "pct") {
    // width is in 50ths of a percent (5000 = 100%)
    return (contentWidth * width) / 5000;
  }
  if (widthType === "dxa" || !widthType || widthType === "auto") {
    return Math.round((width / 20) * 1.333);
  }
  return undefined;
}

export function measureTableBlock(
  tableBlock: TableBlock,
  contentWidth: number,
  fieldValues?: ReadonlyMap<number, string>,
): TableMeasure {
  const DEFAULT_CELL_PADDING_X = 7; // Word default: 108 twips ≈ 7px
  const DEFAULT_CELL_PADDING_Y = 0; // OOXML/TableNormal default: top=0, bottom=0

  // columnWidths are already in pixels (converted in toFlowBlocks)
  let columnWidths = tableBlock.columnWidths ?? [];
  const explicitWidthPx = resolveTableWidthPx(tableBlock.width, tableBlock.widthType, contentWidth);

  if (columnWidths.length === 0 && tableBlock.rows.length > 0) {
    // Determine total columns from first row's colSpans
    const colCount = Math.min(
      tableBlock.rows[0]!.cells.reduce(
        // SAFETY: rows.length > 0
        (sum, cell) => sum + (cell.colSpan ?? 1),
        0,
      ),
      MAX_TABLE_COLUMNS,
    );
    const totalWidth = explicitWidthPx ?? contentWidth;
    const equalWidth = totalWidth / Math.max(1, colCount);
    columnWidths = Array.from({ length: colCount }, () => equalWidth);
  } else if (columnWidths.length > 0 && explicitWidthPx) {
    const totalWidth = columnWidths.reduce((sum, w) => sum + w, 0);
    // `w:tblW` is a preferred width under autofit. When the imported grid is
    // already wider, Word preserves that grid instead of scaling it down.
    const preserveExpandedGrid =
      tableBlock.layout !== "fixed" &&
      (tableBlock.widthType === undefined || tableBlock.widthType === "dxa") &&
      totalWidth - explicitWidthPx > 1;
    if (!preserveExpandedGrid && totalWidth > 0 && Math.abs(totalWidth - explicitWidthPx) > 1) {
      const scale = explicitWidthPx / totalWidth;
      columnWidths = columnWidths.map((w) => w * scale);
    }
  }

  const cellGrid = buildTableCellGrid(tableBlock.rows, columnWidths.length);

  // When the columns are pinned (fixed layout or a fully-consumed explicit
  // width), Word cannot widen a column to satisfy `w:noWrap`, so overflowing
  // content wraps. Only an auto-width table honors `w:noWrap` by keeping the
  // cell on one line. The painter reads the same predicate for `white-space`.
  const columnsPinned = tableColumnsArePinned(tableBlock);

  // Calculate cell widths based on colSpan and columnWidths,
  // skipping columns occupied by spanning cells from previous rows.
  const rows = tableBlock.rows.map((row, rowIdx) => {
    let columnIndex = getFirstAvailableColumn(cellGrid, rowIdx, row.gridBefore ?? 0);

    return {
      cells: row.cells.map((cell) => {
        const colSpan = cell.colSpan ?? 1;
        // Calculate cell width as sum of spanned columns
        let cellWidth = 0;
        for (let c = 0; c < colSpan && columnIndex + c < columnWidths.length; c++) {
          cellWidth += columnWidths[columnIndex + c] ?? 0;
        }
        // Fallback to cell.width or default if columnWidths not available
        if (cellWidth === 0) {
          cellWidth = cell.width ?? 100;
        }
        columnIndex = getFirstAvailableColumn(cellGrid, rowIdx, columnIndex + colSpan);

        const padLeft = cell.padding?.left ?? DEFAULT_CELL_PADDING_X;
        const padRight = cell.padding?.right ?? DEFAULT_CELL_PADDING_X;
        const cellContentWidth = Math.max(1, cellWidth - padLeft - padRight);
        // `w:noWrap` (§17.4.30): in an auto-width table Word honors it by
        // widening the column, so measure against an effectively unbounded
        // width and the line breaker keeps the cell on one MeasuredLine (the
        // painter pairs this with `white-space: nowrap`). When the columns are
        // pinned Word cannot widen, so overflowing content wraps: measure at the
        // real content width like an ordinary cell.
        const keepSingleLine = cell.noWrap === true && !columnsPinned;
        const measureWidth = keepSingleLine ? NO_WRAP_MEASURE_WIDTH : cellContentWidth;
        const cellMeasure: TableCellMeasure = {
          blocks: cell.blocks.map((b) =>
            measureBlock(b, measureWidth, undefined, undefined, fieldValues),
          ),
          width: cellWidth,
          height: 0, // Calculated below
        };
        if (cell.colSpan !== undefined) {
          cellMeasure.colSpan = cell.colSpan;
        }
        if (cell.rowSpan !== undefined) {
          cellMeasure.rowSpan = cell.rowSpan;
        }
        return cellMeasure;
      }),
      height: 0,
    };
  });

  // Calculate cell heights, respecting explicit row height rules
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const row = rows[rowIdx]!; // SAFETY: rowIdx < rows.length
    const sourceRow = tableBlock.rows[rowIdx];
    if (sourceRow?.hidden) {
      row.height = 0;
      row.cells = sourceRow.cells.map(() => ({
        blocks: [],
        width: 0,
        height: 0,
      }));
      continue;
    }
    const sourceRowCells = tableBlock.rows[rowIdx]?.cells;
    // Take the max over per-cell totals (content + padding + vertical borders),
    // not the sum of an independent content-max and border-max: those two maxes
    // can come from different cells, which over-allocates the row when the
    // tallest-content cell is not also the tallest-border cell. `cell.height`
    // stays content + padding (the painter lays cell content out within it); the
    // border only contributes to the row total here.
    let maxCellHeightWithBorders = 0;
    let maxCellInsets = 0;
    for (let cellIdx = 0; cellIdx < row.cells.length; cellIdx++) {
      const cell = row.cells[cellIdx]!; // SAFETY: cellIdx < row.cells.length
      const sourceCell = sourceRowCells?.[cellIdx];
      cell.height = 0;
      const flowState = createTableCellFlowState();
      for (let blockIdx = 0; blockIdx < cell.blocks.length; blockIdx++) {
        const sourceBlock = sourceCell?.blocks[blockIdx];
        const blockMeasure = cell.blocks[blockIdx];
        if (!sourceBlock || !blockMeasure) {
          continue;
        }
        placeTableCellBlock(flowState, sourceBlock, blockMeasure);
      }
      cell.height = finishTableCellFlow(flowState);
      const padTop = sourceCell?.padding?.top ?? DEFAULT_CELL_PADDING_Y;
      const padBottom = sourceCell?.padding?.bottom ?? DEFAULT_CELL_PADDING_Y;
      cell.height += padTop + padBottom;
      if ((sourceCell?.rowSpan ?? 1) > 1) {
        continue;
      }
      const borderHeight = getTableCellVerticalBorderHeight(cellGrid, sourceCell, rowIdx);
      maxCellHeightWithBorders = Math.max(maxCellHeightWithBorders, cell.height + borderHeight);
      maxCellInsets = Math.max(maxCellInsets, padTop + padBottom + borderHeight);
    }

    // Apply heightRule from the source row
    const explicitHeight = sourceRow?.height;
    const heightRule = sourceRow?.heightRule;

    if (explicitHeight && heightRule === "exact") {
      row.height = explicitHeight;
    } else if (explicitHeight) {
      // Compatibility layouts treat both 'atLeast' and auto-with-val as a
      // content floor; cell padding and horizontal borders remain outside it.
      row.height = Math.max(maxCellHeightWithBorders, explicitHeight + maxCellInsets);
    } else {
      // No explicit height — use content height directly.
      row.height = maxCellHeightWithBorders;
    }
  }

  // A vertically merged cell occupies the combined height of all rows it
  // spans. Its content must constrain that combined area, not inflate the
  // first row independently. Apply any remaining deficit to the last
  // non-exact row in the span after every row has its own base height.
  for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
    const sourceRow = tableBlock.rows[rowIdx];
    const measuredRow = rows[rowIdx];
    if (!sourceRow || !measuredRow) {
      continue;
    }
    for (let cellIdx = 0; cellIdx < sourceRow.cells.length; cellIdx++) {
      const sourceCell = sourceRow.cells[cellIdx];
      const measuredCell = measuredRow.cells[cellIdx];
      const rowSpan = sourceCell?.rowSpan ?? 1;
      if (!sourceCell || !measuredCell || rowSpan <= 1) {
        continue;
      }
      const spanEnd = Math.min(rows.length, rowIdx + rowSpan);
      let combinedHeight = 0;
      for (let spannedRowIdx = rowIdx; spannedRowIdx < spanEnd; spannedRowIdx++) {
        combinedHeight += rows[spannedRowIdx]?.height ?? 0;
      }
      const requiredHeight =
        measuredCell.height + getTableCellVerticalBorderHeight(cellGrid, sourceCell, rowIdx);
      const deficit = requiredHeight - combinedHeight;
      if (deficit <= 0) {
        continue;
      }
      for (let spannedRowIdx = spanEnd - 1; spannedRowIdx >= rowIdx; spannedRowIdx--) {
        if (tableBlock.rows[spannedRowIdx]?.heightRule === "exact") {
          continue;
        }
        rows[spannedRowIdx]!.height += deficit;
        break;
      }
    }
  }

  const totalHeight = rows.reduce((h, r) => h + r.height, 0);
  const totalWidth = columnWidths.reduce((w, cw) => w + cw, 0) || explicitWidthPx || contentWidth;

  return {
    kind: "table",
    rows,
    columnWidths,
    totalWidth,
    totalHeight,
  };
}

/**
 * Extract floating image exclusion zones from all blocks.
 * Called before measurement to determine line width reductions.
 *
 * For images with vertical align="top" relative to margin, they're at Y=0.
 * The exclusion zones define the areas where text lines need reduced widths.
 */
/**
 * Extended floating zone info that includes anchor block index
 */
type FloatingZoneWithAnchor = {
  /** Block index where this floating image is anchored */
  anchorBlockIndex: number;
  /** If true, zone is positioned relative to margin/page and applies to all blocks */
  isMarginRelative?: boolean;
} & FloatingImageZone;

function perBlockNumberValue(
  value: number | number[],
  blockIndex: number,
  fallback: number,
): number {
  if (Array.isArray(value)) {
    return value[blockIndex] ?? fallback;
  }
  return value;
}

type TextBoxWrapSideOptions = {
  box: TextBoxBlock;
  contentX: number;
  boxWidth: number;
  contentWidth: number;
};

function textBoxWrapSide({
  box,
  contentX,
  boxWidth,
  contentWidth,
}: TextBoxWrapSideOptions): "left" | "right" {
  if (box.wrapText === "left") {
    return "left";
  }
  if (box.wrapText === "right") {
    return "right";
  }

  const leftSpace = contentX;
  const rightSpace = contentWidth - contentX - boxWidth;
  if (box.wrapText === "largest") {
    return leftSpace >= rightSpace ? "left" : "right";
  }
  return contentX + boxWidth / 2 < contentWidth / 2 ? "right" : "left";
}

// Page geometry the band extraction needs to resolve page/margin-pinned
// topAndBottom anchors (bottom-strip frames, centered/bottom align). Per-block
// because sections can vary page size and margins. eigenpal #694.
type BandPageGeometry = {
  pageWidth?: number | number[];
  pageHeight: number | number[];
  marginLeft?: number | number[];
  marginRight?: number | number[];
  marginBottom: number | number[];
};

function extractFloatingZones(
  blocks: FlowBlock[],
  contentWidth: number | number[],
  marginTop: number | number[] = 0,
  pageGeometry?: BandPageGeometry,
): FloatingZoneWithAnchor[] {
  const zones: FloatingZoneWithAnchor[] = [];
  const defaultContentWidth = Array.isArray(contentWidth) ? (contentWidth[0] ?? 0) : contentWidth;
  const textBoxGroupAnchors = new Map<string, number>();
  const paragraphFrameGroupSizes = new Map<string, number>();
  for (const block of blocks) {
    if (block.kind !== "textBox" || !isParagraphFrameTextBox(block)) {
      continue;
    }
    const groupId = getTextBoxGroupId(block);
    if (groupId !== undefined) {
      paragraphFrameGroupSizes.set(groupId, (paragraphFrameGroupSizes.get(groupId) ?? 0) + 1);
    }
  }
  const defaultMarginTop = Array.isArray(marginTop) ? (marginTop[0] ?? 0) : marginTop;
  const pageWidthInput = pageGeometry?.pageWidth ?? defaultContentWidth;
  const pageHeightInput = pageGeometry?.pageHeight ?? 0;
  const marginLeftInput = pageGeometry?.marginLeft ?? 0;
  const marginRightInput = pageGeometry?.marginRight ?? 0;
  const marginBottomInput = pageGeometry?.marginBottom ?? 0;
  const defaultPageWidth = Array.isArray(pageWidthInput)
    ? (pageWidthInput[0] ?? defaultContentWidth)
    : pageWidthInput;
  const defaultPageHeight = Array.isArray(pageHeightInput)
    ? (pageHeightInput[0] ?? 0)
    : pageHeightInput;
  const defaultMarginLeft = Array.isArray(marginLeftInput)
    ? (marginLeftInput[0] ?? 0)
    : marginLeftInput;
  const defaultMarginRight = Array.isArray(marginRightInput)
    ? (marginRightInput[0] ?? 0)
    : marginRightInput;
  const defaultMarginBottom = Array.isArray(marginBottomInput)
    ? (marginBottomInput[0] ?? 0)
    : marginBottomInput;

  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex]!; // SAFETY: blockIndex < blocks.length
    if (block.kind !== "paragraph") {
      continue;
    }

    const paragraphBlock = block as ParagraphBlock;

    for (const run of paragraphBlock.runs) {
      if (run.kind !== "image") {
        continue;
      }
      const imgRun = run as ImageRun;

      const positionedBand = imgRun.wrapType === "topAndBottom" && imgRun.position !== undefined;
      if (!isSideWrappingImageRun(imgRun) && !positionedBand) {
        continue;
      }

      // Calculate Y position based on vertical alignment
      let topY = 0;
      const position = imgRun.position;
      const distTop = imgRun.distTop ?? 0;
      const distBottom = imgRun.distBottom ?? 0;
      const distLeft = imgRun.distLeft ?? 12;
      const distRight = imgRun.distRight ?? 12;

      if (positionedBand && position !== undefined) {
        const vertical = position.vertical;
        const blockMarginTop = perBlockNumberValue(marginTop, blockIndex, defaultMarginTop);
        const pageFrameRelative = isPageFrameRelativeAnchor(vertical?.relativeTo);
        const rawTop = pageFrameRelative
          ? bandTopContentY(vertical, {
              pageHeight: perBlockNumberValue(pageHeightInput, blockIndex, defaultPageHeight),
              marginTop: blockMarginTop,
              marginBottom: perBlockNumberValue(marginBottomInput, blockIndex, defaultMarginBottom),
              boxHeight: imgRun.height,
            })
          : emuToPixels(vertical?.posOffset ?? 0);
        const bottomY = rawTop + imgRun.height + distBottom;
        if (bottomY > 0) {
          zones.push({
            leftMargin: 0,
            rightMargin: 0,
            topY: Math.max(0, rawTop - distTop),
            bottomY,
            anchorBlockIndex: blockIndex,
            ...(pageFrameRelative ? { isMarginRelative: true } : {}),
            fullWidthBlock: true,
          });
        }
        continue;
      }

      if (position?.vertical) {
        const v = position.vertical;
        if (v.align === "top" && v.relativeTo === "margin") {
          // Image at top of content area
          topY = 0;
        } else if (v.posOffset !== undefined) {
          topY = emuToPixels(v.posOffset);
        }
        // Other cases (paragraph-relative) are harder to handle without knowing paragraph positions
      }

      const bottomY = topY + imgRun.height;

      // Calculate margins based on horizontal position
      let leftMargin = 0;
      let rightMargin = 0;

      if (position?.horizontal) {
        const h = position.horizontal;
        if (h.align === "left") {
          // Image on left - text needs left margin
          leftMargin = imgRun.width + distRight;
        } else if (h.align === "right") {
          // Image on right - text needs right margin
          rightMargin = imgRun.width + distLeft;
        } else if (h.posOffset !== undefined) {
          const x = emuToPixels(h.posOffset);
          if (x < defaultContentWidth / 2) {
            leftMargin = x + imgRun.width + distRight;
          } else {
            rightMargin = defaultContentWidth - x + distLeft;
          }
        }
      } else if (imgRun.cssFloat === "left") {
        leftMargin = imgRun.width + distRight;
      } else if (imgRun.cssFloat === "right") {
        rightMargin = imgRun.width + distLeft;
      }

      if (leftMargin > 0 || rightMargin > 0) {
        // Images positioned relative to margin/page apply globally (before their anchor paragraph)
        const isMarginRelative =
          position?.vertical?.relativeTo === "margin" || position?.vertical?.relativeTo === "page";
        zones.push({
          leftMargin,
          rightMargin,
          topY: topY - distTop,
          bottomY: bottomY + distBottom,
          anchorBlockIndex: blockIndex,
          isMarginRelative,
        });
      }
    }
  }

  // Floating tables (block-level) - treat them as exclusion zones for subsequent text
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex]!; // SAFETY: blockIndex < blocks.length
    if (block.kind !== "table") {
      continue;
    }

    const tableBlock = block as TableBlock;
    const floating = tableBlock.floating;
    if (!floating) {
      continue;
    }

    const tableMeasure = measureTableBlock(tableBlock, defaultContentWidth);
    const tableWidth = tableMeasure.totalWidth;
    const tableHeight = tableMeasure.totalHeight;

    const distLeft = floating.leftFromText ?? 12;
    const distRight = floating.rightFromText ?? 12;
    const distTop = floating.topFromText ?? 0;
    const distBottom = floating.bottomFromText ?? 0;

    let leftMargin = 0;
    let rightMargin = 0;

    // Determine horizontal position relative to content area
    const x = resolveFloatingTableX(
      floating,
      tableBlock.justification,
      tableWidth,
      defaultContentWidth,
    );

    if (x < defaultContentWidth / 2) {
      leftMargin = x + tableWidth + distRight;
    } else {
      rightMargin = defaultContentWidth - x + distLeft;
    }

    const topY = floating.tblpY ?? 0;
    const bottomY = topY + tableHeight;

    zones.push({
      leftMargin,
      rightMargin,
      topY: topY - distTop,
      bottomY: bottomY + distBottom,
      anchorBlockIndex: blockIndex,
    });
  }

  // Positioned side-wrapped text boxes carve the same horizontal exclusion
  // during pagination that the painter applies during its final re-measure.
  // Without this pre-scan, paragraphs paint around the box but retain their
  // original page assignment, so a page-anchored frame can overlap later body
  // content or create a false extra page.
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex]!; // SAFETY: blockIndex < blocks.length
    if (block.kind !== "textBox") {
      continue;
    }
    const tb = block as TextBoxBlock;
    if (!floatingTextBoxWrapsText(tb) || tb.position === undefined) {
      continue;
    }

    const measure = measureTextBoxBlock(tb);
    const blockMarginTop = perBlockNumberValue(marginTop, blockIndex, defaultMarginTop);
    const blockMarginLeft = perBlockNumberValue(marginLeftInput, blockIndex, defaultMarginLeft);
    const blockMarginRight = perBlockNumberValue(marginRightInput, blockIndex, defaultMarginRight);
    const blockPageWidth = perBlockNumberValue(pageWidthInput, blockIndex, defaultPageWidth);
    const blockContentWidth = perBlockNumberValue(contentWidth, blockIndex, defaultContentWidth);
    const vertical = tb.position.vertical;
    const pageFrameRelative = isPageFrameRelativeAnchor(vertical?.relativeTo);
    const topY = pageFrameRelative
      ? bandTopContentY(vertical, {
          pageHeight: perBlockNumberValue(pageHeightInput, blockIndex, defaultPageHeight),
          marginTop: blockMarginTop,
          marginBottom: perBlockNumberValue(marginBottomInput, blockIndex, defaultMarginBottom),
          boxHeight: measure.height,
        })
      : emuToPixels(vertical?.posOffset ?? 0);
    const groupId = getTextBoxGroupId(tb);
    const anchorBlockIndex = groupId
      ? (textBoxGroupAnchors.get(groupId) ?? blockIndex)
      : blockIndex;
    if (groupId && !textBoxGroupAnchors.has(groupId)) {
      textBoxGroupAnchors.set(groupId, blockIndex);
    }
    if (
      isParagraphFrameTextBox(tb) &&
      groupId !== undefined &&
      (paragraphFrameGroupSizes.get(groupId) ?? 0) > 1
    ) {
      zones.push({
        leftMargin: 0,
        rightMargin: 0,
        topY: 0,
        bottomY: topY + measure.height + (tb.distBottom ?? 0),
        anchorBlockIndex,
        ...(pageFrameRelative ? { isMarginRelative: true } : {}),
        fullWidthBlock: true,
      });
      continue;
    }

    const horizontal = tb.position.horizontal;
    const pageX = horizontal
      ? bandFragmentX(horizontal, {
          pageWidth: blockPageWidth,
          marginLeft: blockMarginLeft,
          marginRight: blockMarginRight,
          boxWidth: measure.width,
        })
      : blockMarginLeft;
    const contentX = pageX - blockMarginLeft;
    const wrapSide = textBoxWrapSide({
      box: tb,
      contentX,
      boxWidth: measure.width,
      contentWidth: blockContentWidth,
    });
    const leftMargin = wrapSide === "right" ? contentX + measure.width + (tb.distRight ?? 12) : 0;
    const rightMargin =
      wrapSide === "left" ? blockContentWidth - contentX + (tb.distLeft ?? 12) : 0;
    zones.push({
      leftMargin: Math.max(0, leftMargin),
      rightMargin: Math.max(0, rightMargin),
      topY: topY - (tb.distTop ?? 0),
      bottomY: topY + measure.height + (tb.distBottom ?? 0),
      anchorBlockIndex,
      ...(pageFrameRelative ? { isMarginRelative: true } : {}),
    });
  }

  // Positioned topAndBottom text boxes reserve a full-width band so body text
  // flows above and below the anchored shape. Multiple boxes extracted from
  // one shape-only paragraph share an anchor and remain active together.
  for (let blockIndex = 0; blockIndex < blocks.length; blockIndex++) {
    const block = blocks[blockIndex]!; // SAFETY: blockIndex < blocks.length
    if (block.kind !== "textBox") {
      continue;
    }
    const tb = block as TextBoxBlock;
    if (!floatingTextBoxReservesBand(tb) || tb.position === undefined) {
      continue;
    }
    const v = tb.position?.vertical;

    const height = measureTextBoxBlock(tb).height;
    const distTop = tb.distTop ?? 0;
    const distBottom = tb.distBottom ?? 0;
    const blockMarginTop = perBlockNumberValue(marginTop, blockIndex, defaultMarginTop);
    const pageFrameRelative = isPageFrameRelativeAnchor(v?.relativeTo);
    // Shared with layoutTextBox so the reserved band and the painted box agree.
    const rawTop = pageFrameRelative
      ? bandTopContentY(v, {
          pageHeight: perBlockNumberValue(pageHeightInput, blockIndex, defaultPageHeight),
          marginTop: blockMarginTop,
          marginBottom: perBlockNumberValue(marginBottomInput, blockIndex, defaultMarginBottom),
          boxHeight: height,
        })
      : emuToPixels(v?.posOffset ?? 0);
    const bottomY = rawTop + height + distBottom;
    if (bottomY <= 0) {
      continue;
    }
    const groupId = getTextBoxGroupId(tb);
    const anchorBlockIndex = groupId
      ? (textBoxGroupAnchors.get(groupId) ?? blockIndex)
      : blockIndex;
    if (groupId && !textBoxGroupAnchors.has(groupId)) {
      textBoxGroupAnchors.set(groupId, blockIndex);
    }
    zones.push({
      leftMargin: 0,
      rightMargin: 0,
      topY: Math.max(0, rawTop - distTop),
      bottomY,
      anchorBlockIndex,
      ...(pageFrameRelative ? { isMarginRelative: true } : {}),
      fullWidthBlock: true,
    });
  }

  return zones;
}

/**
 * Measure a block based on its type.
 */
export function measureBlock(
  block: FlowBlock,
  contentWidth: number,
  floatingZones?: FloatingImageZone[],
  cumulativeY?: number,
  fieldValues?: ReadonlyMap<number, string>,
): Measure {
  switch (block.kind) {
    case "paragraph": {
      const pBlock = block as ParagraphBlock;

      // Cache paragraph measurements when no floating zones affect this block.
      // Safe because without floating zones the result depends only on content
      // and contentWidth (both captured in the cache key). When floating zones
      // ARE present, we always measure fresh since zones depend on inter-block
      // layout context (cumulative Y, neighboring floating tables/images).
      // Resolved field values change measured field width and are not part of
      // the cache key. Only paragraphs with field runs need a fresh measure in
      // the stabilization pass; field-free paragraphs stay cacheable.
      const hasFieldRuns = pBlock.runs.some((run) => run.kind === "field");
      const cacheable =
        (!floatingZones || floatingZones.length === 0) && (!fieldValues || !hasFieldRuns);
      if (cacheable) {
        const cached = getCachedParagraphMeasure(pBlock, contentWidth);
        if (cached) {
          return cached;
        }
      }

      const measureOpts: Parameters<typeof measureParagraph>[2] = {
        paragraphYOffset: cumulativeY ?? 0,
      };
      if (floatingZones) {
        measureOpts.floatingZones = floatingZones;
      }
      if (fieldValues) {
        measureOpts.fieldValues = fieldValues;
      }
      const result = measureParagraph(pBlock, contentWidth, measureOpts);

      if (cacheable) {
        setCachedParagraphMeasure(pBlock, contentWidth, result);
      }

      return result;
    }

    case "table": {
      return measureTableBlock(block as TableBlock, contentWidth, fieldValues);
    }

    case "image": {
      const imageBlock = block as ImageBlock;
      return {
        kind: "image",
        width: imageBlock.width,
        height: imageBlock.height,
      };
    }

    case "textBox": {
      return measureTextBoxBlock(block as TextBoxBlock, fieldValues);
    }

    case "pageBreak":
      return { kind: "pageBreak" };

    case "columnBreak":
      return { kind: "columnBreak" };

    case "sectionBreak":
      return { kind: "sectionBreak" };

    default:
      // Unknown block type - return empty paragraph measure
      return {
        kind: "paragraph",
        lines: [],
        totalHeight: 0,
      };
  }
}

export function measureTextBoxBlock(
  tb: TextBoxBlock,
  fieldValues?: ReadonlyMap<number, string>,
): TextBoxMeasure {
  const margins = tb.margins ?? DEFAULT_TEXTBOX_MARGINS;
  const innerWidth = tb.width - margins.left - margins.right;
  const innerMeasures = tb.content.map((block) => {
    if (block.kind === "table") {
      return measureTableBlock(block, innerWidth, fieldValues);
    }
    return measureParagraph(block, innerWidth, fieldValues ? { fieldValues } : undefined);
  });
  const contentHeight = layoutTextBoxContent(tb.content, innerMeasures).totalHeight;
  const contentBoxHeight = contentHeight + margins.top + margins.bottom;
  const totalHeight =
    tb.autoFit === "shape"
      ? Math.max(tb.height ?? 0, contentBoxHeight)
      : (tb.height ?? contentBoxHeight);
  return {
    kind: "textBox",
    width: tb.width,
    height: totalHeight,
    innerMeasures,
  };
}

/**
 * `true` for a section break that explicitly starts a new page
 * (`nextPage`/`evenPage`/`oddPage`). An absent type follows
 * `scheduleSectionBreak` and stays continuous, so measurement must retain any
 * active exclusion zones across it. Used by `measureBlocks` to reset the
 * running Y for a page-pinned band that lands right after the break.
 * eigenpal #694.
 */
function isNextPageSectionBreak(block: FlowBlock): boolean {
  return block.kind === "sectionBreak" && block.type !== undefined && block.type !== "continuous";
}

/**
 * Measure all blocks with floating image support.
 *
 * Pre-scans all blocks to find floating images and creates exclusion zones.
 * Then measures each block, passing the zones so paragraphs can calculate
 * per-line widths based on vertical overlap with floating images.
 */
export function measureBlocks(
  blocks: FlowBlock[],
  contentWidth: number | number[],
  marginTop: number | number[] = 0,
  pageGeometry?: BandPageGeometry,
  fieldValues?: ReadonlyMap<number, string>,
): Measure[] {
  const defaultWidth = Array.isArray(contentWidth) ? (contentWidth[0] ?? 0) : contentWidth;
  // Pre-extract floating image exclusion zones with anchor block indices
  const floatingZonesWithAnchors = extractFloatingZones(
    blocks,
    contentWidth,
    marginTop,
    pageGeometry,
  );

  // Margin-relative zones (positioned relative to page/margin) on the same vertical
  // position are likely on the same page. Group them and activate all from the earliest
  // anchor so text wraps around ALL images from the first paragraph onward.
  // e.g. left-aligned and right-aligned images at margin top should both affect text
  // starting from the first anchor paragraph, not just the one containing each image.
  // Full-width topAndBottom bands are excluded: each pins to its own text box, so a
  // second band sharing the same topY (e.g. body banners in different sections) must
  // not be rewritten to the earliest anchor, or earlier pages would reserve a band
  // that is painted elsewhere. They keep their own anchor below. eigenpal #694.
  const marginRelative = floatingZonesWithAnchors.filter(
    (z) => z.isMarginRelative && !z.fullWidthBlock,
  );
  const ownAnchorZones = floatingZonesWithAnchors.filter(
    (z) => !z.isMarginRelative || z.fullWidthBlock,
  );

  // Group margin-relative zones by topY and move all to earliest anchor in group
  const marginByTopY = new Map<number, FloatingZoneWithAnchor[]>();
  for (const z of marginRelative) {
    const group = marginByTopY.get(z.topY) ?? [];
    group.push(z);
    marginByTopY.set(z.topY, group);
  }

  const adjustedZones: FloatingZoneWithAnchor[] = [...ownAnchorZones];
  for (const group of marginByTopY.values()) {
    const minAnchor = Math.min(...group.map((z) => z.anchorBlockIndex));
    for (const z of group) {
      adjustedZones.push({ ...z, anchorBlockIndex: minAnchor });
    }
  }

  // Group zones by effective anchor block index
  const zonesByAnchor = new Map<number, FloatingImageZone[]>();
  for (const z of adjustedZones) {
    const existing = zonesByAnchor.get(z.anchorBlockIndex) ?? [];
    // Strip the anchor-tracking fields; the rest IS a FloatingImageZone. Spread
    // (rather than copying each field) so fullWidthBlock/segments can't be
    // dropped here.
    const { anchorBlockIndex: _anchorBlockIndex, isMarginRelative: _isMarginRelative, ...zone } = z;
    existing.push(zone);
    zonesByAnchor.set(z.anchorBlockIndex, existing);
  }

  const anchorIndices = new Set(zonesByAnchor.keys());

  // Two running Y cursors for floating-zone overlap:
  //  - cumulativeY resets to 0 at each floating-image/table anchor, giving that
  //    object a local frame for its side-wrap zone.
  //  - pageRelativeY is the real page cursor; it resets only at hard page breaks
  //    (never at a float anchor), so a page-pinned topAndBottom band always
  //    measures from a true page-relative position even when a float was
  //    anchored earlier on the same page. eigenpal #694.
  let cumulativeY = 0;
  let pageRelativeY = 0;
  let activeZones: FloatingImageZone[] = [];

  return blocks.map((block, blockIndex) => {
    recordMeasureBlock(blockIndex, block);

    // A hard page/section break starts a fresh page. Any active zone — including
    // a page-pinned topAndBottom band — belongs to the page it was anchored on,
    // so drop the active zones and restart both cursors at the new page top. A
    // band anchor on the new page re-establishes its own zone below; without
    // this, the first block after the break would be measured against a stale
    // band (a phantom float-skip) while layout paints no band there, opening a
    // gap. eigenpal #694.
    if (block.kind === "pageBreak" || isNextPageSectionBreak(block)) {
      activeZones = [];
      cumulativeY = 0;
      pageRelativeY = 0;
    }

    // Check if this block is an anchor for floating images
    // If so, replace active zones (old zones from previous anchors are invalid
    // after a Y reset since their topY/bottomY are in the old coordinate system).
    if (anchorIndices.has(blockIndex)) {
      activeZones = zonesByAnchor.get(blockIndex) ?? [];
      // Floating-image anchors open a fresh local frame (cumulativeY → 0). A
      // page/margin-pinned band instead reserves against the page, so it
      // measures from pageRelativeY — the real page cursor, which a prior float
      // anchor has not reset. This is a no-op when no float precedes the band on
      // the page (the two cursors agree) and 0 right after a hard break, so the
      // band still reserves from the real cursor down rather than re-reserving
      // the whole band over content that already precedes its anchor. eigenpal #694.
      const bandOnlyAnchor = activeZones.length > 0 && activeZones.every((z) => z.fullWidthBlock);
      cumulativeY = bandOnlyAnchor ? pageRelativeY : 0;
    }

    const zones = activeZones.length > 0 ? activeZones : undefined;

    try {
      const blockWidth = Array.isArray(contentWidth)
        ? (contentWidth[blockIndex] ?? defaultWidth)
        : contentWidth;
      const measure = measureBlock(block, blockWidth, zones, cumulativeY, fieldValues);

      // Paragraphs clear floating exclusions internally (findClearLineY inside
      // measureParagraph). An in-flow table cannot reflow its cells around a
      // side exclusion, so require room for the whole table. Block images keep
      // the existing full-width-band behavior.
      const clearingZones =
        measure.kind === "table" ? zones : zones?.filter((zone) => zone.fullWidthBlock);
      if (
        clearingZones?.length &&
        (measure.kind === "image" || (measure.kind === "table" && !(block as TableBlock).floating))
      ) {
        const blockHeight = measure.kind === "image" ? measure.height : measure.totalHeight;
        const requiredWidth =
          measure.kind === "table"
            ? Math.min(blockWidth, measure.totalWidth)
            : MIN_WRAP_SEGMENT_WIDTH;
        const skip =
          findClearLineY(
            cumulativeY,
            blockHeight,
            clearingZones,
            blockWidth,
            Math.max(MIN_WRAP_SEGMENT_WIDTH, requiredWidth),
          ) - cumulativeY;
        if (skip > 0) {
          measure.bandSkipBefore = skip;
          cumulativeY += skip;
          pageRelativeY += skip;
        }
      }

      // Advance both cursors for the next block.
      if ("totalHeight" in measure && !(block.kind === "table" && (block as TableBlock).floating)) {
        cumulativeY += measure.totalHeight;
        pageRelativeY += measure.totalHeight;
      }

      return measure;
    } catch (error) {
      // A single bad block must not crash pagination, so fall back to a minimal
      // real measure (downstream layout treats it as a single-line paragraph of
      // fixed height). Route the failure through the layout instrumentation hook
      // so it stays traceable instead of vanishing silently.
      recordMeasureBlockError(blockIndex, block, error);
      const fallback: ParagraphMeasure = {
        kind: "paragraph",
        lines: [],
        totalHeight: 20,
      };
      cumulativeY += fallback.totalHeight;
      pageRelativeY += fallback.totalHeight;
      return fallback;
    }
  });
}

export function measureSingleBlockWithoutFloatingZones(
  block: FlowBlock,
  blockWidth: number,
  blockIndex: number,
): Measure {
  recordMeasureBlock(blockIndex, block);
  return measureBlock(block, blockWidth);
}
