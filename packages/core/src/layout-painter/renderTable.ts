/**
 * Table Renderer
 *
 * Renders table fragments to DOM. Handles:
 * - Multi-row tables split across pages
 * - Cell content (paragraphs within cells)
 * - Column widths and cell spans
 * - Basic cell styling (borders, backgrounds)
 */

import { measureParagraph } from "../layout-engine/measure";
import {
  buildTableCellFloatingZones,
  getTableCellContentWidth,
  getTableCellFloatingImages,
} from "../layout-engine/measure/tableCellFloating";
import type {
  TableFragment,
  TableBlock,
  TableMeasure,
  TableCell,
  CellBorderSpec,
  TableCellMeasure,
  ParagraphBlock,
  ParagraphMeasure,
  ParagraphFragment,
  TextBoxBlock,
  TextBoxMeasure,
  TextBoxFragment,
} from "../layout-engine/types";
import {
  getTableRowLeadingWidth,
  isFloatingImageRun,
  isFloatingTextBoxBlock,
  tableColumnsArePinned,
} from "../layout-engine/types";
import { emuToPixels } from "../utils/units";
import { getAutomaticTextColorForBackground } from "./documentColors";
import { renderParagraphFragment } from "./renderParagraph";
import { renderTextBoxFragment } from "./renderTextBox";
import type { RenderContext } from "./renderUtils";

/**
 * CSS class names for table elements
 */
export const TABLE_CLASS_NAMES = {
  table: "layout-table",
  row: "layout-table-row",
  cell: "layout-table-cell",
  cellContent: "layout-table-cell-content",
  resizeHandle: "layout-table-resize-handle",
  rowResizeHandle: "layout-table-row-resize-handle",
  tableEdgeHandleBottom: "layout-table-edge-handle-bottom",
  tableEdgeHandleRight: "layout-table-edge-handle-right",
};

const CELL_DIAGONAL_BORDER_CLASS = "layout-table-cell-diagonal-border";

/**
 * Options for rendering a table fragment
 */
export type RenderTableFragmentOptions = {
  document?: Document;
};

/**
 * Render cell content (paragraphs and nested tables)
 */
type RenderedCellContent = {
  content: HTMLElement;
  floatingLayers: HTMLElement[];
};

type RenderCellContentOptions = {
  cell: TableCell;
  cellMeasure: TableCellMeasure;
  context: RenderContext;
  doc: Document;
  contentWidthOverride?: number;
};

function renderCellContent({
  cell,
  cellMeasure,
  context,
  doc,
  contentWidthOverride,
}: RenderCellContentOptions): RenderedCellContent {
  const contentEl = doc.createElement("div");
  contentEl.className = TABLE_CLASS_NAMES.cellContent;
  contentEl.style.position = "relative";
  // Content width must account for cell padding since the cell uses border-box sizing.
  // Without this, content is wider than the available area, causing centering and
  // clipping issues (especially for nested tables).
  const contentWidth = contentWidthOverride ?? getTableCellContentWidth(cell, cellMeasure);
  contentEl.style.width = `${contentWidth}px`;

  // Extract floating images from cell paragraphs
  const cellFloatingImages = getTableCellFloatingImages(cell, cellMeasure, contentWidth);

  // Build floating zones for measurement and render floating layer
  const floatingZones = buildTableCellFloatingZones(cellFloatingImages, contentWidth);
  const floatingLayers: HTMLElement[] = [];
  if (cellFloatingImages.length > 0) {
    // Render floating image layer within the cell
    const floatingLayer = doc.createElement("div");
    floatingLayer.className = "layout-cell-floating-images-layer";
    floatingLayer.style.position = "absolute";
    floatingLayer.style.top = "0";
    floatingLayer.style.left = "0";
    floatingLayer.style.width = "100%";
    floatingLayer.style.height = "100%";
    floatingLayer.style.pointerEvents = "none";
    floatingLayer.style.zIndex = "10";
    floatingLayer.style.overflow = "visible";

    for (const img of cellFloatingImages) {
      const imgContainer = doc.createElement("div");
      imgContainer.className = "layout-cell-floating-image";
      imgContainer.style.position = "absolute";
      imgContainer.style.left = `${img.x}px`;
      imgContainer.style.top = `${img.y}px`;
      imgContainer.style.pointerEvents = "auto";
      if (img.pmStart !== undefined) {
        imgContainer.dataset["pmStart"] = String(img.pmStart);
      }
      if (img.pmEnd !== undefined) {
        imgContainer.dataset["pmEnd"] = String(img.pmEnd);
      }

      const imgEl = doc.createElement("img");
      imgEl.src = img.src;
      imgEl.style.width = `${img.width}px`;
      imgEl.style.height = `${img.height}px`;
      imgEl.style.display = "block";
      if (img.alt) {
        imgEl.alt = img.alt;
      }
      if (img.transform) {
        imgEl.style.transform = img.transform;
      }
      imgContainer.append(imgEl);
      floatingLayer.append(imgContainer);
    }

    floatingLayers.push(floatingLayer);
  }

  let cumulativeY = 0;
  let anchorParagraphY = 0;
  let floatingTextBoxesLayer: HTMLElement | undefined;
  for (let i = 0; i < cell.blocks.length; i++) {
    const block = cell.blocks[i];
    const measure = cellMeasure.blocks[i];

    if (block?.kind === "paragraph" && measure?.kind === "paragraph") {
      const paragraphBlock = block as ParagraphBlock;
      let paragraphMeasure = measure as ParagraphMeasure;

      // Re-measure when wrapping width changes or floating exclusions apply.
      if (floatingZones.length > 0 || contentWidthOverride !== undefined) {
        paragraphMeasure = measureParagraph(paragraphBlock, contentWidth, {
          floatingZones,
          paragraphYOffset: cumulativeY,
        });
      }

      // Create synthetic fragment for the paragraph
      const syntheticFragment: ParagraphFragment = {
        kind: "paragraph",
        blockId: paragraphBlock.id,
        x: 0,
        y: 0,
        width: contentWidth,
        height: paragraphMeasure.totalHeight,
        fromLine: 0,
        toLine: paragraphMeasure.lines.length,
        ...(paragraphBlock.pmStart !== undefined ? { pmStart: paragraphBlock.pmStart } : {}),
        ...(paragraphBlock.pmEnd !== undefined ? { pmEnd: paragraphBlock.pmEnd } : {}),
      };

      const cellContext = { ...context, insideTableCell: true as const };
      const fragEl = renderParagraphFragment(
        syntheticFragment,
        paragraphBlock,
        paragraphMeasure,
        cellContext,
        { document: doc },
      );

      fragEl.style.position = "relative";
      fragEl.style.boxSizing = "border-box";
      fragEl.style.height = `${paragraphMeasure.totalHeight}px`;
      const spaceBefore = paragraphBlock.attrs?.spacing?.before ?? 0;
      if (spaceBefore > 0) {
        fragEl.style.paddingTop = `${spaceBefore}px`;
      }
      contentEl.append(fragEl);
      anchorParagraphY = cumulativeY;
      cumulativeY += paragraphMeasure.totalHeight;
    } else if (block?.kind === "table" && measure?.kind === "table") {
      // Nested table - render in normal document flow.
      // Avoid cumulative marginTop offsets here: cell content already flows vertically,
      // and compounding offsets can produce enormous heights on deeply nested tables.
      const tableBlock = block as TableBlock;
      const tableMeasure = measure as TableMeasure;

      const nestedTableEl = renderNestedTable(tableBlock, tableMeasure, context, doc);
      nestedTableEl.style.position = "relative";
      contentEl.append(nestedTableEl);
      cumulativeY += (measure as TableMeasure).totalHeight;
      // A standalone anchored shape after a nested table belongs to a
      // shape-only host paragraph at the current flow position. That host is
      // omitted from the ProseMirror cell, so the nested table's bottom is the
      // anchor Y for the following floating text box.
      anchorParagraphY = cumulativeY;
    } else if (block?.kind === "textBox" && measure?.kind === "textBox") {
      const textBoxBlock = block as TextBoxBlock;
      const textBoxMeasure = measure as TextBoxMeasure;
      const syntheticFragment: TextBoxFragment = {
        kind: "textBox",
        blockId: textBoxBlock.id,
        x: 0,
        y: 0,
        width: textBoxMeasure.width,
        height: textBoxMeasure.height,
        ...(textBoxBlock.pmStart !== undefined ? { pmStart: textBoxBlock.pmStart } : {}),
        ...(textBoxBlock.pmEnd !== undefined ? { pmEnd: textBoxBlock.pmEnd } : {}),
      };
      const textBoxEl = renderTextBoxFragment(
        syntheticFragment,
        textBoxBlock,
        textBoxMeasure,
        { ...context, insideTableCell: true },
        { document: doc },
      );

      if (isFloatingTextBoxBlock(textBoxBlock)) {
        floatingTextBoxesLayer ??= createCellFloatingTextBoxesLayer(doc);
        textBoxEl.style.left = `${resolveCellTextBoxX(textBoxBlock, contentWidth)}px`;
        textBoxEl.style.top = `${anchorParagraphY + resolveCellTextBoxY(textBoxBlock)}px`;
        textBoxEl.style.pointerEvents = "auto";
        floatingTextBoxesLayer.append(textBoxEl);
        continue;
      }

      textBoxEl.style.position = "relative";
      textBoxEl.style.left = "0";
      textBoxEl.style.top = "0";
      contentEl.append(textBoxEl);
      cumulativeY += textBoxMeasure.height;
      anchorParagraphY = cumulativeY;
    }
  }

  if (floatingTextBoxesLayer) {
    floatingLayers.push(floatingTextBoxesLayer);
  }

  return {
    content: contentEl,
    floatingLayers,
  };
}

function createCellFloatingTextBoxesLayer(doc: Document): HTMLElement {
  const layer = doc.createElement("div");
  layer.className = "layout-cell-floating-text-boxes-layer";
  layer.style.position = "absolute";
  layer.style.inset = "0";
  layer.style.pointerEvents = "none";
  layer.style.zIndex = "10";
  layer.style.overflow = "visible";
  return layer;
}

function resolveCellTextBoxX(block: TextBoxBlock, contentWidth: number): number {
  const horizontal = block.position?.horizontal;
  if (horizontal?.posOffset !== undefined) {
    return emuToPixels(horizontal.posOffset);
  }
  if (horizontal?.align === "center") {
    return (contentWidth - block.width) / 2;
  }
  if (horizontal?.align === "right" || horizontal?.align === "outside") {
    return contentWidth - block.width;
  }
  return 0;
}

function resolveCellTextBoxY(block: TextBoxBlock): number {
  const vertical = block.position?.vertical;
  if (vertical?.posOffset !== undefined) {
    return emuToPixels(vertical.posOffset);
  }
  return 0;
}

function tableHasFloatingCellContent(block: TableBlock): boolean {
  for (const row of block.rows) {
    for (const cell of row.cells) {
      for (const cellBlock of cell.blocks) {
        if (cellBlock.kind === "textBox" && isFloatingTextBoxBlock(cellBlock)) {
          return true;
        }
        if (
          cellBlock.kind === "paragraph" &&
          cellBlock.runs.some((run) => run.kind === "image" && isFloatingImageRun(run))
        ) {
          return true;
        }
        if (cellBlock.kind === "table" && tableHasFloatingCellContent(cellBlock)) {
          return true;
        }
      }
    }
  }
  return false;
}

/**
 * Render a nested table (within a cell)
 */
function renderNestedTable(
  block: TableBlock,
  measure: TableMeasure,
  context: RenderContext,
  doc: Document,
): HTMLElement {
  const tableEl = doc.createElement("div");
  tableEl.className = `${TABLE_CLASS_NAMES.table} layout-nested-table`;

  // Positioning (relative, not absolute)
  tableEl.style.position = "relative";
  tableEl.style.width = `${measure.totalWidth}px`;
  tableEl.style.display = "block";

  if (block.justification === "center") {
    tableEl.style.marginLeft = "auto";
    tableEl.style.marginRight = "auto";
  } else if (block.justification === "right") {
    tableEl.style.marginLeft = "auto";
  } else if (block.indent) {
    tableEl.style.marginLeft = `${block.indent}px`;
  }

  // Store metadata
  tableEl.dataset["blockId"] = String(block.id);

  if (block.pmStart !== undefined) {
    tableEl.dataset["pmStart"] = String(block.pmStart);
  }
  if (block.pmEnd !== undefined) {
    tableEl.dataset["pmEnd"] = String(block.pmEnd);
  }

  // Build row Y positions for rowSpan height calculation
  const rowYPositions: number[] = [];
  let yPos = 0;
  for (const i_item of measure.rows) {
    rowYPositions.push(yPos);
    yPos += i_item.height;
  }
  rowYPositions.push(yPos);

  // Track spanning cells across rows
  const spanningCells = new Map<string, SpanningCell>();

  // Render all rows
  const columnsPinned = tableColumnsArePinned(block);
  const cellGrid = buildTableCellGrid(block, measure.columnWidths.length);
  let y = 0;
  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex++) {
    const row = block.rows[rowIndex];
    const rowMeasure = measure.rows[rowIndex];

    if (!row || !rowMeasure) {
      continue;
    }
    if (row.hidden) {
      y += rowMeasure.height;
      continue;
    }

    const rowEl = renderTableRow(
      row,
      rowMeasure,
      rowIndex,
      y,
      measure.columnWidths,
      block.rows.length,
      context,
      doc,
      spanningCells,
      rowYPositions,
      undefined,
      block.bidi,
      columnsPinned,
      cellGrid,
    );
    tableEl.append(rowEl);
    y += rowMeasure.height;
  }

  tableEl.style.height = `${y}px`;

  return tableEl;
}

/**
 * Apply a single border to an element.
 */
function applyBorder(
  el: HTMLElement,
  side: "top" | "right" | "bottom" | "left",
  border: { width?: number; color?: string; style?: string } | undefined,
): void {
  const styleProp = `border${side.charAt(0).toUpperCase() + side.slice(1)}` as
    | "borderTop"
    | "borderRight"
    | "borderBottom"
    | "borderLeft";

  if (!border || border.style === "none" || border.style === "nil") {
    el.style[styleProp] = "none";
  } else {
    const width = Math.max(1, border.width ?? 1);
    const color = border.color ?? "#000000";
    const style = border.style ?? "solid";
    el.style[styleProp] = `${width}px ${style} ${color}`;
  }
}

type CellDiagonalDirection = "top-left-to-bottom-right" | "top-right-to-bottom-left";

function renderCellDiagonalBorder(
  border: CellBorderSpec | undefined,
  direction: CellDiagonalDirection,
  cellWidth: number,
  cellHeight: number,
  doc: Document,
): HTMLElement | null {
  if (!hasVisibleBorder(border)) {
    return null;
  }

  const line = doc.createElement("div");
  const strokeWidth = Math.max(1, border?.width ?? 1);
  const color = border?.color ?? "#000000";
  const length = Math.hypot(cellWidth, cellHeight);
  const angle = Math.atan2(cellHeight, cellWidth);

  line.className = CELL_DIAGONAL_BORDER_CLASS;
  line.dataset["direction"] = direction;
  line.style.position = "absolute";
  line.style.left = "0";
  line.style.top =
    direction === "top-left-to-bottom-right"
      ? `${-strokeWidth / 2}px`
      : `${cellHeight - strokeWidth / 2}px`;
  line.style.width = `${length}px`;
  line.style.height = `${strokeWidth}px`;
  line.style.transformOrigin = "0 50%";
  line.style.transform = `rotate(${direction === "top-left-to-bottom-right" ? angle : -angle}rad)`;
  line.style.pointerEvents = "none";
  line.style.zIndex = "1";

  switch (border?.style) {
    case "dashed":
      line.style.backgroundImage = `repeating-linear-gradient(to right, ${color} 0, ${color} ${strokeWidth * 3}px, transparent ${strokeWidth * 3}px, transparent ${strokeWidth * 5}px)`;
      break;
    case "dotted":
      line.style.backgroundImage = `radial-gradient(circle at ${strokeWidth / 2}px 50%, ${color} ${strokeWidth / 2}px, transparent ${strokeWidth / 2}px)`;
      line.style.backgroundSize = `${strokeWidth * 2}px ${strokeWidth}px`;
      break;
    case "double":
      line.style.backgroundImage = `linear-gradient(to bottom, ${color} 0, ${color} 33%, transparent 33%, transparent 67%, ${color} 67%, ${color} 100%)`;
      break;
    default:
      line.style.backgroundColor = color;
      break;
  }

  return line;
}

/**
 * Render a single table cell
 */
function renderTableCell(
  cell: TableCell,
  cellMeasure: TableCellMeasure,
  x: number,
  rowHeight: number,
  borderFlags: {
    drawTop: boolean;
    isLastRow: boolean;
    drawLeft: boolean;
    isLastCol: boolean;
  },
  columnsPinned: boolean,
  context: RenderContext,
  doc: Document,
): HTMLElement {
  const cellEl = doc.createElement("div");
  cellEl.className = TABLE_CLASS_NAMES.cell;

  // Positioning
  cellEl.style.position = "absolute";
  cellEl.style.left = `${x}px`;
  cellEl.style.top = "0";
  cellEl.style.width = `${cellMeasure.width}px`;
  cellEl.style.height = `${rowHeight}px`;
  cellEl.style.overflow = "hidden";
  cellEl.style.boxSizing = "border-box";
  // Use per-cell padding from DOCX margins, default to Word's visual rendering
  const padTop = cell.padding?.top ?? 1;
  const padRight = cell.padding?.right ?? 7;
  const padBottom = cell.padding?.bottom ?? 1;
  const padLeft = cell.padding?.left ?? 7;
  cellEl.style.padding = `${padTop}px ${padRight}px ${padBottom}px ${padLeft}px`;

  // Apply borders - use cell borders if available, otherwise no border
  if (cell.borders) {
    // Collapse shared borders to avoid double-thick lines.
    // Strategy: "bottom wins" for rows, "right wins" for columns.
    // Each cell's bottom border represents the shared edge with the row below.
    // Each cell's right border represents the shared edge with the column to its right.
    // Top/left edges are retained when the adjacent cell does not already
    // own the shared edge. This also preserves partial boxes that begin inside
    // a table instead of limiting those sides to the table perimeter.
    if (borderFlags.drawTop) {
      applyBorder(cellEl, "top", cell.borders.top);
    }
    applyBorder(cellEl, "right", cell.borders.right);
    applyBorder(cellEl, "bottom", cell.borders.bottom);
    if (borderFlags.drawLeft) {
      applyBorder(cellEl, "left", cell.borders.left);
    }
  }
  // No default border - cells without explicit borders should be borderless

  // Background color
  if (cell.background) {
    cellEl.style.backgroundColor = cell.background;
    const automaticTextColor = getAutomaticTextColorForBackground(cell.background);
    if (automaticTextColor) {
      cellEl.style.color = automaticTextColor;
    }
  }

  // `w:noWrap` (§17.4.30): forbid soft-wrapping inside the cell. This applies
  // only when the columns can grow to honor it (auto-width table); measurement
  // then collapses the cell to a single MeasuredLine (see NO_WRAP_MEASURE_WIDTH)
  // and this style is the inline-content guard that stops the browser from
  // re-wrapping the single line div (e.g. user zoom, font fallback widening).
  // When the columns are pinned Word cannot honor `w:noWrap`, so measurement
  // wrapped the cell; emitting `nowrap` here would collapse those lines and
  // desync the painted height from the measured height. eigenpal #424.
  if (cell.noWrap && !columnsPinned) {
    cellEl.style.whiteSpace = "nowrap";
  }

  // Vertical alignment
  if (cell.verticalAlign) {
    cellEl.style.display = "flex";
    cellEl.style.flexDirection = "column";
    switch (cell.verticalAlign) {
      case "top":
        cellEl.style.justifyContent = "flex-start";
        break;
      case "center":
        cellEl.style.justifyContent = "center";
        break;
      case "bottom":
        cellEl.style.justifyContent = "flex-end";
        break;
      default:
        break;
    }
  }

  // Render cell content
  const contentWidthOverride =
    cell.textDirection === "btLr" ? Math.max(1, rowHeight - padTop - padBottom) : undefined;
  const renderedContent = renderCellContent({
    cell,
    cellMeasure,
    context,
    doc,
    ...(contentWidthOverride !== undefined ? { contentWidthOverride } : {}),
  });
  if (cell.textDirection === "btLr") {
    renderedContent.content.style.position = "absolute";
    renderedContent.content.style.left = "50%";
    renderedContent.content.style.top = "50%";
    renderedContent.content.style.transform = "translate(-50%, -50%) rotate(-90deg)";
  }
  if (renderedContent.floatingLayers.length > 0) {
    renderedContent.content.style.height = "100%";
    renderedContent.content.style.overflow = "hidden";
    cellEl.style.overflow = "visible";
  }
  cellEl.append(renderedContent.content);
  const topLeftToBottomRight = renderCellDiagonalBorder(
    cell.borders?.topLeftToBottomRight,
    "top-left-to-bottom-right",
    cellMeasure.width,
    rowHeight,
    doc,
  );
  if (topLeftToBottomRight) {
    cellEl.append(topLeftToBottomRight);
  }
  const topRightToBottomLeft = renderCellDiagonalBorder(
    cell.borders?.topRightToBottomLeft,
    "top-right-to-bottom-left",
    cellMeasure.width,
    rowHeight,
    doc,
  );
  if (topRightToBottomLeft) {
    cellEl.append(topRightToBottomLeft);
  }
  for (const floatingLayer of renderedContent.floatingLayers) {
    floatingLayer.style.left = `${padLeft}px`;
    floatingLayer.style.top = `${padTop}px`;
    cellEl.append(floatingLayer);
  }

  // Store PM positions for selection
  if (cell.blocks.length > 0) {
    const firstBlock = cell.blocks.at(0);
    const lastBlock = cell.blocks.at(-1);
    const pmStart =
      firstBlock !== undefined && "pmStart" in firstBlock ? firstBlock.pmStart : undefined;
    if (pmStart !== undefined) {
      cellEl.dataset["pmStart"] = String(pmStart);
    }
    if (lastBlock && "pmEnd" in lastBlock) {
      cellEl.dataset["pmEnd"] = String(lastBlock.pmEnd);
    }
  }

  return cellEl;
}

/**
 * Track cells that span multiple rows
 */
type SpanningCell = {
  cell: TableCell;
  cellMeasure: TableCellMeasure;
  columnIndex: number;
  startRow: number;
  rowSpan: number;
  colSpan: number;
  x: number;
  totalHeight: number;
};

type TableCellGrid = Map<string, TableCell>;

const tableCellGridKey = (rowIndex: number, columnIndex: number): string =>
  `${rowIndex}:${columnIndex}`;

function buildTableCellGrid(block: TableBlock, columnCount: number): TableCellGrid {
  const grid: TableCellGrid = new Map();

  for (let rowIndex = 0; rowIndex < block.rows.length; rowIndex++) {
    const row = block.rows[rowIndex];
    if (!row) {
      continue;
    }
    let columnIndex = 0;
    for (const cell of row.cells) {
      while (grid.has(tableCellGridKey(rowIndex, columnIndex))) {
        columnIndex += 1;
      }
      const colSpan = cell.colSpan ?? 1;
      const rowSpan = cell.rowSpan ?? 1;
      const rowEnd = Math.min(block.rows.length, rowIndex + rowSpan);
      const columnEnd = Math.min(columnCount, columnIndex + colSpan);
      for (let gridRow = rowIndex; gridRow < rowEnd; gridRow++) {
        for (let gridColumn = columnIndex; gridColumn < columnEnd; gridColumn++) {
          grid.set(tableCellGridKey(gridRow, gridColumn), cell);
        }
      }
      columnIndex += colSpan;
    }
  }

  return grid;
}

const hasVisibleBorder = (border: { width?: number; style?: string } | undefined): boolean =>
  border !== undefined && border.style !== "none" && border.style !== "nil";

/**
 * Render a table row with rowSpan support
 */
function renderTableRow(
  row: TableBlock["rows"][number],
  rowMeasure: TableMeasure["rows"][number],
  rowIndex: number,
  y: number,
  columnWidths: number[],
  totalRows: number,
  context: RenderContext,
  doc: Document,
  spanningCells?: Map<string, SpanningCell>,
  rowYPositions?: number[],
  isFirstRowInFragment?: boolean,
  bidi = false,
  columnsPinned = false,
  cellGrid?: TableCellGrid,
): HTMLElement {
  const rowEl = doc.createElement("div");
  rowEl.className = TABLE_CLASS_NAMES.row;

  // RTL tables (w:bidiVisual) mirror columns: logical column 0 paints on the
  // right. Geometry-only — cell content and saved DOCX are unchanged
  // (eigenpal/docx-editor#940).
  const tableWidth = bidi ? columnWidths.reduce((sum, columnWidth) => sum + columnWidth, 0) : 0;

  // Positioning
  rowEl.style.position = "absolute";
  rowEl.style.left = "0";
  rowEl.style.top = `${y}px`;
  rowEl.style.width = "100%";
  rowEl.style.height = `${rowMeasure.height}px`;

  // Data attributes
  rowEl.dataset["rowIndex"] = String(rowIndex);

  // Build set of columns occupied by spanning cells from previous rows
  const occupiedColumns = new Set<number>();
  if (spanningCells) {
    for (const [, spanCell] of spanningCells) {
      // Check if this spanning cell covers the current row
      if (spanCell.startRow < rowIndex && spanCell.startRow + spanCell.rowSpan > rowIndex) {
        for (let c = 0; c < spanCell.colSpan; c++) {
          occupiedColumns.add(spanCell.columnIndex + c);
        }
      }
    }
  }

  // Render cells
  // Track actual column index separately from cell index
  // because cells with colSpan > 1 span multiple columns
  const gridBefore = row.gridBefore ?? 0;
  let x = getTableRowLeadingWidth(row, columnWidths);
  let columnIndex = gridBefore;

  // Skip columns occupied by spanning cells
  while (occupiedColumns.has(columnIndex)) {
    x += columnWidths[columnIndex] ?? 0;
    columnIndex++;
  }

  for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex++) {
    const cell = row.cells[cellIndex];
    const cellMeasure = rowMeasure.cells[cellIndex];

    if (!cell || !cellMeasure) {
      continue;
    }

    const colSpan = cell.colSpan ?? 1;
    const rowSpan = cell.rowSpan ?? 1;

    // Calculate cell height - for spanning cells, use total height of spanned rows
    let cellHeight = rowMeasure.height;
    if (rowSpan > 1 && rowYPositions) {
      cellHeight = 0;
      for (let r = rowIndex; r < rowIndex + rowSpan && r < rowYPositions.length - 1; r++) {
        cellHeight += (rowYPositions[r + 1] ?? 0) - (rowYPositions[r] ?? 0);
      }
      // Fallback if rowYPositions doesn't have enough entries
      if (cellHeight === 0) {
        cellHeight = rowMeasure.height * rowSpan;
      }
    }

    let cellWidth = 0;
    for (let c = 0; c < colSpan && columnIndex + c < columnWidths.length; c++) {
      cellWidth += columnWidths[columnIndex + c] ?? 0;
    }
    const cellLeft = bidi ? tableWidth - x - cellWidth : x;

    const isFirstRow = rowIndex === 0 || isFirstRowInFragment === true;
    const isLastRow = rowIndex + rowSpan >= totalRows;
    // Outer-border sides are visual: in RTL the logical last column is the
    // visual first (leftmost), so swap which cell draws the left/right border.
    const atLogicalStart = columnIndex === 0;
    const atLogicalEnd = columnIndex + colSpan >= columnWidths.length;
    const isFirstCol = bidi ? atLogicalEnd : atLogicalStart;
    const isLastCol = bidi ? atLogicalStart : atLogicalEnd;
    const aboveCell = cellGrid?.get(tableCellGridKey(rowIndex - 1, columnIndex));
    const leftNeighborColumn = bidi ? columnIndex + colSpan : columnIndex - 1;
    const leftCell = cellGrid?.get(tableCellGridKey(rowIndex, leftNeighborColumn));
    const drawTop = isFirstRow || !hasVisibleBorder(aboveCell?.borders?.bottom);
    const drawLeft = isFirstCol || !hasVisibleBorder(leftCell?.borders?.right);

    const cellEl = renderTableCell(
      cell,
      cellMeasure,
      cellLeft,
      cellHeight,
      { drawTop, isLastRow, drawLeft, isLastCol },
      columnsPinned,
      context,
      doc,
    );
    cellEl.dataset["cellIndex"] = String(cellIndex);
    cellEl.dataset["columnIndex"] = String(columnIndex);

    // Store rowSpan info for styling
    if (rowSpan > 1) {
      cellEl.dataset["rowSpan"] = String(rowSpan);
    }

    rowEl.append(cellEl);

    // Track this cell as spanning if it spans multiple rows
    if (rowSpan > 1 && spanningCells) {
      const key = `${rowIndex}-${columnIndex}`;
      spanningCells.set(key, {
        cell,
        cellMeasure,
        columnIndex,
        startRow: rowIndex,
        rowSpan,
        colSpan,
        x: cellLeft,
        totalHeight: cellHeight,
      });
    }

    // Advance the logical running offset by the columns this cell spans
    x += cellWidth;

    // Advance column index by colSpan
    columnIndex += colSpan;

    // Skip columns occupied by spanning cells
    while (occupiedColumns.has(columnIndex)) {
      x += columnWidths[columnIndex] ?? 0;
      columnIndex++;
    }
  }

  return rowEl;
}

/**
 * Render a table fragment to DOM
 *
 * @param fragment - The table fragment to render
 * @param block - The full table block
 * @param measure - The full table measure
 * @param context - Rendering context
 * @param options - Rendering options
 * @returns The table DOM element
 */
export function renderTableFragment(
  fragment: TableFragment,
  block: TableBlock,
  measure: TableMeasure,
  context: RenderContext,
  options: RenderTableFragmentOptions = {},
): HTMLElement {
  const doc = options.document ?? document;

  const tableEl = doc.createElement("div");
  tableEl.className = TABLE_CLASS_NAMES.table;

  // Basic table styling
  tableEl.style.position = "absolute";
  tableEl.style.width = `${fragment.width}px`;
  tableEl.style.height = `${fragment.height}px`;
  tableEl.style.overflow = tableHasFloatingCellContent(block) ? "visible" : "hidden";

  // Store metadata
  tableEl.dataset["blockId"] = String(fragment.blockId);
  tableEl.dataset["fromRow"] = String(fragment.fromRow);
  tableEl.dataset["toRow"] = String(fragment.toRow);

  if (fragment.pmStart !== undefined) {
    tableEl.dataset["pmStart"] = String(fragment.pmStart);
  }
  if (fragment.pmEnd !== undefined) {
    tableEl.dataset["pmEnd"] = String(fragment.pmEnd);
  }

  // Add column resize handles at each column boundary. For RTL tables the
  // boundary at the running left offset maps to the mirrored visual position
  // so handles stay on the visible column edges (eigenpal/docx-editor#940).
  const tableColumnsWidth = block.bidi
    ? measure.columnWidths.reduce((sum, columnWidth) => sum + columnWidth, 0)
    : 0;
  let handleX = 0;
  for (let col = 0; col < measure.columnWidths.length - 1; col++) {
    handleX += measure.columnWidths[col] ?? 0;
    const handleLeft = block.bidi ? tableColumnsWidth - handleX : handleX;
    const handle = doc.createElement("div");
    handle.className = TABLE_CLASS_NAMES.resizeHandle;
    handle.style.position = "absolute";
    handle.style.left = `${handleLeft - 3}px`;
    handle.style.top = "0";
    handle.style.width = "6px";
    handle.style.height = "100%";
    handle.style.cursor = "col-resize";
    handle.style.zIndex = "10";
    handle.dataset["columnIndex"] = String(col);
    handle.dataset["tableBlockId"] = String(fragment.blockId);
    // RTL boundary: the two columns are visually mirrored, so the resize path
    // must invert the drag delta (eigenpal/docx-editor#940).
    if (block.bidi) {
      handle.dataset["bidi"] = "true";
    }
    if (fragment.pmStart !== undefined) {
      handle.dataset["tablePmStart"] = String(fragment.pmStart);
    }
    tableEl.append(handle);
  }

  // Build row Y positions for rowSpan height calculation
  const rowYPositions: number[] = [];
  let yPos = 0;
  for (const i_item of measure.rows) {
    rowYPositions.push(yPos);
    yPos += i_item.height;
  }
  rowYPositions.push(yPos); // Add final position for height calculation

  // Track spanning cells across rows
  const spanningCells = new Map<string, SpanningCell>();
  const columnsPinned = tableColumnsArePinned(block);
  const cellGrid = buildTableCellGrid(block, measure.columnWidths.length);

  // Render repeated header rows for continuation fragments. For a mid-content
  // row break (eigenpal #698), keep repeated headers pinned to the fragment top
  // and offset only the content row stack by topClip.
  const headerRowCount = fragment.headerRowCount ?? 0;
  let y = 0;
  if (headerRowCount > 0 && fragment.continuesFromPrev) {
    for (let hdrIdx = 0; hdrIdx < headerRowCount; hdrIdx++) {
      const hdrRow = block.rows[hdrIdx];
      const hdrRowMeasure = measure.rows[hdrIdx];
      if (!hdrRow || !hdrRowMeasure) {
        continue;
      }
      if (hdrRow.hidden) {
        y += hdrRowMeasure.height;
        continue;
      }

      const rowEl = renderTableRow(
        hdrRow,
        hdrRowMeasure,
        hdrIdx,
        y,
        measure.columnWidths,
        block.rows.length,
        context,
        doc,
        spanningCells,
        rowYPositions,
        hdrIdx === 0, // first header row draws top border
        block.bidi,
        columnsPinned,
        cellGrid,
      );
      rowEl.dataset["repeatedHeader"] = "true";
      tableEl.append(rowEl);
      y += hdrRowMeasure.height;
    }
  }

  const topClip = fragment.topClip ?? 0;
  let contentParent: HTMLElement = tableEl;
  if (headerRowCount > 0 && fragment.continuesFromPrev && topClip > 0) {
    const contentClipEl = doc.createElement("div");
    contentClipEl.style.position = "absolute";
    contentClipEl.style.left = "0";
    contentClipEl.style.top = `${y}px`;
    contentClipEl.style.width = "100%";
    contentClipEl.style.height = `${Math.max(0, fragment.height - y)}px`;
    contentClipEl.style.overflow = "hidden";
    tableEl.append(contentClipEl);
    contentParent = contentClipEl;
    y = -topClip;
  } else {
    y -= topClip;
  }

  // Render content rows from fragment.fromRow to fragment.toRow
  for (let rowIndex = fragment.fromRow; rowIndex < fragment.toRow; rowIndex++) {
    const row = block.rows[rowIndex];
    const rowMeasure = measure.rows[rowIndex];

    if (!row || !rowMeasure) {
      continue;
    }
    if (row.hidden) {
      y += rowMeasure.height;
      continue;
    }

    // First content row in a continuation fragment with headers should draw top border
    const isFirstRowInFragment =
      headerRowCount > 0 && fragment.continuesFromPrev
        ? false // header rows already drawn, content rows are not "first"
        : fragment.continuesFromPrev && rowIndex === fragment.fromRow;

    const rowEl = renderTableRow(
      row,
      rowMeasure,
      rowIndex,
      y,
      measure.columnWidths,
      block.rows.length,
      context,
      doc,
      spanningCells,
      rowYPositions,
      isFirstRowInFragment,
      block.bidi,
      columnsPinned,
      cellGrid,
    );

    contentParent.append(rowEl);
    y += rowMeasure.height;
  }

  // Add row resize handles at each row boundary (between consecutive rows)
  let handleY = 0;
  for (let rowIdx = fragment.fromRow; rowIdx < fragment.toRow; rowIdx++) {
    if (block.rows[rowIdx]?.hidden) {
      continue;
    }
    handleY += measure.rows[rowIdx]?.height ?? 0;

    // Don't add a handle after the last row in this fragment (unless it's the table's last row — that's the bottom edge)
    if (rowIdx < fragment.toRow - 1) {
      const rowHandle = doc.createElement("div");
      rowHandle.className = TABLE_CLASS_NAMES.rowResizeHandle;
      rowHandle.style.position = "absolute";
      rowHandle.style.left = "0";
      rowHandle.style.top = `${handleY - 3}px`;
      rowHandle.style.width = "100%";
      rowHandle.style.height = "6px";
      rowHandle.style.cursor = "row-resize";
      rowHandle.style.zIndex = "10";
      rowHandle.dataset["rowIndex"] = String(rowIdx);
      rowHandle.dataset["tableBlockId"] = String(fragment.blockId);
      if (fragment.pmStart !== undefined) {
        rowHandle.dataset["tablePmStart"] = String(fragment.pmStart);
      }
      tableEl.append(rowHandle);
    }
  }

  // Bottom edge handle (only on fragments containing the last row)
  if (fragment.toRow === block.rows.length) {
    const bottomHandle = doc.createElement("div");
    bottomHandle.className = TABLE_CLASS_NAMES.tableEdgeHandleBottom;
    bottomHandle.style.position = "absolute";
    bottomHandle.style.left = "0";
    bottomHandle.style.top = `${handleY - 3}px`;
    bottomHandle.style.width = "100%";
    bottomHandle.style.height = "6px";
    bottomHandle.style.cursor = "row-resize";
    bottomHandle.style.zIndex = "10";
    bottomHandle.dataset["rowIndex"] = String(block.rows.length - 1);
    bottomHandle.dataset["tableBlockId"] = String(fragment.blockId);
    bottomHandle.dataset["isEdge"] = "bottom";
    if (fragment.pmStart !== undefined) {
      bottomHandle.dataset["tablePmStart"] = String(fragment.pmStart);
    }
    tableEl.append(bottomHandle);
  }

  // Right edge handle (only on fragments containing the last row)
  if (fragment.toRow === block.rows.length) {
    const totalWidth = measure.columnWidths.reduce((w, cw) => w + cw, 0);
    const rightHandle = doc.createElement("div");
    rightHandle.className = TABLE_CLASS_NAMES.tableEdgeHandleRight;
    rightHandle.style.position = "absolute";
    rightHandle.style.left = `${totalWidth - 3}px`;
    rightHandle.style.top = "0";
    rightHandle.style.width = "6px";
    rightHandle.style.height = "100%";
    rightHandle.style.cursor = "col-resize";
    rightHandle.style.zIndex = "10";
    // The visual right edge belongs to logical column 0 in an RTL table, so the
    // edge handle resizes that column instead of the last one
    // (eigenpal/docx-editor#940).
    rightHandle.dataset["columnIndex"] = String(block.bidi ? 0 : measure.columnWidths.length - 1);
    rightHandle.dataset["tableBlockId"] = String(fragment.blockId);
    rightHandle.dataset["isEdge"] = "right";
    if (fragment.pmStart !== undefined) {
      rightHandle.dataset["tablePmStart"] = String(fragment.pmStart);
    }
    tableEl.append(rightHandle);
  }

  return tableEl;
}
