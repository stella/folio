/**
 * Table fragments → primitives.
 *
 * Order per cell is fixed here rather than left to a backend: background,
 * borders, content, diagonals. `renderTable.ts` gets that order from CSS
 * (a background paints under its own border, a `z-index: 1` diagonal over the
 * content); nothing in a display list would reproduce it by accident.
 *
 * One deliberate divergence from the painter: `renderCellDiagonalBorder`
 * (`renderTable.ts:533`) emulates a diagonal with a rotated bar whose dash
 * pattern is a CSS gradient. Here it is a real diagonal `line` with the
 * authored stroke pattern, so a backend draws a line rather than reverse
 * engineering a gradient.
 */

import {
  placeTableCellBlock,
  createTableCellFlowState,
} from "../../layout-engine/measure/tableCellFlow";
import {
  buildTableCellPlacements,
  getSourceCellAt,
  buildTableCellGrid,
} from "../../layout-engine/measure/tableCellGrid";
import { resolveTableCellPadding } from "../../layout-engine/types";
import type {
  FlowBlock,
  Measure,
  TableBlock,
  TableCell,
  TableFragment,
  TableMeasure,
  TableRow,
} from "../../layout-engine/types";
import type { DisplayPrimitive, DisplayRect, DisplayStroke } from "../types";
import type { BuildContext } from "./buildContext";
import { parseDisplayColor } from "./colors";
import { paintParagraphFragment } from "./paragraphPrimitives";
import { resolveBorderStroke } from "./strokes";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** A cell edge is "visible" for collapse purposes on style alone, as the painter decides it. */
const hasVisibleBorder = (border: { style?: string } | undefined): boolean =>
  border !== undefined && border.style !== "none" && border.style !== "nil";

type CellBox = {
  readonly xPx: number;
  readonly yPx: number;
  readonly widthPx: number;
  readonly heightPx: number;
};

type BorderWidths = {
  readonly top: number;
  readonly right: number;
  readonly bottom: number;
  readonly left: number;
};

const strokeFor = (
  border: { width?: number; style?: string; color?: string } | undefined,
  context: BuildContext,
  label: string,
): DisplayStroke | undefined => {
  if (!hasVisibleBorder(border) || !border) {
    return undefined;
  }
  const { stroke, unresolvedColor } = resolveBorderStroke(border);
  if (unresolvedColor !== undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.unresolvedColor,
      context.pageIndex,
      `${label} colour ${unresolvedColor}`,
    );
  }
  return stroke;
};

type PaintCellOptions = {
  readonly cell: TableCell;
  readonly cellMeasure: { blocks: Measure[]; width: number; height: number };
  readonly box: CellBox;
  readonly context: BuildContext;
  readonly sides: {
    readonly top: boolean;
    readonly bottom: boolean;
    readonly left: boolean;
    readonly right: boolean;
  };
};

const paintCell = ({
  cell,
  cellMeasure,
  box,
  context,
  sides,
}: PaintCellOptions): readonly DisplayPrimitive[] => {
  const primitives: DisplayPrimitive[] = [];

  if (cell.background) {
    const fill = parseDisplayColor(cell.background);
    if (fill) {
      primitives.push({
        kind: "rect",
        rect: { xPx: box.xPx, yPx: box.yPx, widthPx: box.widthPx, heightPx: box.heightPx },
        fill,
      });
    } else {
      context.unsupported.report(
        UNSUPPORTED_CONSTRUCT.unresolvedColor,
        context.pageIndex,
        `table cell background ${cell.background}`,
      );
    }
  }

  const borders = cell.borders;
  const strokes = {
    top: sides.top ? strokeFor(borders?.top, context, "table cell top border") : undefined,
    bottom: sides.bottom
      ? strokeFor(borders?.bottom, context, "table cell bottom border")
      : undefined,
    left: sides.left ? strokeFor(borders?.left, context, "table cell left border") : undefined,
    right: strokeFor(borders?.right, context, "table cell right border"),
  };

  // `box-sizing: border-box` puts each edge inside the cell box; a display-list
  // stroke is centred on its path, so each edge sits half a thickness inside.
  const widths: BorderWidths = {
    top: strokes.top?.thicknessPx ?? 0,
    right: strokes.right?.thicknessPx ?? 0,
    bottom: strokes.bottom?.thicknessPx ?? 0,
    left: strokes.left?.thicknessPx ?? 0,
  };

  if (strokes.top) {
    primitives.push({
      kind: "line",
      x1Px: box.xPx,
      y1Px: box.yPx + widths.top / 2,
      x2Px: box.xPx + box.widthPx,
      y2Px: box.yPx + widths.top / 2,
      stroke: strokes.top,
    });
  }
  if (strokes.bottom) {
    primitives.push({
      kind: "line",
      x1Px: box.xPx,
      y1Px: box.yPx + box.heightPx - widths.bottom / 2,
      x2Px: box.xPx + box.widthPx,
      y2Px: box.yPx + box.heightPx - widths.bottom / 2,
      stroke: strokes.bottom,
    });
  }
  if (strokes.left) {
    primitives.push({
      kind: "line",
      x1Px: box.xPx + widths.left / 2,
      y1Px: box.yPx,
      x2Px: box.xPx + widths.left / 2,
      y2Px: box.yPx + box.heightPx,
      stroke: strokes.left,
    });
  }
  if (strokes.right) {
    primitives.push({
      kind: "line",
      x1Px: box.xPx + box.widthPx - widths.right / 2,
      y1Px: box.yPx,
      x2Px: box.xPx + box.widthPx - widths.right / 2,
      y2Px: box.yPx + box.heightPx,
      stroke: strokes.right,
    });
  }

  const padding = resolveTableCellPadding(cell);
  const contentXPx = box.xPx + widths.left + padding.left;
  const contentYPx = box.yPx + widths.top + padding.top;
  // Measurement used `cellWidth - padLeft - padRight`, so paint must too:
  // subtracting the border widths as well would wrap text the engine did not.
  const contentWidthPx = Math.max(0, cellMeasure.width - padding.left - padding.right);

  if (cell.textDirection === "btLr" || cell.textDirection === "tbRl") {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.verticalCellText,
      context.pageIndex,
      `w:textDirection ${cell.textDirection} paints horizontally`,
    );
  }

  primitives.push(
    ...paintCellBlocks({
      blocks: cell.blocks,
      measures: cellMeasure.blocks,
      xPx: contentXPx,
      yPx: contentYPx,
      widthPx: contentWidthPx,
      context,
    }),
  );

  const diagonals = [
    { border: borders?.topLeftToBottomRight, from: "topLeft" as const },
    { border: borders?.topRightToBottomLeft, from: "topRight" as const },
  ];
  for (const { border, from } of diagonals) {
    const stroke = strokeFor(border, context, "table cell diagonal border");
    if (!stroke) {
      continue;
    }
    primitives.push({
      kind: "line",
      x1Px: box.xPx,
      y1Px: from === "topLeft" ? box.yPx : box.yPx + box.heightPx,
      x2Px: box.xPx + box.widthPx,
      y2Px: from === "topLeft" ? box.yPx + box.heightPx : box.yPx,
      stroke,
    });
  }

  return primitives;
};

type PaintCellBlocksOptions = {
  readonly blocks: readonly FlowBlock[];
  readonly measures: readonly Measure[];
  readonly xPx: number;
  readonly yPx: number;
  readonly widthPx: number;
  readonly context: BuildContext;
};

/**
 * A cell's inner blocks, stacked by the same flow state the measurer used, so
 * paragraph spacing collapses in paint exactly where it collapsed in layout.
 */
const paintCellBlocks = ({
  blocks,
  measures,
  xPx,
  yPx,
  widthPx,
  context,
}: PaintCellBlocksOptions): readonly DisplayPrimitive[] => {
  const primitives: DisplayPrimitive[] = [];
  const flowState = createTableCellFlowState();

  for (let index = 0; index < blocks.length; index += 1) {
    const block = blocks[index];
    const measure = measures[index];
    if (!block || !measure) {
      continue;
    }
    const placement = placeTableCellBlock(flowState, block, measure);

    if (block.kind === "paragraph" && measure.kind === "paragraph") {
      const previous = blocks[index - 1];
      const next = blocks[index + 1];
      primitives.push(
        ...paintParagraphFragment({
          fragment: {
            kind: "paragraph",
            blockId: block.id,
            x: xPx,
            y: yPx + placement.contentTop,
            width: widthPx,
            height: placement.contentHeight,
            fromLine: 0,
            toLine: measure.lines.length,
          },
          block,
          measure,
          context,
          ...(previous?.kind === "paragraph" && previous.attrs?.borders !== undefined
            ? { prevBorders: previous.attrs.borders }
            : {}),
          ...(next?.kind === "paragraph" && next.attrs?.borders !== undefined
            ? { nextBorders: next.attrs.borders }
            : {}),
        }),
      );
      continue;
    }

    if (block.kind === "table" && measure.kind === "table") {
      primitives.push(
        ...paintTableBlock({ block, measure, xPx, yPx: yPx + placement.contentTop, context }),
      );
      continue;
    }

    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.fragmentKind,
      context.pageIndex,
      `table cell content of kind ${block.kind} is not painted`,
    );
  }

  return primitives;
};

type PaintTableBodyOptions = {
  readonly block: TableBlock;
  readonly measure: TableMeasure;
  readonly xPx: number;
  readonly yPx: number;
  readonly fromRow: number;
  readonly toRow: number;
  readonly headerRowCount: number;
  readonly context: BuildContext;
};

const rowIsPainted = (row: TableRow | undefined): row is TableRow =>
  row !== undefined && row.hidden !== true;

const paintTableBody = ({
  block,
  measure,
  xPx,
  yPx,
  fromRow,
  toRow,
  headerRowCount,
  context,
}: PaintTableBodyOptions): readonly DisplayPrimitive[] => {
  const primitives: DisplayPrimitive[] = [];
  const grid = buildTableCellGrid(block.rows, measure.columnWidths.length);
  const placements = buildTableCellPlacements({
    grid,
    columnWidths: measure.columnWidths,
    bidi: block.bidi === true,
  });

  // A rowSpan cell is as tall as the rows it covers, measured over the whole
  // table (not the fragment), matching `rowYPositions` in the painter.
  const rowTops: number[] = [0];
  for (const rowMeasure of measure.rows) {
    rowTops.push((rowTops.at(-1) ?? 0) + rowMeasure.height);
  }

  const paintRow = (rowIndex: number, rowYPx: number): void => {
    const row = block.rows[rowIndex];
    const rowMeasure = measure.rows[rowIndex];
    if (!rowIsPainted(row) || !rowMeasure) {
      return;
    }

    for (let cellIndex = 0; cellIndex < row.cells.length; cellIndex += 1) {
      const cell = row.cells[cellIndex];
      const cellMeasure = rowMeasure.cells[cellIndex];
      if (!cell || !cellMeasure) {
        continue;
      }
      const placement = placements.get(cell);
      if (!placement) {
        continue;
      }

      const rowSpan = Math.max(1, Math.trunc(cell.rowSpan ?? 1));
      const spanEnd = Math.min(rowTops.length - 1, rowIndex + rowSpan);
      const heightPx =
        // SAFETY: both indices are inside `rowTops`, which has rows.length + 1 entries.
        spanEnd > rowIndex ? rowTops[spanEnd]! - rowTops[rowIndex]! : rowMeasure.height;

      const atLogicalStart = placement.sourceColumn === 0;
      const atLogicalEnd =
        placement.sourceColumn + placement.columnSpan >= measure.columnWidths.length;
      const isFirstColumn = block.bidi === true ? atLogicalEnd : atLogicalStart;
      const aboveCell = getSourceCellAt(grid, rowIndex - 1, placement.sourceColumn);
      const leftNeighborColumn =
        block.bidi === true
          ? placement.sourceColumn + placement.columnSpan
          : placement.sourceColumn - 1;
      const leftCell = getSourceCellAt(grid, rowIndex, leftNeighborColumn);

      primitives.push(
        ...paintCell({
          cell,
          cellMeasure,
          box: {
            xPx: xPx + placement.left,
            yPx: rowYPx,
            widthPx: placement.width,
            heightPx,
          },
          context,
          sides: {
            // The shared edge belongs to the upper / leading cell; the other
            // side suppresses its own only when that owner actually draws one.
            top: rowIndex === fromRow || !hasVisibleBorder(aboveCell?.borders?.bottom),
            bottom: true,
            left: isFirstColumn || !hasVisibleBorder(leftCell?.borders?.right),
            right: true,
          },
        }),
      );
    }
  };

  let cursorYPx = yPx;
  for (let rowIndex = 0; rowIndex < headerRowCount; rowIndex += 1) {
    const rowMeasure = measure.rows[rowIndex];
    if (!rowMeasure) {
      continue;
    }
    paintRow(rowIndex, cursorYPx);
    cursorYPx += rowMeasure.height;
  }

  for (let rowIndex = fromRow; rowIndex < toRow; rowIndex += 1) {
    const rowMeasure = measure.rows[rowIndex];
    if (!rowMeasure) {
      continue;
    }
    paintRow(rowIndex, cursorYPx);
    cursorYPx += rowMeasure.height;
  }

  return primitives;
};

export type TableBlockPaintOptions = {
  readonly block: TableBlock;
  readonly measure: TableMeasure;
  readonly xPx: number;
  readonly yPx: number;
  readonly context: BuildContext;
};

/** A whole, unpaginated table: what a nested table inside a cell or box is. */
export const paintTableBlock = ({
  block,
  measure,
  xPx,
  yPx,
  context,
}: TableBlockPaintOptions): readonly DisplayPrimitive[] =>
  paintTableBody({
    block,
    measure,
    xPx,
    yPx,
    fromRow: 0,
    toRow: block.rows.length,
    headerRowCount: 0,
    context,
  });

export type TablePaintOptions = {
  readonly fragment: TableFragment;
  readonly block: TableBlock;
  readonly measure: TableMeasure;
  readonly context: BuildContext;
};

/**
 * A table fragment's primitives. A fragment that cuts a row mid-content wraps
 * its rows in a `clipGroup` and starts them `topClip` px higher, which is what
 * the painter's negative row-stack offset plus `overflow: hidden` does.
 */
export const paintTableFragment = ({
  fragment,
  block,
  measure,
  context,
}: TablePaintOptions): readonly DisplayPrimitive[] => {
  const headerRowCount = fragment.continuesFromPrev === true ? (fragment.headerRowCount ?? 0) : 0;
  let headerHeightPx = 0;
  for (let rowIndex = 0; rowIndex < headerRowCount; rowIndex += 1) {
    headerHeightPx += measure.rows[rowIndex]?.height ?? 0;
  }

  const topClipPx = fragment.topClip ?? 0;
  const body = paintTableBody({
    block,
    measure,
    xPx: fragment.x,
    yPx: fragment.y + headerHeightPx - topClipPx,
    fromRow: fragment.fromRow,
    toRow: fragment.toRow,
    headerRowCount: 0,
    context,
  });

  const headers =
    headerRowCount === 0
      ? []
      : paintTableBody({
          block,
          measure,
          xPx: fragment.x,
          yPx: fragment.y,
          fromRow: 0,
          toRow: 0,
          headerRowCount,
          context,
        });

  if (fragment.topClip === undefined && fragment.bottomClip === undefined) {
    return [...headers, ...body];
  }

  const clip: DisplayRect = {
    xPx: fragment.x,
    yPx: fragment.y + headerHeightPx,
    widthPx: fragment.width,
    heightPx: Math.max(0, fragment.height - headerHeightPx),
  };
  return [...headers, { kind: "clipGroup", rect: clip, children: body }];
};
