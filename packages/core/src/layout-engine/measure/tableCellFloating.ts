import { emuToPixels } from "../../utils/units";
import { createTableCellFlowState, placeTableCellBlock } from "./tableCellFlow";
import {
  isFloatingImageRun,
  resolveTableCellPadding,
  type ImageRun,
  type TableCell,
  type TableCellMeasure,
} from "../types";
import { clampFloatingWrapMargins } from "./clampFloatingWrapMargins";
import type { FloatingImageZone } from "./floatingZones";

export type TableCellFloatingImage = {
  src: string;
  width: number;
  height: number;
  alt?: string;
  transform?: string;
  opacity?: number;
  brightness?: number;
  contrast?: number;
  cropTop?: number;
  cropRight?: number;
  cropBottom?: number;
  cropLeft?: number;
  x: number;
  y: number;
  side: "left" | "right";
  distTop: number;
  distBottom: number;
  distLeft: number;
  distRight: number;
  wrapText?: "bothSides" | "left" | "right" | "largest";
  /** Authored wrap mode; decides whether and how the picture excludes text. */
  wrapType?: ImageRun["wrapType"];
  pmStart?: number;
  pmEnd?: number;
};

type TableCellFloatingPosition = Pick<TableCellFloatingImage, "x" | "y" | "side">;

export type ResolveTableCellFloatingPosition = (
  run: ImageRun,
  paragraphY: number,
) => TableCellFloatingPosition;

type ResolveCellScopedPositionOptions = {
  run: ImageRun;
  paragraphY: number;
  contentWidth: number;
};

const resolveCellScopedPosition = ({
  run,
  paragraphY,
  contentWidth,
}: ResolveCellScopedPositionOptions): TableCellFloatingPosition => {
  const position = run.position;
  let side: "left" | "right" = "left";
  let x = 0;
  if (position?.horizontal) {
    const horizontal = position.horizontal;
    if (horizontal.align === "right") {
      side = "right";
      x = contentWidth - run.width;
    } else if (horizontal.align === "center") {
      x = (contentWidth - run.width) / 2;
    } else if (horizontal.posOffset !== undefined) {
      x = emuToPixels(horizontal.posOffset);
      side = x > contentWidth / 2 ? "right" : "left";
    }
  } else if (run.cssFloat === "right") {
    side = "right";
    x = contentWidth - run.width;
  }

  let y = paragraphY;
  if (position?.vertical) {
    const vertical = position.vertical;
    if (vertical.posOffset !== undefined) {
      y = paragraphY + emuToPixels(vertical.posOffset);
    } else if (vertical.align === "top") {
      y = 0;
    }
  }

  return {
    side,
    x,
    y,
  };
};

export function getTableCellContentWidth(
  cell: TableCell | undefined,
  cellMeasure: TableCellMeasure,
): number {
  const { left: padLeft, right: padRight } = resolveTableCellPadding(cell);
  return Math.max(0, cellMeasure.width - padLeft - padRight);
}

export function getTableCellFloatingImages(
  cell: TableCell,
  cellMeasure: TableCellMeasure,
  contentWidth: number,
  resolvePosition?: ResolveTableCellFloatingPosition,
): TableCellFloatingImage[] {
  const result: TableCellFloatingImage[] = [];
  const flowState = createTableCellFlowState();

  for (let blockIndex = 0; blockIndex < cell.blocks.length; blockIndex++) {
    const block = cell.blocks[blockIndex];
    const blockMeasure = cellMeasure.blocks[blockIndex];
    if (!block || !blockMeasure) {
      continue;
    }
    const placement = placeTableCellBlock(flowState, block, blockMeasure);
    if (block.kind !== "paragraph") {
      continue;
    }

    for (const run of block.runs) {
      if (run.kind !== "image" || !isFloatingImageRun(run)) {
        continue;
      }

      const distTop = run.distTop ?? 0;
      const distBottom = run.distBottom ?? 0;
      const distLeft = run.distLeft ?? 12;
      const distRight = run.distRight ?? 12;
      const verticalOriginY =
        run.position?.vertical?.relativeTo === "paragraph" ? placement.top : placement.contentTop;

      if (run.layoutInCell === false && !resolvePosition) {
        // Anchored relative to the page rather than this cell, but no page
        // geometry is available here (e.g. row-break measurement, selection
        // rects). Approximating it as cell-scoped would disagree with the
        // page-relative position the painter resolves once page geometry is
        // known, so leave it out of this cell's wrap/break geometry instead.
        continue;
      }

      const resolved =
        run.layoutInCell === false && resolvePosition
          ? resolvePosition(run, verticalOriginY)
          : resolveCellScopedPosition({
              run,
              paragraphY: verticalOriginY,
              contentWidth,
            });

      let wrapText: "bothSides" | "left" | "right" | "largest" = "bothSides";
      if (run.cssFloat === "left") {
        wrapText = "right";
      } else if (run.cssFloat === "right") {
        wrapText = "left";
      }

      result.push({
        src: run.src,
        width: run.width,
        height: run.height,
        ...(run.alt !== undefined ? { alt: run.alt } : {}),
        ...(run.transform !== undefined ? { transform: run.transform } : {}),
        ...(run.opacity != null ? { opacity: run.opacity } : {}),
        ...(run.brightness != null ? { brightness: run.brightness } : {}),
        ...(run.contrast != null ? { contrast: run.contrast } : {}),
        ...(run.cropTop != null ? { cropTop: run.cropTop } : {}),
        ...(run.cropRight != null ? { cropRight: run.cropRight } : {}),
        ...(run.cropBottom != null ? { cropBottom: run.cropBottom } : {}),
        ...(run.cropLeft != null ? { cropLeft: run.cropLeft } : {}),
        x: resolved.x,
        y: resolved.y,
        side: resolved.side,
        distTop,
        distBottom,
        distLeft,
        distRight,
        wrapText,
        ...(run.wrapType !== undefined ? { wrapType: run.wrapType } : {}),
        ...(run.pmStart !== undefined ? { pmStart: run.pmStart } : {}),
        ...(run.pmEnd !== undefined ? { pmEnd: run.pmEnd } : {}),
      });
    }
  }

  return result;
}

/**
 * The text exclusion one cell-scoped picture casts, in cell content
 * coordinates, or `undefined` when text ignores it.
 *
 * - `wp:wrapNone` (`behind` / `inFront`): text paints over or under the
 *   picture, so nothing is excluded.
 * - `wp:wrapTopAndBottom`: no text beside the picture; lines it overlaps move
 *   below it.
 * - `wp:wrapSquare` and the contour modes: text runs on the picture's open
 *   side. When the picture and its wrap distances cover the cell's content
 *   width there is no open side, and the text moves below the picture as for
 *   `wrapTopAndBottom`.
 */
function tableCellFloatingZone(
  img: TableCellFloatingImage,
  contentWidth: number,
): FloatingImageZone | undefined {
  if (img.wrapType === "behind" || img.wrapType === "inFront") {
    return undefined;
  }
  const rectLeft = img.x - img.distLeft;
  const rectRight = img.x + img.width + img.distRight;
  const topY = img.y - img.distTop;
  const bottomY = img.y + img.height + img.distBottom;
  const belowPicture: FloatingImageZone = {
    leftMargin: 0,
    rightMargin: 0,
    topY,
    bottomY,
    fullWidthBlock: true,
  };
  if (img.wrapType === "topAndBottom") {
    return belowPicture;
  }

  let leftMargin = 0;
  let rightMargin = 0;
  const wrapText = img.wrapText ?? "bothSides";
  if (wrapText === "right") {
    leftMargin = rectRight;
  } else if (wrapText === "left") {
    rightMargin = contentWidth - rectLeft;
  } else if (img.side === "left") {
    leftMargin = rectRight;
  } else {
    rightMargin = contentWidth - rectLeft;
  }

  const overlapsContent = rectRight > 0 && rectLeft < contentWidth;
  if (overlapsContent && Math.max(leftMargin, rightMargin) >= contentWidth) {
    return belowPicture;
  }

  const clamped = clampFloatingWrapMargins(leftMargin, rightMargin, contentWidth);
  return {
    leftMargin: clamped.leftMargin,
    rightMargin: clamped.rightMargin,
    topY,
    bottomY,
  };
}

export function buildTableCellFloatingZones(
  floatingImages: TableCellFloatingImage[],
  contentWidth: number,
): FloatingImageZone[] {
  const zones: FloatingImageZone[] = [];
  for (const img of floatingImages) {
    const zone = tableCellFloatingZone(img, contentWidth);
    if (zone !== undefined) {
      zones.push(zone);
    }
  }
  return zones;
}
