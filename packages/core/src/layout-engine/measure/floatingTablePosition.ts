/**
 * Horizontal placement of a `w:tblpPr` positioned table.
 *
 * Wrap-zone extraction and final floating-table placement must agree on X.
 * Resolution order per `w:tblpPr` (§17.4.57): a `tblpXSpec` keyword supersedes
 * any `tblpX` offset ("that value is ignored"), then the explicit `tblpX`
 * offset, then the table's own `w:jc` justification, then the left margin.
 * `inside`/`outside` resolve as `left`/`right` — the editor does not model
 * facing pages. The result stays signed when the table is wider than its
 * anchor frame; final placement owns the physical-page clamp.
 *
 * @packageDocumentation
 */

import type { FloatingTablePosition, TableBlock } from "../types";

type FloatingTablePageXOptions = {
  anchor: FloatingTablePosition;
  justification: TableBlock["justification"];
  tableWidth: number;
  marginWidth: number;
  pageWidth: number;
  marginLeft: number;
  textFrameWidth: number;
  textFrameLeft: number;
};

/**
 * X of the table's left edge relative to the content-box left edge, px.
 */
export function resolveFloatingTableX(
  anchor: FloatingTablePosition,
  justification: TableBlock["justification"],
  tableWidth: number,
  contentWidth: number,
): number {
  const spec = anchor.tblpXSpec;
  if (spec === "left" || spec === "inside") {
    return 0;
  }
  if (spec === "right" || spec === "outside") {
    return contentWidth - tableWidth;
  }
  if (spec === "center") {
    return (contentWidth - tableWidth) / 2;
  }
  if (anchor.tblpX !== undefined) {
    return anchor.tblpX;
  }
  if (justification === "center") {
    return (contentWidth - tableWidth) / 2;
  }
  if (justification === "right") {
    return contentWidth - tableWidth;
  }
  return 0;
}

/**
 * Absolute page X after resolving the selected anchor frame and constraining
 * the table to the physical page.
 */
export function resolveFloatingTablePageX({
  anchor,
  justification,
  tableWidth,
  marginWidth,
  pageWidth,
  marginLeft,
  textFrameWidth,
  textFrameLeft,
}: FloatingTablePageXOptions): number {
  let frameWidth = marginWidth;
  let frameLeft = marginLeft;
  switch (anchor.horzAnchor) {
    case "page":
      frameWidth = pageWidth;
      frameLeft = 0;
      break;
    case "text":
      frameWidth = textFrameWidth;
      frameLeft = textFrameLeft;
      break;
    case "margin":
    case undefined:
      break;
    default:
      anchor.horzAnchor satisfies never;
  }
  const resolvedX =
    frameLeft + resolveFloatingTableX(anchor, justification, tableWidth, frameWidth);
  const maxX = Math.max(0, pageWidth - tableWidth);
  return Math.max(0, Math.min(resolvedX, maxX));
}
