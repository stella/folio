import type { ImageRun } from "../layout-engine/types";
import { emuToPixels } from "./renderUtils";

export type PageGeometry = {
  pageWidth: number;
  pageHeight: number;
  marginLeft: number;
  marginTop: number;
  marginRight: number;
  marginBottom: number;
  contentWidth: number;
  contentHeight: number;
};

/** Coordinates relative to the page content area's top-left corner. */
export type AnchoredImagePosition = {
  x: number;
  y: number;
  side: "left" | "right";
};

function resolveHorizontalBand(
  relativeTo: string | undefined,
  geometry: PageGeometry,
): { baseX: number; bandWidth: number } {
  switch (relativeTo) {
    case "page":
      return { baseX: -geometry.marginLeft, bandWidth: geometry.pageWidth };
    case "leftMargin":
    case "insideMargin":
      return { baseX: -geometry.marginLeft, bandWidth: geometry.marginLeft };
    case "rightMargin":
    case "outsideMargin":
      return { baseX: geometry.contentWidth, bandWidth: geometry.marginRight };
    case "character":
      return { baseX: 0, bandWidth: 0 };
    default:
      return { baseX: 0, bandWidth: geometry.contentWidth };
  }
}

function resolveVerticalBand(
  relativeTo: string | undefined,
  fragmentY: number,
  geometry: PageGeometry,
): { baseY: number; bandHeight: number } {
  switch (relativeTo) {
    case "page":
      return { baseY: -geometry.marginTop, bandHeight: geometry.pageHeight };
    case "topMargin":
      return { baseY: -geometry.marginTop, bandHeight: geometry.marginTop };
    case "bottomMargin":
      return { baseY: geometry.contentHeight, bandHeight: geometry.marginBottom };
    case "paragraph":
    case "line":
      return { baseY: fragmentY, bandHeight: 0 };
    default:
      return { baseY: 0, bandHeight: geometry.contentHeight };
  }
}

/** Resolve an anchored image into page-content-relative coordinates. */
export function resolveAnchoredImagePosition(
  imgRun: ImageRun,
  fragmentY: number,
  geometry: PageGeometry,
): AnchoredImagePosition {
  const position = imgRun.position;
  const contentWidth = geometry.contentWidth;

  let side: "left" | "right" = "left";
  let x = 0;

  if (position?.horizontal) {
    const h = position.horizontal;
    const { baseX, bandWidth } = resolveHorizontalBand(h.relativeTo, geometry);
    const horizontalHasBand = h.relativeTo !== "character";

    if (h.align === "right" || h.align === "outside") {
      side = "right";
      x = horizontalHasBand ? baseX + bandWidth - imgRun.width : 0;
    } else if (h.align === "left" || h.align === "inside") {
      x = baseX;
    } else if (h.align === "center") {
      x = horizontalHasBand ? baseX + (bandWidth - imgRun.width) / 2 : 0;
    } else if (h.posOffset !== undefined) {
      x = baseX + emuToPixels(h.posOffset);
      side = x > contentWidth / 2 ? "right" : "left";
    } else {
      x = baseX;
    }
  } else if (imgRun.cssFloat === "right") {
    side = "right";
    x = contentWidth - imgRun.width;
  }

  let y: number;
  if (position?.vertical) {
    const v = position.vertical;
    const { baseY, bandHeight } = resolveVerticalBand(v.relativeTo, fragmentY, geometry);
    const verticalHasBand = v.relativeTo !== "paragraph" && v.relativeTo !== "line";

    if (v.align === "top" || v.align === "inside") {
      y = baseY;
    } else if (v.align === "center") {
      y = verticalHasBand ? baseY + (bandHeight - imgRun.height) / 2 : fragmentY;
    } else if (v.align === "bottom" || v.align === "outside") {
      y = verticalHasBand ? baseY + bandHeight - imgRun.height : fragmentY;
    } else if (v.posOffset !== undefined) {
      y = baseY + emuToPixels(v.posOffset);
    } else {
      y = verticalHasBand ? baseY : fragmentY;
    }
  } else {
    y = fragmentY;
  }

  return { x, y, side };
}
