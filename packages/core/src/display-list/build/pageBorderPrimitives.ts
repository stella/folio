/**
 * `w:pgBorders` → four strokes.
 *
 * The painter draws the border as one overlay `div` with `box-sizing:
 * border-box`, so each side's stroke lies inside the overlay box. A display
 * list stroke is centred on its path, so each side is emitted as a line half a
 * thickness inside the box edge, which puts the painted band in the same place.
 *
 * One deliberate divergence, the same one `strokes.ts` already states: the
 * painter floors a `double` border at 3px because browsers collapse a narrower
 * one into a single line. A display list has no device pixel and a PDF strokes
 * the authored width, so the authored width survives here.
 */

import type { Page } from "../../layout-engine/types";
import type { RenderPageOptions } from "../../layout-painter/renderPage";
import type { BorderSpec, Theme } from "../../types/document";
import { resolveColor } from "../../utils/colorResolver";
import { eighthsToPixels, pointsToPixels } from "../../utils/units";
import type { DisplayPrimitive, DisplayStroke } from "../types";
import type { BuildContext } from "./buildContext";
import { resolveBorderStroke } from "./strokes";
import { UNSUPPORTED_CONSTRUCT } from "./unsupported";

/** The section's `w:pgBorders`, as the painter already receives them. */
export type PageBorders = NonNullable<RenderPageOptions["pageBorders"]>;

/** `borderToStyle`: a border with no `w:sz` paints a hairline, not nothing. */
const HAIRLINE_WIDTH_PX = 1;

const SIDES = ["top", "right", "bottom", "left"] as const;

type Side = (typeof SIDES)[number];

/** `pageBorderShouldRender`: `w:display` is about the physical page number. */
const shouldRender = (pageNumber: number, display: PageBorders["display"]): boolean => {
  const rule = display ?? "allPages";
  switch (rule) {
    case "firstPage":
      return pageNumber === 1;
    case "notFirstPage":
      return pageNumber !== 1;
    case "allPages":
      return true;
    default:
      rule satisfies never;
      return true;
  }
};

const spacePx = (border: BorderSpec | undefined): number =>
  border?.space === undefined ? 0 : pointsToPixels(border.space);

const strokeOf = (
  border: BorderSpec | undefined,
  theme: Theme | null | undefined,
  context: BuildContext,
): DisplayStroke | undefined => {
  if (!border) {
    return undefined;
  }
  const { stroke, unresolvedColor } = resolveBorderStroke({
    width:
      border.size === undefined || border.size === 0
        ? HAIRLINE_WIDTH_PX
        : eighthsToPixels(border.size),
    style: border.style,
    color: resolveColor(border.color, theme),
  });
  if (unresolvedColor !== undefined) {
    context.unsupported.report(
      UNSUPPORTED_CONSTRUCT.unresolvedColor,
      context.pageIndex,
      `page border colour ${unresolvedColor}`,
    );
  }
  return stroke;
};

export type PageBorderPaintOptions = {
  readonly page: Page;
  readonly borders: PageBorders;
  readonly theme?: Theme | null;
  readonly context: BuildContext;
};

/**
 * The four page-border strokes, or nothing when this page does not carry them.
 * `offsetFrom` decides what the box is measured from: `page` insets `w:space`
 * from the paper edge, `text` insets the margin less `w:space` and the stroke
 * itself, so the authored gap between text and border survives a thick rule.
 */
export const paintPageBorders = ({
  page,
  borders,
  theme,
  context,
}: PageBorderPaintOptions): readonly DisplayPrimitive[] => {
  if (!shouldRender(page.number, borders.display)) {
    return [];
  }

  const strokes = {
    top: strokeOf(borders.top, theme, context),
    right: strokeOf(borders.right, theme, context),
    bottom: strokeOf(borders.bottom, theme, context),
    left: strokeOf(borders.left, theme, context),
  } as const satisfies Record<Side, DisplayStroke | undefined>;

  if (SIDES.every((side) => strokes[side] === undefined)) {
    return [];
  }

  const insetPx = (side: Side): number => {
    const border = borders[side];
    if ((borders.offsetFrom ?? "text") === "page") {
      return spacePx(border);
    }
    const marginPx = page.margins[side];
    return Math.max(0, marginPx - spacePx(border) - (strokes[side]?.thicknessPx ?? 0));
  };

  const leftPx = insetPx("left");
  const topPx = insetPx("top");
  const rightPx = page.size.w - insetPx("right");
  const bottomPx = page.size.h - insetPx("bottom");

  const primitives: DisplayPrimitive[] = [];
  for (const side of SIDES) {
    const stroke = strokes[side];
    if (stroke === undefined) {
      continue;
    }
    const half = stroke.thicknessPx / 2;
    switch (side) {
      case "top":
        primitives.push({
          kind: "line",
          x1Px: leftPx,
          y1Px: topPx + half,
          x2Px: rightPx,
          y2Px: topPx + half,
          stroke,
        });
        break;
      case "bottom":
        primitives.push({
          kind: "line",
          x1Px: leftPx,
          y1Px: bottomPx - half,
          x2Px: rightPx,
          y2Px: bottomPx - half,
          stroke,
        });
        break;
      case "left":
        primitives.push({
          kind: "line",
          x1Px: leftPx + half,
          y1Px: topPx,
          x2Px: leftPx + half,
          y2Px: bottomPx,
          stroke,
        });
        break;
      case "right":
        primitives.push({
          kind: "line",
          x1Px: rightPx - half,
          y1Px: topPx,
          x2Px: rightPx - half,
          y2Px: bottomPx,
          stroke,
        });
        break;
      default:
        side satisfies never;
    }
  }

  return primitives;
};
