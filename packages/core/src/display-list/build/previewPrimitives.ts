/**
 * The primitives a preview descriptor draws.
 *
 * A drawing with no image data used to be given one: a megapixel PNG built the
 * moment something was about to paint it, which the DOM backend showed as an
 * `<img>` and the PDF exporter embedded as a bitmap — a photograph of a
 * drawing that is vector geometry.
 *
 * The descriptor already holds rectangles, polygons, and polylines in the
 * drawing's own coordinate space, and each maps to native display-list
 * primitives. The preview is drawn rather than sampled: the extent maps
 * affinely onto the image's box, and no raster exists at any point, so the
 * drawing stays a drawing at any zoom.
 */

import { panic } from "better-result";

import type { PreviewDescriptor, PreviewShape } from "../../types/document";
import type { DisplayColor, DisplayPrimitive, DisplayRect } from "../types";
import { parseDisplayColor } from "./colors";

/** The flat backdrop a preview is drawn on, behind every shape. */
const PREVIEW_BACKGROUND: DisplayColor = { r: 0xee, g: 0xf2, b: 0xf7, a: 1 };

/** A shape whose `a:srgbClr` was missing or is not a colour. */
const PREVIEW_SHAPE_FILL: DisplayColor = { r: 0xe8, g: 0xee, b: 0xf7, a: 1 };

const clampFraction = (value: number): number => Math.min(1, Math.max(0, value));

type ShapeRectOptions = {
  readonly shape: PreviewShape;
  readonly extent: PreviewDescriptor["extent"];
  readonly box: DisplayRect;
};

/**
 * One shape's box, as the fraction of the extent it covers applied to the
 * image's box. `undefined` for a shape the box keeps nothing of, which is the
 * clip the raster did by clamping to its own bounds.
 */
const shapeRect = ({ shape, extent, box }: ShapeRectOptions): DisplayRect | undefined => {
  if (shape.type !== "rect") return undefined;
  const left = clampFraction(shape.x / extent.width);
  const right = clampFraction((shape.x + shape.width) / extent.width);
  const top = clampFraction(shape.y / extent.height);
  const bottom = clampFraction((shape.y + shape.height) / extent.height);
  if (right <= left || bottom <= top) {
    return undefined;
  }
  return {
    xPx: box.xPx + left * box.widthPx,
    yPx: box.yPx + top * box.heightPx,
    widthPx: (right - left) * box.widthPx,
    heightPx: (bottom - top) * box.heightPx,
  };
};

/**
 * The marks a descriptor makes inside `box`, back to front: the backdrop, then
 * each shape in the order the drawing lists them.
 */
export const paintPreview = (
  descriptor: PreviewDescriptor,
  box: DisplayRect,
): readonly DisplayPrimitive[] => {
  switch (descriptor.kind) {
    case "chart":
    case "diagram": {
      const primitives: DisplayPrimitive[] = [
        { kind: "rect", rect: box, fill: PREVIEW_BACKGROUND },
      ];
      const { extent } = descriptor;
      // A descriptor reaches a build from persisted editor state as well as
      // from a parse, and an extent with no area maps every shape to nothing.
      // The backdrop still occupies the box the drawing reserved.
      if (!(extent.width > 0 && extent.height > 0)) {
        return primitives;
      }
      for (const shape of descriptor.shapes) {
        const color = parseDisplayColor(shape.color) ?? PREVIEW_SHAPE_FILL;
        if (shape.type === "polygon") {
          primitives.push({
            kind: "polygon",
            points: shape.points.map(({ x, y }) => ({
              xPx: box.xPx + clampFraction(x / extent.width) * box.widthPx,
              yPx: box.yPx + clampFraction(y / extent.height) * box.heightPx,
            })),
            fill: color,
          });
          continue;
        }
        if (shape.type === "polyline") {
          const points = shape.points.map(({ x, y }) => ({
            xPx: box.xPx + clampFraction(x / extent.width) * box.widthPx,
            yPx: box.yPx + clampFraction(y / extent.height) * box.heightPx,
          }));
          const stroke = {
            color,
            thicknessPx:
              shape.strokeWidth *
              Math.min(box.widthPx / extent.width, box.heightPx / extent.height),
            pattern: "solid" as const,
          };
          for (let index = 1; index < points.length; index += 1)
            primitives.push({
              kind: "line",
              x1Px: points[index - 1]!.xPx,
              y1Px: points[index - 1]!.yPx,
              x2Px: points[index]!.xPx,
              y2Px: points[index]!.yPx,
              stroke,
            });
          continue;
        }
        const rect = shapeRect({ shape, extent, box });
        if (rect === undefined) {
          continue;
        }
        primitives.push({
          kind: "rect",
          rect,
          fill: color,
        });
      }
      return primitives;
    }
    default: {
      const unreachable: never = descriptor.kind;
      return panic(`unhandled preview kind: ${String(unreachable)}`);
    }
  }
};
