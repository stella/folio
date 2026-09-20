/**
 * The primitives a preview descriptor draws.
 *
 * A drawing with no image data used to be given one: a megapixel PNG built the
 * moment something was about to paint it, which the DOM backend showed as an
 * `<img>` and the PDF exporter embedded as a bitmap — a photograph of a
 * drawing that is nothing but flat rectangles.
 *
 * The descriptor already holds those rectangles, in the drawing's own
 * coordinate space, and `rect` is a primitive both backends draw natively. So
 * the preview is drawn rather than sampled: the extent maps affinely onto the
 * image's box, every shape becomes a filled rectangle inside it, and no raster
 * exists at any point. The DOM backend paints each one as the absolutely
 * positioned div it paints every other `rect` as, and the PDF backend emits
 * `re`/`f` operators, so the drawing stays a drawing at any zoom.
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
        const rect = shapeRect({ shape, extent, box });
        if (rect === undefined) {
          continue;
        }
        primitives.push({
          kind: "rect",
          rect,
          fill: parseDisplayColor(shape.color) ?? PREVIEW_SHAPE_FILL,
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
