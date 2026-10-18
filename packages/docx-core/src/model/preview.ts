/**
 * How a drawing with no image data is drawn, without drawing it.
 *
 * A `w:drawing` whose graphic is a diagram carries no `a:blip`, so there is no
 * image to resolve and the model has nothing to say the drawing looks like.
 * Rasterising a placeholder at parse time answers that, at the cost of holding
 * a megapixel PNG — megabytes per drawing — for the life of the document,
 * whether or not anything ever paints it.
 *
 * A descriptor says the same thing as a list of filled rectangles, filled
 * polygons, and stroked polylines in the drawing's own coordinate space. It
 * is hundreds of bytes against megabytes, and it is what a renderer draws: no
 * picture is built at parse time or at paint time, so a backend with vector
 * operators draws vectors.
 */

/** One point in the coordinate space of {@link PreviewDescriptor.extent}. */
export type PreviewPoint = { readonly x: number; readonly y: number };

export type PreviewShape =
  | {
      readonly type: "rect";
      readonly x: number;
      readonly y: number;
      readonly width: number;
      readonly height: number;
      /** Six hexadecimal digits, no leading `#`, as `a:srgbClr@val` spells it. */
      readonly color: string;
    }
  | {
      readonly type: "polygon";
      readonly points: readonly PreviewPoint[];
      readonly color: string;
    }
  | {
      readonly type: "polyline";
      readonly points: readonly PreviewPoint[];
      readonly color: string;
      readonly strokeWidth: number;
    };

/**
 * The kind identifies the drawing source while the renderers share the same
 * bounded vector primitives for every source.
 */
export type PreviewDescriptor = {
  readonly kind: "chart" | "diagram";
  /**
   * The drawing's extent, in the units `shapes` coordinates use (EMU).
   *
   * Held here rather than read from the owning `Image.size` so a descriptor
   * that reaches a renderer alone still maps its shapes onto the box it is
   * drawn in.
   */
  readonly extent: { readonly width: number; readonly height: number };
  /** Already bounded by the shape and point caps, in paint order. */
  readonly shapes: readonly PreviewShape[];
};
