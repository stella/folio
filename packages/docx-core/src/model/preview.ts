/**
 * How a drawing with no image data is drawn, without drawing it.
 *
 * A `w:drawing` whose graphic is a diagram carries no `a:blip`, so there is no
 * image to resolve and the model has nothing to say the drawing looks like.
 * Rasterising a placeholder at parse time answers that, at the cost of holding
 * a megapixel PNG — megabytes per drawing — for the life of the document,
 * whether or not anything ever paints it.
 *
 * A descriptor says the same thing as a list of flat rectangles in the
 * drawing's own coordinate space. It is hundreds of bytes against megabytes,
 * and it is what a renderer draws: no picture is built at parse time or at
 * paint time, so a backend with vector operators draws vectors.
 */

/** One flat rectangle of a preview, in the coordinate space of {@link PreviewDescriptor.extent}. */
export type PreviewShape = {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
  /** Six hexadecimal digits, no leading `#`, as `a:srgbClr@val` spells it. */
  readonly color: string;
};

/**
 * Discriminated on `kind` rather than shaped as a diagram record: chart frames
 * and OLE previews are the same problem and will want the same field, and a
 * `kind` now is cheaper than migrating a bare `diagram` later.
 */
export type PreviewDescriptor = {
  readonly kind: "diagram";
  /**
   * The drawing's extent, in the units `shapes` coordinates use (EMU).
   *
   * Held here rather than read from the owning `Image.size` so a descriptor
   * that reaches a renderer alone still maps its shapes onto the box it is
   * drawn in.
   */
  readonly extent: { readonly width: number; readonly height: number };
  /** Already clipped to the shape cap, in paint order. */
  readonly shapes: readonly PreviewShape[];
};
