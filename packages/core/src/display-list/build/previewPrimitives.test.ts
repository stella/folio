/**
 * What a preview descriptor draws.
 *
 * The geometry is stated as exact fractions of the image's box, because that
 * is the definition of the picture: there is no raster whose pixel grid an
 * edge could be rounded onto, so an edge that is off is off by an amount a
 * test can name.
 */

import { describe, expect, test } from "bun:test";

import type { PreviewDescriptor, PreviewShape } from "../../types/document";
import type { DisplayPrimitive, DisplayRect } from "../types";
import { paintPreview } from "./previewPrimitives";

const BOX: DisplayRect = { xPx: 100, yPx: 50, widthPx: 400, heightPx: 200 };
const EXTENT = { width: 4_000_000, height: 2_000_000 };

const descriptorOf = (shapes: readonly PreviewShape[]): PreviewDescriptor => ({
  kind: "diagram",
  extent: EXTENT,
  shapes,
});

const rectsOf = (primitives: readonly DisplayPrimitive[]) =>
  primitives.map((primitive) => {
    if (primitive.kind !== "rect") {
      throw new Error(`a preview drew a ${primitive.kind}, which no backend expects of one`);
    }
    return primitive;
  });

describe("a preview is drawn", () => {
  test("the backdrop is the first mark and fills the box", () => {
    const [backdrop] = rectsOf(paintPreview(descriptorOf([]), BOX));

    expect(backdrop?.rect).toEqual(BOX);
    expect(backdrop?.fill).toEqual({ r: 0xee, g: 0xf2, b: 0xf7, a: 1 });
    expect(backdrop?.stroke).toBeUndefined();
  });

  test("a shape covers the fraction of the box it covers of the extent", () => {
    const shape = {
      x: EXTENT.width / 4,
      y: EXTENT.height / 2,
      width: EXTENT.width / 2,
      height: EXTENT.height / 4,
      color: "70AD47",
    };
    const [, drawn] = rectsOf(paintPreview(descriptorOf([shape]), BOX));

    expect(drawn?.rect).toEqual({
      xPx: BOX.xPx + BOX.widthPx / 4,
      yPx: BOX.yPx + BOX.heightPx / 2,
      widthPx: BOX.widthPx / 2,
      heightPx: BOX.heightPx / 4,
    });
    expect(drawn?.fill).toEqual({ r: 0x70, g: 0xad, b: 0x47, a: 1 });
  });

  test("shapes are drawn in the order the drawing lists them", () => {
    const at = (x: number, color: string): PreviewShape => ({
      x,
      y: 0,
      width: EXTENT.width / 8,
      height: EXTENT.height,
      color,
    });
    const drawn = rectsOf(
      paintPreview(descriptorOf([at(0, "70AD47"), at(0, "C00000"), at(0, "4472C4")]), BOX),
    );

    expect(drawn.slice(1).map((primitive) => primitive.fill)).toEqual([
      { r: 0x70, g: 0xad, b: 0x47, a: 1 },
      { r: 0xc0, g: 0x00, b: 0x00, a: 1 },
      { r: 0x44, g: 0x72, b: 0xc4, a: 1 },
    ]);
  });

  test("a shape reaching past the extent is clipped to the box", () => {
    const [, drawn] = rectsOf(
      paintPreview(
        descriptorOf([
          {
            x: -EXTENT.width,
            y: EXTENT.height / 2,
            width: EXTENT.width * 3,
            height: EXTENT.height * 3,
            color: "70AD47",
          },
        ]),
        BOX,
      ),
    );

    expect(drawn?.rect).toEqual({
      xPx: BOX.xPx,
      yPx: BOX.yPx + BOX.heightPx / 2,
      widthPx: BOX.widthPx,
      heightPx: BOX.heightPx / 2,
    });
  });

  test("a shape wholly outside the extent draws nothing", () => {
    const outside = paintPreview(
      descriptorOf([
        { x: EXTENT.width, y: 0, width: EXTENT.width, height: EXTENT.height, color: "70AD47" },
      ]),
      BOX,
    );

    expect(outside).toHaveLength(1);
  });

  /**
   * The raster read each channel with `parseInt(...) || fallback`, so a zero
   * channel took the backdrop's value: pure red came out yellow-green and
   * black came out pale blue. A drawn rectangle keeps the colour the drawing
   * authored.
   */
  test("a channel of zero is a channel of zero", () => {
    const [, drawn] = rectsOf(
      paintPreview(
        descriptorOf([{ x: 0, y: 0, width: EXTENT.width, height: EXTENT.height, color: "FF0000" }]),
        BOX,
      ),
    );

    expect(drawn?.fill).toEqual({ r: 255, g: 0, b: 0, a: 1 });
  });

  test("a colour the drawing did not state falls back to the preview fill", () => {
    const [, drawn] = rectsOf(
      paintPreview(
        descriptorOf([{ x: 0, y: 0, width: EXTENT.width, height: EXTENT.height, color: "" }]),
        BOX,
      ),
    );

    expect(drawn?.fill).toEqual({ r: 0xe8, g: 0xee, b: 0xf7, a: 1 });
  });

  test("an extent with no area leaves the backdrop alone in the box it reserved", () => {
    const degenerate = paintPreview(
      {
        kind: "diagram",
        extent: { width: 0, height: 0 },
        shapes: [{ x: 0, y: 0, width: 10, height: 10, color: "70AD47" }],
      },
      BOX,
    );

    expect(rectsOf(degenerate)).toHaveLength(1);
    expect(rectsOf(degenerate).at(0)?.rect).toEqual(BOX);
  });
});
