/**
 * What a preview descriptor draws, and that it draws what the raster showed.
 *
 * The geometry is stated as exact fractions of the image's box, because that
 * is now the definition of the picture rather than a consequence of how many
 * pixels a raster happened to have. The second half of the file compares those
 * rectangles against the raster they replace, decoding it rather than
 * restating its arithmetic: a second copy of the formula would agree with
 * itself however wrong both copies were.
 */

import { inflateSync } from "node:zlib";

import { describe, expect, test } from "bun:test";

import { previewRasterSize, rasterizePreview } from "../../docx/previewRaster";
import type { PreviewDescriptor, PreviewShape } from "../../types/document";
import type { DisplayPrimitive, DisplayRect } from "../types";
import { paintPreview } from "./previewPrimitives";

const BOX: DisplayRect = { xPx: 100, yPx: 50, widthPx: 400, heightPx: 200 };
const EXTENT = { width: 4_000_000, height: 2_000_000 };

const descriptorOf = (shapes: readonly PreviewShape[]): PreviewDescriptor => ({
  kind: "diagram",
  extent: EXTENT,
  shapes,
  ...previewRasterSize(EXTENT.width, EXTENT.height),
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
        pixelWidth: 1,
        pixelHeight: 1,
      },
      BOX,
    );

    expect(rectsOf(degenerate)).toHaveLength(1);
    expect(rectsOf(degenerate).at(0)?.rect).toEqual(BOX);
  });
});

// ---------------------------------------------------------------------------
// The raster this replaces
// ---------------------------------------------------------------------------

type Raster = {
  readonly width: number;
  readonly height: number;
  readonly pixels: Uint8Array;
};

/**
 * The preview PNG's samples. It is always 8-bit RGBA with no interlacing and
 * one filter-zero scanline per row, so reversing it needs no filter machinery.
 */
const decodePreviewPng = (bytes: Uint8Array): Raster => {
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const width = view.getUint32(16);
  const height = view.getUint32(20);
  let cursor = 8;
  const payload: Uint8Array[] = [];
  while (cursor + 8 <= bytes.length) {
    const length = view.getUint32(cursor);
    const type = String.fromCharCode(...bytes.subarray(cursor + 4, cursor + 8));
    if (type === "IDAT") {
      payload.push(bytes.subarray(cursor + 8, cursor + 8 + length));
    }
    cursor += 12 + length;
  }
  const joined = new Uint8Array(payload.reduce((total, chunk) => total + chunk.length, 0));
  let offset = 0;
  for (const chunk of payload) {
    joined.set(chunk, offset);
    offset += chunk.length;
  }
  const rows = inflateSync(joined);
  const pixels = new Uint8Array(width * height * 4);
  for (let y = 0; y < height; y += 1) {
    pixels.set(rows.subarray(y * (width * 4 + 1) + 1, (y + 1) * (width * 4 + 1)), y * width * 4);
  }
  return { width, height, pixels };
};

/** The bounding box, in raster pixels, of every pixel of one colour. */
const rasterBoundsOf = (raster: Raster, color: readonly [number, number, number]) => {
  let minX = raster.width;
  let minY = raster.height;
  let maxX = 0;
  let maxY = 0;
  for (let y = 0; y < raster.height; y += 1) {
    for (let x = 0; x < raster.width; x += 1) {
      const index = (y * raster.width + x) * 4;
      if (
        raster.pixels[index] === color[0] &&
        raster.pixels[index + 1] === color[1] &&
        raster.pixels[index + 2] === color[2]
      ) {
        minX = Math.min(minX, x);
        minY = Math.min(minY, y);
        maxX = Math.max(maxX, x + 1);
        maxY = Math.max(maxY, y + 1);
      }
    }
  }
  return { minX, minY, maxX, maxY };
};

/**
 * The raster quantised every edge to a whole pixel of its own grid, so the
 * drawn rectangles can only agree with it to that grid. One pixel of the
 * raster, measured in the box, is the tolerance; anything larger would be the
 * rectangles landing somewhere else.
 */
describe("a drawn preview matches the raster it replaces", () => {
  const shapes: readonly PreviewShape[] = [
    { x: 200_000, y: 100_000, width: 1_400_000, height: 700_000, color: "70AD47" },
    { x: 2_000_000, y: 1_000_000, width: 1_600_000, height: 800_000, color: "4472C4" },
  ];
  const descriptor = descriptorOf(shapes);
  const raster = decodePreviewPng(rasterizePreview(descriptor));
  const drawn = rectsOf(paintPreview(descriptor, BOX));

  test("the descriptor's raster is the one the shapes were clipped into", () => {
    expect(raster.width).toBe(descriptor.pixelWidth);
    expect(raster.height).toBe(descriptor.pixelHeight);
  });

  test.each([
    [0, [0x70, 0xad, 0x47]],
    [1, [0x44, 0x72, 0xc4]],
  ] as const)("shape %i occupies the rectangle the raster painted", (index, color) => {
    const bounds = rasterBoundsOf(raster, color);
    const rect = drawn.at(index + 1)?.rect;
    const xTolerancePx = BOX.widthPx / raster.width;
    const yTolerancePx = BOX.heightPx / raster.height;
    const edges = [
      ["left", rect?.xPx, BOX.xPx + (bounds.minX / raster.width) * BOX.widthPx, xTolerancePx],
      ["top", rect?.yPx, BOX.yPx + (bounds.minY / raster.height) * BOX.heightPx, yTolerancePx],
      [
        "width",
        rect?.widthPx,
        ((bounds.maxX - bounds.minX) / raster.width) * BOX.widthPx,
        xTolerancePx,
      ],
      [
        "height",
        rect?.heightPx,
        ((bounds.maxY - bounds.minY) / raster.height) * BOX.heightPx,
        yTolerancePx,
      ],
    ] as const;

    for (const [edge, actual, expected, tolerancePx] of edges) {
      expect(Math.abs((actual ?? Number.NaN) - expected), `${edge} edge`).toBeLessThanOrEqual(
        tolerancePx,
      );
    }
  });
});
