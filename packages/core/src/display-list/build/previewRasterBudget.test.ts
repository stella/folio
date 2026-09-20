/**
 * The preview raster's bound, at the place the rasters now are.
 *
 * It used to sit on the parse, as a cap on the base64 characters a package
 * could retain, because the parse built the rasters. It builds descriptions
 * now, so a package with a thousand diagrams costs the parse a thousand
 * bounded shape lists and nothing else; what has to stay bounded is the one
 * display list that paints them.
 */

import { describe, expect, test } from "bun:test";

import { MAX_PREVIEW_PIXELS } from "../../docx/previewRaster";
import type { PreviewDescriptor } from "../../types/document";
import { ImageTable, MAX_BUILD_PREVIEW_PIXELS } from "./imagePrimitives";

const descriptor = (pixelWidth: number, pixelHeight: number): PreviewDescriptor => ({
  kind: "diagram",
  extent: { width: pixelWidth * 100, height: pixelHeight * 100 },
  shapes: [{ x: 0, y: 0, width: pixelWidth * 10, height: pixelHeight * 10, color: "70AD47" }],
  pixelWidth,
  pixelHeight,
});

describe("preview raster budget", () => {
  test("one descriptor rasterises once however often it is interned", () => {
    const table = new ImageTable();
    const one = descriptor(200, 100);
    expect(table.internPreview(one)).toBe(0);
    expect(table.internPreview(one)).toBe(0);
    expect(table.internPreview(one)).toBe(0);
    expect(table.snapshot()).toHaveLength(1);
  });

  test("two descriptors of the same size are still two pictures", () => {
    const table = new ImageTable();
    expect(table.internPreview(descriptor(200, 100))).toBe(0);
    expect(table.internPreview(descriptor(200, 100))).toBe(1);
    expect(table.snapshot()).toHaveLength(2);
  });

  test("a build's rasters are bounded however many diagrams ask for one", () => {
    // Room for three of these, offered twenty.
    const table = new ImageTable(3 * 200 * 100);
    const refs = Array.from({ length: 20 }, () => table.internPreview(descriptor(200, 100)));

    expect(refs.filter((ref) => ref !== undefined)).toHaveLength(3);
    expect(
      table.snapshot().reduce((total, source) => total + source.pixelWidth * source.pixelHeight, 0),
    ).toBe(3 * 200 * 100);
  });

  test("the allowance is spent, not merely tripped: nothing paints after it runs out", () => {
    const table = new ImageTable(200 * 100);
    expect(table.internPreview(descriptor(200, 100))).toBe(0);
    // A preview small enough to fit in what a naive remaining-count would have
    // left still gets nothing, because the allowance is gone.
    expect(table.internPreview(descriptor(20, 10))).toBeUndefined();
  });

  /**
   * Asserted on the constants rather than by rasterising, because rasterising
   * the answer means building the fourteen megapixels the question is about.
   * Ten previews at the per-preview cap is the most any package in the public
   * corpus asks one build for, and the default allowance has to clear it or
   * the bound refuses a file that works today.
   */
  test("the default allowance clears the heaviest package in the public corpus", () => {
    expect(MAX_BUILD_PREVIEW_PIXELS).toBeGreaterThanOrEqual(10 * MAX_PREVIEW_PIXELS);
  });
});
