/**
 * What a package with diagrams costs to parse, and what it retains.
 *
 * The SmartArt preview used to be a raster the parse built, so every defect in
 * it was a defect per byte and the package's cost followed the extent its
 * author chose rather than anything the document contained. The parse now
 * builds a description instead, and the raster is built by whoever paints it.
 *
 * These tests state that as properties of the code rather than of the machine,
 * because a millisecond budget cannot say it on shared hardware. Two stand in
 * for time and memory: the number of iterator steps a parse takes over byte
 * buffers, which is zero once nothing walks a raster, and the number of bytes
 * a package retains, which must follow its diagram count and not its extent.
 */

import { describe, expect, test } from "bun:test";

import type { MediaFile, PreviewDescriptor } from "../types/document";
import { ImageTable } from "../display-list/build/imagePrimitives";
import { parseDiagramPreview } from "./diagramPreview";
import { MAX_PREVIEW_PIXELS } from "./previewRaster";
import { parseRelationships } from "./relsParser";
import { parseXmlDocument } from "./xmlParser";

const DIAGRAM_RELATIONSHIPS = parseRelationships(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/></Relationships>`,
);

/** A diagram drawing whose raster is `extent` by `extent` device pixels. */
const diagramDrawing = (extent: number) => {
  const drawing = parseXmlDocument(
    `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="${String(extent)}" cy="${String(extent)}"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic></wp:inline></w:drawing>`,
  );
  if (!drawing) {
    throw new Error("diagram fixture did not parse");
  }
  return drawing;
};

/** `shapeCount` filled rectangles for the drawing part the preview reads. */
const diagramMedia = (shapeCount: number): Map<string, MediaFile> => {
  const shapes = Array.from(
    { length: shapeCount },
    (_unused, index) =>
      `<dsp:sp><dsp:spPr><a:xfrm><a:off x="${String(index * 10)}" y="${String(index * 10)}"/><a:ext cx="400" cy="200"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr></dsp:sp>`,
  ).join("");
  const drawingPart = new TextEncoder().encode(
    `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree>${shapes}</dsp:spTree></dsp:drawing>`,
  ).buffer;
  return new Map([
    [
      "word/diagrams/drawing1.xml",
      {
        path: "word/diagrams/drawing1.xml",
        filename: "drawing1.xml",
        mimeType: "application/xml",
        data: drawingPart,
      },
    ],
  ]);
};

/**
 * Count the iterator steps a call takes over byte buffers.
 *
 * `Uint8Array.prototype[Symbol.iterator]` is the one interception point that
 * needs no production seam, and it is the exact protocol a per-byte `for...of`
 * drives, so the counter is zero for code that indexes and proportional to the
 * buffer for code that does not.
 */
const countByteIteratorSteps = (render: () => void): number => {
  const prototype = Object.getPrototypeOf(new Uint8Array()) as {
    [Symbol.iterator]: () => IterableIterator<number>;
  };
  const original = prototype[Symbol.iterator];
  let steps = 0;
  prototype[Symbol.iterator] = function counted(this: Uint8Array) {
    const inner = original.call(this);
    return {
      [Symbol.iterator]() {
        return this;
      },
      next() {
        steps += 1;
        return inner.next();
      },
    } as IterableIterator<number>;
  };
  try {
    render();
  } finally {
    prototype[Symbol.iterator] = original;
  }
  return steps;
};

/**
 * Enough headroom for the handful of whole-buffer walks the assembly does
 * (chunk concatenation and the like), and far below any per-pixel figure: the
 * smallest raster these tests render is already ten thousand pixels.
 */
const MAX_ITERATOR_STEPS_PER_PREVIEW = 1_000;

/**
 * What one diagram may add to a parsed package, in bytes of retained JSON.
 *
 * A descriptor is a bounded shape list, so this is a constant rather than a
 * function of the drawing: the cap is the shape cap times a generous per-shape
 * figure. The number to compare it against is what the same preview used to
 * retain, which was the base64 of a megapixel raster: roughly 7.3 MB, five
 * hundred times this budget and unrelated to the package's own size.
 */
const MAX_RETAINED_BYTES_PER_DIAGRAM = 16_384;

const descriptorBytes = (descriptor: PreviewDescriptor): number =>
  JSON.stringify(descriptor).length;

const parsePreview = (extent: number, media: Map<string, MediaFile>): PreviewDescriptor => {
  const preview = parseDiagramPreview(
    diagramDrawing(extent),
    DIAGRAM_RELATIONSHIPS,
    media,
  )?.preview;
  if (!preview) {
    throw new Error("diagram fixture produced no preview");
  }
  return preview;
};

describe("SmartArt preview scaling", () => {
  test("a parse takes no iterator steps over byte buffers, at any raster size", () => {
    const media = diagramMedia(4);
    const counts = [200, 400, 800, 1600].map((extent) =>
      countByteIteratorSteps(() => {
        expect(parsePreview(extent, media).kind).toBe("diagram");
      }),
    );

    // Sixty-four times the pixels must not buy a single extra step.
    for (const steps of counts) {
      expect(steps).toBeLessThan(MAX_ITERATOR_STEPS_PER_PREVIEW);
    }
    expect(new Set(counts).size).toBe(1);
  });

  test("what a parse retains per diagram is bounded and does not follow the extent", () => {
    const media = diagramMedia(4);
    const sizes = [200, 400, 800, 1600, 100_000].map((extent) =>
      descriptorBytes(parsePreview(extent, media)),
    );
    for (const size of sizes) {
      expect(size).toBeLessThan(MAX_RETAINED_BYTES_PER_DIAGRAM);
    }

    // Five hundred times the area, and the only difference is the digits in
    // the two recorded raster dimensions.
    expect(Math.max(...sizes) - Math.min(...sizes)).toBeLessThan(32);
  });

  test("a package's retained preview bytes stay proportional to its diagram count", () => {
    const media = diagramMedia(4);
    const single = descriptorBytes(parsePreview(800, media));

    for (const diagrams of [1, 2, 4, 8, 64]) {
      const total = Array.from({ length: diagrams }, () => parsePreview(800, media)).reduce(
        (bytes, preview) => bytes + descriptorBytes(preview),
        0,
      );
      expect(total).toBe(single * diagrams);
      expect(total).toBeLessThan(MAX_RETAINED_BYTES_PER_DIAGRAM * diagrams);
    }
  });

  test("the raster the parse no longer builds is still built when a backend asks", () => {
    const table = new ImageTable();
    const steps = countByteIteratorSteps(() => {
      expect(table.internPreview(parsePreview(800, diagramMedia(4)))).toBe(0);
    });
    expect(steps).toBeLessThan(MAX_ITERATOR_STEPS_PER_PREVIEW);

    const source = table.snapshot().at(0);
    expect(source?.format).toBe("png");
    expect((source?.pixelWidth ?? 0) * (source?.pixelHeight ?? 0)).toBeLessThanOrEqual(
      MAX_PREVIEW_PIXELS,
    );
  });

  /**
   * The picture is not meant to change. A fixed descriptor pins the exact PNG,
   * so a later edit to the rasteriser that alters what a diagram looks like has
   * to say so here.
   */
  test("a fixed descriptor rasterises to fixed bytes", () => {
    const table = new ImageTable();
    table.internPreview({
      kind: "diagram",
      extent: { width: 400, height: 200 },
      shapes: [{ x: 40, y: 20, width: 160, height: 80, color: "70AD47" }],
      pixelWidth: 400,
      pixelHeight: 200,
    });
    const bytes = table.snapshot().at(0)?.bytes as Uint8Array;
    expect(bytes.byteLength).toBe(320_288);
    expect(Bun.SHA256.hash(bytes, "hex")).toBe(
      "3f9ec77a688d41bb75e467e64e03264ed2fe185aa93391b5ce05ddb0965cf370",
    );
  });
});
