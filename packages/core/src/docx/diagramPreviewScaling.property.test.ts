/**
 * The SmartArt preview is a raster, so every defect in it is a defect per byte.
 *
 * A millisecond budget cannot say that on shared hardware, and an allocation
 * count cannot say it at all: walking a megapixel buffer with
 * `for (const byte of buffer)` allocates one iterator and then pays the
 * protocol five million times. What does say it is the number of iterator
 * steps the render takes, which is a property of the code rather than of the
 * machine. These tests scale the raster and assert that the step count does
 * not follow it, and that the bytes a package retains stay proportional to the
 * number of diagrams rather than to anything larger.
 */

import { describe, expect, test } from "bun:test";

import type { MediaFile } from "../types/document";
import { parseDiagramPreview } from "./diagramPreview";
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
 * Count the iterator steps a render takes over byte buffers.
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

describe("SmartArt preview scaling", () => {
  test("iterator steps do not follow the raster's area", () => {
    const media = diagramMedia(4);
    const counts = [200, 400, 800, 1600].map((extent) =>
      countByteIteratorSteps(() => {
        expect(
          parseDiagramPreview(diagramDrawing(extent), DIAGRAM_RELATIONSHIPS, media),
        ).not.toBeNull();
      }),
    );

    // Sixty-four times the pixels must not buy a single extra step.
    for (const steps of counts) {
      expect(steps).toBeLessThan(MAX_ITERATOR_STEPS_PER_PREVIEW);
    }
    expect(new Set(counts).size).toBe(1);
  });

  test("a package's preview bytes stay proportional to its diagram count", () => {
    const media = diagramMedia(4);
    const drawing = diagramDrawing(800);
    const single = parseDiagramPreview(drawing, DIAGRAM_RELATIONSHIPS, media)?.src ?? "";
    expect(single.length).toBeGreaterThan(0);

    for (const diagrams of [1, 2, 4, 8]) {
      const total = Array.from({ length: diagrams }, () =>
        parseDiagramPreview(drawing, DIAGRAM_RELATIONSHIPS, media),
      ).reduce((bytes, image) => bytes + (image?.src?.length ?? 0), 0);
      expect(total).toBe(single.length * diagrams);
    }
  });

  test("shape count drives the render, not the raster", () => {
    const drawing = diagramDrawing(800);
    const counts = [1, 2, 4, 8].map((shapes) =>
      countByteIteratorSteps(() => {
        parseDiagramPreview(drawing, DIAGRAM_RELATIONSHIPS, diagramMedia(shapes));
      }),
    );
    for (const steps of counts) {
      expect(steps).toBeLessThan(MAX_ITERATOR_STEPS_PER_PREVIEW);
    }
  });
});
