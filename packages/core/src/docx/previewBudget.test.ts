/**
 * A preview the budget does not recognize is a preview with no bound, which is
 * how the SmartArt raster came to retain tens of megabytes from a package of
 * twenty kilobytes. Recognition depends on the producer and the budget
 * agreeing about three strings, so these tests drive a real producer and check
 * that what it emits is what the budget charges for.
 *
 * The SmartArt producer no longer emits a raster, so the character budget has
 * nothing of its to charge and the bound moved to `ImageTable`, where the
 * rasters now are; that half is tested beside it in
 * `display-list/build/previewRasterBudget.test.ts`. What stays here is the
 * producer agreement itself, and the VML preview, which is still a `src`.
 */

import { describe, expect, test } from "bun:test";

import type { MediaFile } from "../types/document";
import { parseDiagramPreview } from "./diagramPreview";
import { PREVIEW_KINDS, enforcePackagePreviewBudget } from "./previewBudget";
import { parseRelationships } from "./relsParser";
import { parseXmlDocument } from "./xmlParser";

const RELATIONSHIPS = parseRelationships(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/></Relationships>`,
);

const MEDIA: Map<string, MediaFile> = new Map([
  [
    "word/diagrams/drawing1.xml",
    {
      path: "word/diagrams/drawing1.xml",
      filename: "drawing1.xml",
      mimeType: "application/xml",
      data: new TextEncoder().encode(
        `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="400" cy="200"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`,
      ).buffer,
    },
  ],
]);

const diagramDrawing = () => {
  const drawing = parseXmlDocument(
    `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="400" cy="200"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic></wp:inline></w:drawing>`,
  );
  if (!drawing) {
    throw new Error("diagram fixture did not parse");
  }
  return drawing;
};

const smartArtPreview = () => {
  const image = parseDiagramPreview(diagramDrawing(), RELATIONSHIPS, MEDIA);
  if (!image?.preview) {
    throw new Error("diagram fixture produced no preview");
  }
  return image;
};

describe("package preview budget", () => {
  test("the SmartArt producer emits a description, not a raster", () => {
    const image = smartArtPreview();
    expect(image.mimeType).toBe(PREVIEW_KINDS.smartArt.mimeType);
    expect(image.filename).toBe(PREVIEW_KINDS.smartArt.filename);
    expect(image.src).toBeUndefined();
    expect(image.preview?.kind).toBe("diagram");
  });

  /**
   * The budget may not drop a descriptor. Dropping a `src` gave up a raster
   * that could be rebuilt from the package; dropping a descriptor gives up the
   * only record of what the drawing looks like, and rebuilding it means
   * re-reading parts the parse has finished with.
   */
  test("the character budget leaves a descriptor alone at any allowance", () => {
    const image = smartArtPreview();
    enforcePackagePreviewBudget({ image }, { smartArt: 0 });
    expect(image.preview?.kind).toBe("diagram");
    expect(image.preview?.extent).toEqual({ width: 400, height: 200 });
    expect(image.size).toEqual({ width: 400, height: 200 });
  });

  test("keeps the drawing when it drops a preview", () => {
    const image = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    };
    const drawing = { image, rawXml: "<w:pict/>" };
    enforcePackagePreviewBudget(drawing, { vmlShape: 0 });
    expect(image.src).toBeUndefined();
    expect(drawing.rawXml).toBe("<w:pict/>");
  });

  test("a package's retained VML preview text is bounded however many shapes it has", () => {
    const preview = () => ({
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    });
    const previews = Array.from({ length: 8 }, preview);
    const single = previews[0]?.src.length ?? 0;
    expect(single).toBeGreaterThan(0);

    // Room for three, offered eight.
    enforcePackagePreviewBudget({ previews }, { vmlShape: single * 3 });
    expect(previews.filter((image) => image.src !== undefined)).toHaveLength(3);
  });

  test("leaves a relationship-backed image alone", () => {
    const real = {
      type: "image",
      rId: "rId7",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    };
    enforcePackagePreviewBudget({ real }, { vmlShape: 0 });
    expect(real.src).toBe(`${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`);
  });

  test("one kind's allowance does not spend another's", () => {
    const vml = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    };
    const smartArt = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.smartArt.srcPrefix}AAAA`,
      mimeType: PREVIEW_KINDS.smartArt.mimeType,
      filename: PREVIEW_KINDS.smartArt.filename,
    };
    enforcePackagePreviewBudget({ vml, smartArt }, { smartArt: 0 });
    expect(smartArt.src).toBeUndefined();
    expect(vml.src).toBe(`${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`);
  });
});
