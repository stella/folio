/**
 * A preview the budget does not recognize is a preview with no bound, which is
 * how the SmartArt raster came to retain tens of megabytes from a package of
 * twenty kilobytes. Recognition depends on the producer and the budget
 * agreeing about three strings, so these tests drive a real producer and check
 * that what it emits is what the budget charges for.
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
  if (!image?.src) {
    throw new Error("diagram fixture produced no preview");
  }
  return image;
};

describe("package preview budget", () => {
  test("charges the SmartArt preview a producer actually emits", () => {
    const image = smartArtPreview();
    expect(image.mimeType).toBe(PREVIEW_KINDS.smartArt.mimeType);
    expect(image.filename).toBe(PREVIEW_KINDS.smartArt.filename);
    expect(image.src).toStartWith(PREVIEW_KINDS.smartArt.srcPrefix);

    // One character short of this preview: it is the first, so it fits exactly
    // at its own length and is dropped below it.
    const kept = { image: smartArtPreview() };
    enforcePackagePreviewBudget(kept, { smartArt: image.src?.length });
    expect(kept.image.src).toBeDefined();

    const dropped = { image: smartArtPreview() };
    enforcePackagePreviewBudget(dropped, { smartArt: (image.src?.length ?? 0) - 1 });
    expect(dropped.image.src).toBeUndefined();
  });

  test("a package's retained preview text is bounded however many diagrams it has", () => {
    const previews = Array.from({ length: 8 }, () => smartArtPreview());
    const single = previews[0]?.src?.length ?? 0;
    expect(single).toBeGreaterThan(0);

    // Room for three, offered eight.
    enforcePackagePreviewBudget({ previews }, { smartArt: single * 3 });
    const retained = previews.reduce((total, image) => total + (image.src?.length ?? 0), 0);
    expect(retained).toBe(single * 3);
    expect(previews.filter((image) => image.src !== undefined)).toHaveLength(3);
  });

  test("keeps the drawing when it drops the preview", () => {
    const image = smartArtPreview();
    const drawing = { image, rawXml: "<w:drawing/>" };
    enforcePackagePreviewBudget(drawing, { smartArt: 0 });
    expect(image.src).toBeUndefined();
    expect(image.size).toEqual({ width: 400, height: 200 });
    expect(drawing.rawXml).toBe("<w:drawing/>");
  });

  test("leaves a relationship-backed image alone", () => {
    const real = {
      type: "image",
      rId: "rId7",
      src: `${PREVIEW_KINDS.smartArt.srcPrefix}AAAA`,
      mimeType: PREVIEW_KINDS.smartArt.mimeType,
      filename: PREVIEW_KINDS.smartArt.filename,
    };
    enforcePackagePreviewBudget({ real }, { smartArt: 0 });
    expect(real.src).toBe(`${PREVIEW_KINDS.smartArt.srcPrefix}AAAA`);
  });

  test("one kind's allowance does not spend another's", () => {
    const vml = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    };
    const diagram = smartArtPreview();
    enforcePackagePreviewBudget({ vml, diagram }, { smartArt: 0 });
    expect(diagram.src).toBeUndefined();
    expect(vml.src).toBe(`${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`);
  });
});
