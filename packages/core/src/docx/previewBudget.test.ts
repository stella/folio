/**
 * A preview the budget does not recognize is a preview with no bound, which is
 * how the SmartArt raster came to retain tens of megabytes from a package of
 * twenty kilobytes. Recognition depends on the producer and the budget
 * agreeing about three strings, so these tests drive a real producer and check
 * that what it emits is what the budget charges for.
 *
 * The SmartArt producer no longer emits a raster, so the character budget has
 * nothing of its to charge: the drawing is drawn from its descriptor, bounded
 * by the shape cap the parse applies. What stays here is the producer
 * agreement itself, and the two SVG previews, which are still a `src`.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document, DrawingContent, MediaFile, Run } from "../types/document";
import { parseDiagramPreview } from "./diagramPreview";
import { parseDocx } from "./parser";
import { PREVIEW_KINDS, enforcePackagePreviewBudget } from "./previewBudget";
import { RELATIONSHIP_TYPES, parseRelationships } from "./relsParser";
import { repackDocx, validateDocx } from "./rezip";
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

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

/** The group a synthetic package carries, and the drawing folio renders it as. */
const GROUP_DRAWING = `<w:drawing><wp:inline><wp:extent cx="2000000" cy="1000000"/><a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"><wpg:wgp><wps:wsp><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="500000"/></a:xfrm><a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill></wps:spPr></wps:wsp></wpg:wgp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const groupDocx = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `${XML}
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `${XML}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="${RELATIONSHIP_TYPES.officeDocument}" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/_rels/document.xml.rels",
    `${XML}
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `${XML}
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"><w:body><w:p><w:r>${GROUP_DRAWING}</w:r></w:p><w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>`,
  );
  zip.file(
    "word/styles.xml",
    `${XML}
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"></w:styles>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const firstDrawing = (document: Document): DrawingContent | undefined => {
  const block = document.package.document.content.at(0);
  if (block?.type !== "paragraph") {
    return undefined;
  }
  const run = block.content.find((item): item is Run => item.type === "run");
  return run?.content.find((item): item is DrawingContent => item.type === "drawing");
};

describe("package preview budget", () => {
  test("the SmartArt producer emits a description, not a raster", () => {
    const image = smartArtPreview();
    expect(image.src).toBeUndefined();
    expect(image.mimeType).toBeUndefined();
    expect(image.filename).toBeUndefined();
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
    enforcePackagePreviewBudget({ image }, { vmlShape: 0 });
    expect(image.preview?.kind).toBe("diagram");
    expect(image.preview?.extent).toEqual({ width: 400, height: 200 });
    expect(image.size).toEqual({ width: 400, height: 200 });
  });

  /**
   * The table's two branches, read from the budget's side. A descriptor-backed
   * kind declares no `src` strings, so nothing the model carries can be
   * matched to it: the raster's own identity, still attached to a diagram
   * image, is charged to no kind and dropped by none.
   */
  test("a descriptor-backed kind cannot be matched by a src", () => {
    const image = Object.assign(smartArtPreview(), {
      src: "data:image/png;base64,AAAA",
      mimeType: "image/png",
      filename: "smartart-preview.png",
    });
    enforcePackagePreviewBudget({ image }, { vmlShape: 0 });
    expect(image.src).toBe("data:image/png;base64,AAAA");
    expect(image.preview?.kind).toBe("diagram");
    // @ts-expect-error a descriptor-backed kind has no character allowance.
    enforcePackagePreviewBudget({ image }, { smartArt: 0 });
  });

  /**
   * And a source-backed kind has no descriptor path out of the budget: the
   * preview is the `src`, so a descriptor sitting beside it shields nothing.
   */
  test("a source-backed kind is charged whatever else the image carries", () => {
    const image = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
      preview: { kind: "diagram", extent: { width: 400, height: 200 }, shapes: [] },
    };
    enforcePackagePreviewBudget({ image }, { vmlShape: 0 });
    expect(image.src).toBeUndefined();
    expect(image.preview.kind).toBe("diagram");
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

  /**
   * Two kinds share a data-URL prefix and a mime type, so only the filename
   * tells a group preview from a VML one. A budget that matched on the prefix
   * alone would spend one kind's allowance on the other.
   */
  test("one kind's allowance does not spend another's", () => {
    const vml = {
      type: "image",
      rId: "",
      src: `${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.vmlShape.mimeType,
      filename: PREVIEW_KINDS.vmlShape.filename,
    };
    const group = {
      type: "image",
      src: `${PREVIEW_KINDS.wpGroup.srcPrefix}%3Csvg%3E`,
      mimeType: PREVIEW_KINDS.wpGroup.mimeType,
      filename: PREVIEW_KINDS.wpGroup.filename,
    };
    enforcePackagePreviewBudget({ vml, group }, { wpGroup: 0 });
    expect(group.src).toBeUndefined();
    expect(vml.src).toBe(`${PREVIEW_KINDS.vmlShape.srcPrefix}%3Csvg%3E`);
  });

  /**
   * What being over budget costs, end to end: the group preview goes and
   * nothing else does. The drawing keeps the space it reserved on the page,
   * the package saves the group it was authored with, and opening the saved
   * package renders the preview again, because dropping one was an economy in
   * the model and never a loss from the file.
   */
  test("drops a group preview past its allowance and still saves the group", async () => {
    const original = await groupDocx();
    const document = await parseDocx(original, { preloadFonts: false });
    const drawing = firstDrawing(document);
    const preview = drawing?.image.src;
    expect(preview).toStartWith(PREVIEW_KINDS.wpGroup.srcPrefix);
    expect(drawing?.rawXmlMode).toBe("previewOnly");

    enforcePackagePreviewBudget(document.package, { wpGroup: 0 });
    expect(drawing?.image.src).toBeUndefined();
    expect(drawing?.image.size).toEqual({ width: 2_000_000, height: 1_000_000 });

    const saved = await repackDocx(document, { updateModifiedDate: false });
    expect((await validateDocx(saved)).valid).toBe(true);
    const savedXml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("text");
    expect(savedXml).toContain("<wpg:wgp>");
    expect(savedXml).toContain('<a:srgbClr val="DBEDF3"/>');
    expect(savedXml).not.toContain("svg");

    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(firstDrawing(reopened)?.image.src).toBe(preview);
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
});
