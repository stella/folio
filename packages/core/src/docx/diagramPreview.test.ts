import { describe, expect, test } from "bun:test";
import type { MediaFile, RelationshipMap } from "../types/document";
import { parseDiagramPreview } from "./diagramPreview";
import { ImageTable } from "../display-list/build/imagePrimitives";
import { parseRelationships } from "./relsParser";
import { parseXmlDocument } from "./xmlParser";

const drawing = parseXmlDocument(
  `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="rIdData"/></a:graphicData></a:graphic></wp:inline></w:drawing>`,
);

/** The extension in a diagram's data part that names its cached drawing. */
const dataModelExt = (relId: string): string =>
  `<dgm:extLst><a:ext xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" uri="http://schemas.microsoft.com/office/drawing/2008/diagram"><dsp:dataModelExt xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" relId="${relId}"/></a:ext></dgm:extLst>`;

const xmlPart = (path: string, xml: string): [string, MediaFile] => [
  path,
  {
    path,
    filename: path.split("/").at(-1) ?? path,
    mimeType: "application/xml",
    data: new TextEncoder().encode(xml).buffer,
  },
];

describe("SmartArt preview", () => {
  test("builds a bounded PNG from cached diagram shapes while retaining dimensions", () => {
    if (!drawing) throw new Error("fixture did not parse");
    const rels: RelationshipMap = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/></Relationships>`,
    );
    const xml = `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:pt modelId="1"><dgm:txBody><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Organisation Chart</a:t></a:r></a:p></dgm:txBody></dgm:pt>${dataModelExt("rIdDrawing")}</dgm:dataModel>`;
    const data = new TextEncoder().encode(xml).buffer;
    const drawingData = new TextEncoder().encode(
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="10000" y="10000"/><a:ext cx="400000" cy="200000"/></a:xfrm><a:solidFill><a:srgbClr val="70AD47"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`,
    ).buffer;
    const media = new Map<string, MediaFile>([
      [
        "word/diagrams/data1.xml",
        {
          path: "word/diagrams/data1.xml",
          filename: "data1.xml",
          mimeType: "application/xml",
          data,
        },
      ],
      [
        "word/diagrams/drawing1.xml",
        {
          path: "word/diagrams/drawing1.xml",
          filename: "drawing1.xml",
          mimeType: "application/xml",
          data: drawingData,
        },
      ],
    ]);
    const image = parseDiagramPreview(drawing, rels, media);
    expect(image?.mimeType).toBe("image/png");
    expect(image?.size).toEqual({ width: 914400, height: 457200 });
    expect(image?.src).toStartWith("data:image/png;base64,");
    expect(new ImageTable().intern(image?.src ?? "")).toBeDefined();
  });

  test("falls back to a bounded background when cached drawing data is unavailable", () => {
    if (!drawing) throw new Error("fixture did not parse");
    const rels = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/></Relationships>`,
    );
    const media = new Map<string, MediaFile>();
    const image = parseDiagramPreview(drawing, rels, media);
    expect(image?.mimeType).toBe("image/png");
    expect(new ImageTable().intern(image?.src ?? "")).toBeDefined();
  });

  /**
   * Two diagrams in one package each have a data part and a cached drawing.
   * Choosing the drawing by scanning for a relationship type finds two, and the
   * scan refused on more than one, so neither diagram got its shapes.
   */
  test("resolves each of two diagrams to its own cached drawing", () => {
    const diagramDrawing = (colour: string, width: number): string =>
      `<dsp:drawing xmlns:dsp="http://schemas.microsoft.com/office/drawing/2008/diagram" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><dsp:spTree><dsp:sp><dsp:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${String(width)}" cy="400000"/></a:xfrm><a:solidFill><a:srgbClr val="${colour}"/></a:solidFill></dsp:spPr></dsp:sp></dsp:spTree></dsp:drawing>`;
    const rels = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdData1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data1.xml"/><Relationship Id="rIdDrawing1" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing1.xml"/><Relationship Id="rIdData2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" Target="diagrams/data2.xml"/><Relationship Id="rIdDrawing2" Type="http://schemas.microsoft.com/office/2007/relationships/diagramDrawing" Target="diagrams/drawing2.xml"/></Relationships>`,
    );
    const media = new Map<string, MediaFile>([
      xmlPart(
        "word/diagrams/data1.xml",
        `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">${dataModelExt("rIdDrawing1")}</dgm:dataModel>`,
      ),
      xmlPart(
        "word/diagrams/data2.xml",
        `<dgm:dataModel xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram">${dataModelExt("rIdDrawing2")}</dgm:dataModel>`,
      ),
      xmlPart("word/diagrams/drawing1.xml", diagramDrawing("70AD47", 400_000)),
      xmlPart("word/diagrams/drawing2.xml", diagramDrawing("C00000", 800_000)),
    ]);
    const diagramDrawingFor = (dataRelationshipId: string) => {
      const element = parseXmlDocument(
        `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:dgm="http://schemas.openxmlformats.org/drawingml/2006/diagram" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="914400" cy="457200"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/diagram"><dgm:relIds r:dm="${dataRelationshipId}"/></a:graphicData></a:graphic></wp:inline></w:drawing>`,
      );
      if (!element) throw new Error("fixture did not parse");
      return parseDiagramPreview(element, rels, media);
    };

    const first = diagramDrawingFor("rIdData1");
    const second = diagramDrawingFor("rIdData2");

    // Each preview paints its own diagram's shape, so the two differ.
    expect(first?.src).toStartWith("data:image/png;base64,");
    expect(second?.src).toStartWith("data:image/png;base64,");
    expect(first?.src).not.toBe(second?.src);
  });
});
