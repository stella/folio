import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";
import type { MediaFile } from "../types/document";
import { MAX_PREVIEW_SHAPES } from "./diagramPreview";
import { parseChartPreview } from "./chartPreview";
import { parseDocx } from "./parser";
import { parseRelationships } from "./relsParser";
import { createEmptyDocx, repackDocx } from "./rezip";
import { parseRun } from "./runParser";
import { parseXmlDocument } from "./xmlParser";

const NS = "http://schemas.openxmlformats.org/drawingml/2006/chart";
const drawingXml = (namespace = NS) =>
  `<w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="${namespace}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><wp:inline><wp:extent cx="1000" cy="1000"/><a:graphic><a:graphicData uri="${namespace}"><c:chart r:id="rChart"/></a:graphicData></a:graphic></wp:inline></w:drawing>`;
const relationships = parseRelationships(
  `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>`,
);
const xmlFile = (path: string, xml: string): [string, MediaFile] => [
  path,
  {
    path,
    filename: path.split("/").at(-1),
    mimeType: "application/xml",
    data: new TextEncoder().encode(xml).buffer,
  },
];
const cache = (name: "cat" | "val", values: readonly (string | number)[]) =>
  `<c:${name}><c:${name === "cat" ? "strRef" : "numRef"}><c:${name === "cat" ? "strCache" : "numCache"}>${values.map((value, index) => `<c:pt idx="${index}"><c:v>${value}</c:v></c:pt>`).join("")}</c:${name === "cat" ? "strCache" : "numCache"}></c:${name === "cat" ? "strRef" : "numRef"}></c:${name}>`;
const series = (values: readonly (string | number)[], color = "4472C4") =>
  `<c:ser><c:spPr><a:solidFill xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:srgbClr val="${color}"/></a:solidFill></c:spPr>${cache(
    "cat",
    values.map((_, index) => `C${index}`),
  )}${cache("val", values)}</c:ser>`;
const chart = (kind: string, body: string, namespace = NS) =>
  `<c:chartSpace xmlns:c="${namespace}"><c:chart><c:title><c:tx><c:rich><a:p xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:r><a:t>Sample</a:t></a:r></a:p></c:rich></c:tx></c:title><c:plotArea><c:${kind}>${body}</c:${kind}></c:plotArea></c:chart></c:chartSpace>`;
const preview = (xml: string, namespace = NS, extra: [string, MediaFile][] = []) => {
  const drawing = parseXmlDocument(drawingXml(namespace));
  if (!drawing) throw new Error("drawing did not parse");
  return parseChartPreview(
    drawing,
    relationships,
    new Map([xmlFile("word/charts/chart1.xml", xml), ...extra]),
  );
};

describe("chart previews", () => {
  test("column geometry uses a zero baseline and puts negative values below it", () => {
    const image = preview(
      chart("barChart", `<c:barDir val="col"/><c:grouping val="clustered"/>${series([-5, 10])}`),
    );
    const shapes = image?.preview?.shapes ?? [];
    expect(shapes).toHaveLength(2);
    expect(shapes[0]).toEqual({
      x: 123,
      y: 640.0000000000001,
      width: 344,
      height: 280,
      color: "4472C4",
    });
    expect(shapes[1]).toEqual({
      x: 553,
      y: 80,
      width: 344,
      height: 560.0000000000001,
      color: "4472C4",
    });
  });

  test.each([
    [
      "stacked bar",
      "barChart",
      `<c:barDir val="bar"/><c:grouping val="stacked"/>${series([2, 3])}${series([4, 1], "ED7D31")}`,
    ],
    ["line", "lineChart", series([1, 3, 2])],
    ["pie", "pieChart", series([1, 2, 3])],
    ["doughnut", "doughnutChart", series([1, 2, 3])],
  ])("renders a bounded %s chart", (_name, kind, body) => {
    const shapes = preview(chart(kind, body))?.preview?.shapes ?? [];
    expect(shapes.length).toBeGreaterThan(0);
    expect(shapes.length).toBeLessThanOrEqual(MAX_PREVIEW_SHAPES);
    expect(
      shapes.every(
        (shape) =>
          shape.x >= 0 &&
          shape.y >= 0 &&
          shape.x + shape.width <= 1000 &&
          shape.y + shape.height <= 1000,
      ),
    ).toBe(true);
  });

  test("uses a DrawingML picture from an AlternateContent fallback", () => {
    const run = parseXmlDocument(
      `<w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="${NS}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><mc:AlternateContent><mc:Choice Requires="c">${drawingXml()}</mc:Choice><mc:Fallback><w:drawing><wp:inline><wp:extent cx="10" cy="20"/><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:blipFill><a:blip r:embed="rFallback"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Fallback></mc:AlternateContent></w:r>`,
    );
    if (!run) throw new Error("run did not parse");
    const rels = parseRelationships(
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/><Relationship Id="rFallback" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/fallback.png"/></Relationships>`,
    );
    const fallback: MediaFile = {
      path: "word/media/fallback.png",
      filename: "fallback.png",
      mimeType: "image/png",
      data: new ArrayBuffer(1),
      dataUrl: "data:image/png;base64,AA==",
    };
    const parsed = parseRun(
      run,
      null,
      null,
      rels,
      new Map([
        xmlFile("word/charts/chart1.xml", chart("barChart", series([1]))),
        [fallback.path, fallback],
      ]),
    );
    const content = parsed.content.at(0);
    expect(content?.type).toBe("drawing");
    if (content?.type !== "drawing") throw new Error("fallback was not a drawing");
    expect(content.image.src).toBe(fallback.dataUrl);
    expect(content.image.size).toEqual({ width: 1000, height: 1000 });
    expect(content.rawXml).toContain("mc:AlternateContent");
  });

  test("uses a chart-part related image before generated geometry", () => {
    const rels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart.png"/></Relationships>`;
    const image: MediaFile = {
      path: "word/media/chart.png",
      filename: "chart.png",
      mimeType: "image/png",
      data: new ArrayBuffer(1),
      dataUrl: "data:image/png;base64,AA==",
    };
    const result = preview(chart("barChart", series([1])), NS, [
      xmlFile("word/charts/_rels/chart1.xml.rels", rels),
      [image.path, image],
    ]);
    expect(result?.src).toBe(image.dataUrl);
    expect(result?.preview).toBeUndefined();
    expect(result?.size).toEqual({ width: 1000, height: 1000 });
  });

  test("fails closed without caches and for non-finite values", () => {
    expect(preview(chart("barChart", `<c:ser/>`))).toBeNull();
    expect(preview(chart("barChart", series(["NaN"])))).toBeNull();
    expect(preview(chart("barChart", series(["Infinity"])))).toBeNull();
  });

  test("caps oversized cached charts", () => {
    const manySeries = Array.from({ length: 40 }, (_, index) =>
      series(Array.from({ length: 100 }, () => index + 1)),
    ).join("");
    expect(
      preview(chart("barChart", `<c:barDir val="col"/>${manySeries}`))?.preview?.shapes.length,
    ).toBe(MAX_PREVIEW_SHAPES);
  });

  test("recognises the strict chart namespace", () => {
    const strict = "http://purl.oclc.org/ooxml/drawingml/chart";
    expect(
      preview(chart("lineChart", series([1, 2]), strict), strict)?.preview?.shapes.length,
    ).toBe(2);
  });

  test("finite cached series always stay inside the extent and never throw", () => {
    fc.assert(
      fc.property(
        fc.array(fc.double({ noNaN: true, noDefaultInfinity: true, min: -1e6, max: 1e6 }), {
          minLength: 1,
          maxLength: 32,
        }),
        (values) => {
          const shapes =
            preview(chart("barChart", `<c:barDir val="col"/>${series(values)}`))?.preview?.shapes ??
            [];
          return shapes.every(
            (shape) =>
              shape.x >= 0 &&
              shape.y >= 0 &&
              shape.x + shape.width <= 1000.000001 &&
              shape.y + shape.height <= 1000.000001,
          );
        },
      ),
      { numRuns: 100 },
    );
  });
  test("round-trips chart XML and chart relationships unchanged", async () => {
    const chartXml = chart(
      "barChart",
      `<c:barDir val="col"/><c:grouping val="clustered"/>${series([1, 2])}`,
    );
    const chartRels = `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rImage" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/chart.png"/></Relationships>`;
    const zip = await JSZip.loadAsync(await createEmptyDocx());
    zip.file(
      "word/document.xml",
      `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:c="${NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body><w:p><w:r>${drawingXml()}</w:r></w:p><w:sectPr/></w:body></w:document>`,
    );
    zip.file(
      "word/_rels/document.xml.rels",
      `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rChart" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="charts/chart1.xml"/></Relationships>`,
    );
    zip.file("word/charts/chart1.xml", chartXml);
    zip.file("word/charts/_rels/chart1.xml.rels", chartRels);
    zip.file("word/media/chart.png", new Uint8Array([0]));
    const source = await zip.generateAsync({ type: "arraybuffer" });
    const saved = await repackDocx(await parseDocx(source, { preloadFonts: false }), {
      updateModifiedDate: false,
    });
    const result = await JSZip.loadAsync(saved);
    expect(await result.file("word/charts/chart1.xml")?.async("text")).toBe(chartXml);
    expect(await result.file("word/charts/_rels/chart1.xml.rels")?.async("text")).toBe(chartRels);
  });
});
