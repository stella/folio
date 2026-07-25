import { describe, expect, test } from "bun:test";

import {
  parseAnchorPosition,
  parseColorElement,
  parseFill,
  parseOutline,
  parsePositionH,
  parsePositionV,
} from "./drawingUtils";
import { serializeRun } from "./serializer/runSerializer";
import type { XmlElement } from "./xmlParser";
import { parseXmlDocument } from "./xmlParser";

function el(name: string, attributes: Record<string, string> = {}): XmlElement {
  return { name, type: "element", attributes };
}

function wrap(child: XmlElement): XmlElement {
  return { name: "wrapper", type: "element", elements: [child] };
}

describe("drawingUtils position parsing", () => {
  test("normalizes omitted anchor positions to serializer defaults", () => {
    const anchor = parseXmlDocument(`
      <wp:anchor
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      />
    `);
    expect(anchor).not.toBeNull();
    if (!anchor) {
      return;
    }

    expect(parseAnchorPosition(anchor)).toEqual({
      horizontal: { relativeTo: "column", posOffset: 0 },
      vertical: { relativeTo: "paragraph", posOffset: 0 },
    });
  });

  test("normalizes position axes that omit both alignment and offset", () => {
    const positionH = parseXmlDocument(`
      <wp:positionH
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        relativeFrom="character"
      />
    `);
    const positionV = parseXmlDocument(`
      <wp:positionV
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        relativeFrom="page"
      />
    `);

    expect(parsePositionH(positionH)).toEqual({ relativeTo: "character", posOffset: 0 });
    expect(parsePositionV(positionV)).toEqual({ relativeTo: "page", posOffset: 0 });
  });

  test("keeps valid alignments instead of materializing an offset", () => {
    const positionH = parseXmlDocument(`
      <wp:positionH
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        relativeFrom="margin"
      >
        <wp:align>center</wp:align>
      </wp:positionH>
    `);
    const positionV = parseXmlDocument(`
      <wp:positionV
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        relativeFrom="page"
      >
        <wp:align>bottom</wp:align>
      </wp:positionV>
    `);

    expect(parsePositionH(positionH)).toEqual({
      relativeTo: "margin",
      alignment: "center",
    });
    expect(parsePositionV(positionV)).toEqual({
      relativeTo: "page",
      alignment: "bottom",
    });
  });
});

describe("drawingUtils.parseColorElement", () => {
  test("accepts a well-formed srgbClr value and uppercases it", () => {
    const result = parseColorElement(wrap(el("a:srgbClr", { val: "ff8800" })));
    expect(result).toEqual({ rgb: "FF8800" });
  });

  test("rejects srgbClr values that are not exactly six hex digits", () => {
    for (const val of [
      "FF",
      "FFFFFFF",
      "GGGGGG",
      "FFFFF#",
      "FF8800;color:red",
      `FF8800"/><script>alert(1)</script>`,
      "url(javascript:alert(1))",
      "",
    ]) {
      const result = parseColorElement(wrap(el("a:srgbClr", { val })));
      expect(result).toBeUndefined();
    }
  });

  test("falls back to black for sysClr without a hex lastClr", () => {
    const result = parseColorElement(
      wrap(el("a:sysClr", { val: "windowText", lastClr: `not-hex"/>` })),
    );
    expect(result).toEqual({ rgb: "000000" });
  });

  test("uses the validated lastClr from sysClr when present", () => {
    const result = parseColorElement(
      wrap(el("a:sysClr", { val: "windowText", lastClr: "1f497d" })),
    );
    expect(result).toEqual({ rgb: "1F497D" });
  });
});

describe("drawingUtils.parseFill", () => {
  test("projects gradient details while preserving authored DrawingML", () => {
    const shapeProperties = parseXmlDocument(`
      <wps:spPr
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
      >
        <a:gradFill rotWithShape="0">
          <a:gsLst>
            <a:gs pos="0"><a:srgbClr val="FFFFFF"><a:lumMod val="95000"/></a:srgbClr></a:gs>
            <a:gs pos="100000"><a:srgbClr val="5B9BD5"/></a:gs>
          </a:gsLst>
          <a:lin ang="5400000" scaled="0"/>
          <a:tileRect/>
        </a:gradFill>
      </wps:spPr>
    `);
    expect(shapeProperties).not.toBeNull();
    if (!shapeProperties) {
      return;
    }

    const fill = parseFill(shapeProperties);
    expect(fill).toMatchObject({
      type: "gradient",
      gradient: {
        type: "linear",
        angle: 90,
        stops: [
          { position: 0, color: { rgb: "FFFFFF" } },
          { position: 100_000, color: { rgb: "5B9BD5" } },
        ],
      },
    });
    expect(fill?.rawXml).toContain('<a:gradFill rotWithShape="0">');
    expect(fill?.rawXml).toContain('<a:lumMod val="95000"/>');
    expect(fill?.rawXml).toContain('<a:lin ang="5400000" scaled="0"/>');
    expect(fill?.rawXml).toContain("<a:tileRect/>");
    if (!fill?.rawXml) {
      return;
    }

    const serialized = serializeRun({
      type: "run",
      content: [
        {
          type: "shape",
          shape: {
            type: "shape",
            shapeType: "textBox",
            size: { width: 914_400, height: 457_200 },
            fill,
            textBody: { content: [] },
          },
        },
      ],
    });
    expect(serialized).toContain(fill.rawXml);
  });
});

describe("drawingUtils.parseOutline", () => {
  test("projects outline details while preserving authored DrawingML", () => {
    const shapeProperties = parseXmlDocument(`
      <wps:spPr
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
      >
        <a:ln w="12700" cap="rnd" cmpd="dbl" algn="in">
          <a:solidFill>
            <a:schemeClr val="accent1">
              <a:shade val="50000"/>
              <a:satMod val="80000"/>
            </a:schemeClr>
          </a:solidFill>
          <a:prstDash val="dash"/>
          <a:miter lim="800000"/>
          <a:headEnd type="triangle" w="lg" len="sm"/>
          <a:tailEnd type="oval"/>
        </a:ln>
      </wps:spPr>
    `);
    expect(shapeProperties).not.toBeNull();
    if (!shapeProperties) {
      return;
    }

    const outline = parseOutline(shapeProperties);
    expect(outline).toMatchObject({
      width: 12_700,
      cap: "round",
      color: { themeColor: "accent1", themeShade: "80" },
      style: "dash",
      join: "miter",
      headEnd: { type: "triangle", width: "lg", length: "sm" },
      tailEnd: { type: "oval" },
    });
    expect(outline?.rawXml).toContain('<a:ln w="12700" cap="rnd" cmpd="dbl" algn="in">');
    expect(outline?.rawXml).toContain('<a:satMod val="80000"/>');
    expect(outline?.rawXml).toContain('<a:miter lim="800000"/>');
    if (!outline?.rawXml) {
      return;
    }

    const serialized = serializeRun({
      type: "run",
      content: [
        {
          type: "shape",
          shape: {
            type: "shape",
            shapeType: "textBox",
            size: { width: 914_400, height: 457_200 },
            outline,
            textBody: { content: [] },
          },
        },
      ],
    });
    expect(serialized).toContain(outline.rawXml);

    const { rawXml: authoredXml, ...structuredOutline } = outline;
    expect(authoredXml).toBe(outline.rawXml);
    const edited = serializeRun({
      type: "run",
      content: [
        {
          type: "shape",
          shape: {
            type: "shape",
            shapeType: "textBox",
            size: { width: 914_400, height: 457_200 },
            outline: { ...structuredOutline, width: 25_400 },
            textBody: { content: [] },
          },
        },
      ],
    });
    expect(edited).toContain('<a:ln w="25400" cap="rnd">');
    expect(edited).not.toContain('cmpd="dbl"');
  });
});
