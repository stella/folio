import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { Window } from "happy-dom";
import { Schema } from "prosemirror-model";

import { parseShape } from "../../../docx/shapeParser";
import { parseXmlDocument } from "../../../docx/xmlParser";
import { ShapeExtension } from "./ShapeExtension";

const EMU_PER_PIXEL = 9525;
const originalDocument = globalThis.document;
const window = new Window();
const shapeSpec = ShapeExtension().config.nodeSpec;
const schema = new Schema({
  nodes: { doc: { content: "shape*" }, text: {}, shape: shapeSpec },
});

beforeEach(() => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: window.document,
  });
});

afterEach(() => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: originalDocument,
  });
});

type RenderShapeOptions = {
  width: number;
  height: number;
  preset: "line" | "straightConnector1" | "rect";
};

const renderShape = ({ width, height, preset }: RenderShapeOptions) => {
  const xml = parseXmlDocument(`
    <wps:wsp xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
      <wps:spPr>
        <a:xfrm><a:ext cx="${width * EMU_PER_PIXEL}" cy="${height * EMU_PER_PIXEL}"/></a:xfrm>
        <a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>
        <a:noFill/><a:ln w="57150"/>
      </wps:spPr>
    </wps:wsp>
  `);
  if (!xml) throw new Error("Synthetic shape must parse");
  const shape = parseShape(xml);
  const node = schema.node("shape", {
    shapeType: shape.shapeType,
    width: shape.size.width / EMU_PER_PIXEL,
    height: shape.size.height / EMU_PER_PIXEL,
    outlineWidth: (shape.outline?.width ?? 0) / EMU_PER_PIXEL,
    fillType: "none",
  });
  const output = shapeSpec.toDOM?.(node);
  if (!output || typeof output !== "object" || !("dom" in output)) {
    throw new Error("Shape renderer must return a DOM element");
  }
  if (!(output.dom instanceof window.HTMLElement)) {
    throw new Error("Shape renderer must return an HTML wrapper");
  }
  return output.dom;
};

describe("straight shape paint bounds", () => {
  for (const preset of ["line", "straightConnector1"] as const) {
    for (const width of [0, 1, 63]) {
      for (const height of [0, 1, 87]) {
        if (width === 0 && height === 0) continue;
        test(`${preset} keeps its endpoints and visible viewport at ${width} by ${height}`, () => {
          const wrapper = renderShape({ width, height, preset });
          const svg = wrapper.querySelector("svg");
          const line = svg?.querySelector("line");
          expect(line?.getAttribute("x1")).toBe("0");
          expect(line?.getAttribute("y1")).toBe("0");
          expect(line?.getAttribute("x2")).toBe(String(width));
          expect(line?.getAttribute("y2")).toBe(String(height));
          expect(Number(svg?.getAttribute("width"))).toBeGreaterThan(0);
          expect(Number(svg?.getAttribute("height"))).toBeGreaterThan(0);
          expect(svg?.getAttribute("viewBox")).toBe(
            `0 0 ${width === 0 ? 1 : width} ${height === 0 ? 1 : height}`,
          );
          expect(svg?.style.overflow).toBe("visible");
          expect(wrapper.style.width).toBe(`${width}px`);
          expect(wrapper.style.height).toBe(`${height}px`);
        });
      }
    }
  }

  test("a centered rectangle outline can extend past its authored layout box", () => {
    const wrapper = renderShape({ width: 63, height: 87, preset: "rect" });
    const svg = wrapper.querySelector("svg");
    expect(svg?.style.strokeWidth).toBe("6");
    expect(svg?.style.overflow).toBe("visible");
    expect(svg?.getAttribute("viewBox")).toBe("0 0 63 87");
    expect(wrapper.style.width).toBe("63px");
    expect(wrapper.style.height).toBe("87px");
  });
});
