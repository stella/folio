/**
 * A shape outline written from the model, without source markup to replay
 * (a shape read from VML, or one an edit rebuilt), must be the element the
 * parser's verbatim capture of it replays: otherwise the second save of an
 * unedited document differs from the first. And it states only what the
 * model states, so an attribute the source omitted is never written with
 * its schema default.
 */

import { describe, expect, test } from "bun:test";

import type { Shape } from "../types/document";
import { serializeRun } from "./serializer/runSerializer";
import { parseShapeFromDrawing } from "./shapeParser";
import { findDeep, parseXmlDocument } from "./xmlParser";

const DRAWINGML = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS =
  `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
  `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
  `xmlns:a="${DRAWINGML}" ` +
  `xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

const shapeDrawing = (outline: string): string =>
  `<w:drawing ${NS}><wp:inline><wp:extent cx="914400" cy="457200"/>` +
  `<wp:docPr id="9" name="Shape 9"/><a:graphic>` +
  `<a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">` +
  `<wps:wsp><wps:cNvPr id="9" name="Shape 9"/><wps:spPr>` +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${outline}</wps:spPr>` +
  `<wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing>`;

const parsedShape = (drawingXml: string): Shape => {
  // A serializer returns a fragment; the readers resolve `wp:` and `wps:` by
  // namespace, so the bindings a package carries are declared on a wrapper.
  const root = parseXmlDocument(`<fixture ${NS}>${drawingXml}</fixture>`);
  const drawing = root === null ? null : findDeep(root, "w", "drawing");
  const shape = drawing ? parseShapeFromDrawing(drawing) : null;
  if (!shape?.outline) throw new Error("the fixture's outline did not parse");
  return shape;
};

const savedRun = (shape: Shape): string =>
  serializeRun({ type: "run", content: [{ type: "shape", shape }] });

/** The outline's model with no source markup, as a VML import or an edit gives it. */
const modelOnly = (shape: Shape): Shape => {
  const { rawXml: _replayed, ...outline } = shape.outline ?? {};
  return { ...shape, outline };
};

/** Each outline in the element order the schema and the serializer share. */
const OUTLINES: Record<string, string> = {
  "a width alone": `<a:ln w="31750"/>`,
  "a width and a cap": `<a:ln w="9525" cap="rnd"/>`,
  "a join alone": `<a:ln><a:round/></a:ln>`,
  "a colour": `<a:ln w="12700"><a:solidFill><a:srgbClr val="006699"/></a:solidFill></a:ln>`,
  "a preset dash": `<a:ln><a:prstDash val="sysDot"/></a:ln>`,
  "line ends without a type": `<a:ln w="9525"><a:headEnd/><a:tailEnd w="sm" len="lg"/></a:ln>`,
  "line ends with types": `<a:ln w="9525"><a:headEnd type="none"/><a:tailEnd type="triangle" w="med" len="med"/></a:ln>`,
  "every child": `<a:ln w="25400" cap="sq"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/><a:miter/><a:headEnd type="arrow"/><a:tailEnd/></a:ln>`,
};

describe("a shape outline written from the model", () => {
  for (const [name, outline] of Object.entries(OUTLINES)) {
    test(`with ${name} is written as the source spelled it`, () => {
      const source = parsedShape(shapeDrawing(outline));

      const first = savedRun(modelOnly(source));

      // Nothing the source omitted is written, defaults included.
      expect(first).toContain(outline);
    });

    test(`with ${name} saves the same bytes the second time`, () => {
      const first = savedRun(modelOnly(parsedShape(shapeDrawing(outline))));
      const reopened = parsedShape(first);

      // The second save replays the parser's capture of the first save's
      // element, and the first save must already be that capture.
      expect(reopened.outline?.rawXml).toBe(outline);
      expect(savedRun(reopened)).toBe(first);
      expect(modelOnly(reopened).outline).toEqual(
        modelOnly(parsedShape(shapeDrawing(outline))).outline,
      );
    });
  }

  test("line ends the source gives no type keep no type", () => {
    const source = parsedShape(
      shapeDrawing(`<a:ln><a:headEnd/><a:tailEnd w="sm"/></a:ln>`),
    );

    expect(source.outline?.headEnd).toEqual({});
    expect(source.outline?.tailEnd).toEqual({ width: "sm" });
  });
});
