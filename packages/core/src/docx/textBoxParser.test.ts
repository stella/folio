import { describe, expect, test } from "bun:test";

import { parseTextBox } from "./textBoxParser";
import { parseXmlDocument } from "./xmlParser";

const NAMESPACES = `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"`;

const drawingWithBodyProperties = (bodyProperties: string): string => `<w:drawing ${NAMESPACES}>
  <wp:inline>
    <wp:extent cx="914400" cy="457200"/>
    <a:graphic>
      <a:graphicData>
        <wps:wsp>
          <wps:spPr/>
          <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
          ${bodyProperties}
        </wps:wsp>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>`;

describe("text box body properties", () => {
  test.each([
    ["<wps:bodyPr><a:spAutoFit/></wps:bodyPr>", "shape"],
    ["<wps:bodyPr><a:normAutofit/></wps:bodyPr>", "normal"],
    ["<wps:bodyPr><a:noAutofit/></wps:bodyPr>", "none"],
  ] as const)("parses the %s fitting mode", (bodyProperties, expected) => {
    const drawing = parseXmlDocument(drawingWithBodyProperties(bodyProperties));
    expect(drawing).not.toBeNull();
    if (!drawing) {
      return;
    }

    expect(parseTextBox(drawing)?.autoFit).toBe(expected);
  });

  test("leaves fitting unspecified when no mode is authored", () => {
    const drawing = parseXmlDocument(drawingWithBodyProperties("<wps:bodyPr/>"));
    expect(drawing).not.toBeNull();
    if (!drawing) {
      return;
    }

    expect(parseTextBox(drawing)?.autoFit).toBeUndefined();
  });
});
