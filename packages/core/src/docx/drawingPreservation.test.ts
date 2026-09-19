// Drawings Folio cannot model must survive a round trip as verbatim raw XML.
// Anything the dispatch neither models nor preserves is lost on save, so each
// case here asserts both that the run keeps the content and that the content
// carries preservation-only raw XML.

import { describe, expect, test } from "bun:test";

import type { RunContent } from "../types/document";
import { parseDocumentBody } from "./documentParser";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const NS = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:v="urn:schemas-microsoft-com:vml"',
  'xmlns:o="urn:schemas-microsoft-com:office:office"',
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
].join("\n  ");

const documentWith = (runXml: string): string => `${XML_DECLARATION}
<w:document
  ${NS}>
  <w:body><w:p><w:r>${runXml}</w:r></w:p></w:body>
</w:document>`;

const runContentsOf = (xml: string): RunContent[] => {
  const body = parseDocumentBody(xml);
  const paragraph = body.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected a paragraph");
  }
  return paragraph.content
    .filter((content) => content.type === "run")
    .flatMap((run) => run.content);
};

describe("group drawings the rasterizer declines", () => {
  // parseGroupDrawing renders a wpg:wgp preview only when every child it walks
  // produces geometry; a group of unknown child kinds yields no SVG content.
  const unrenderableGroup = `
    <w:drawing>
      <wp:anchor behindDoc="1">
        <wp:extent cx="2000000" cy="1000000"/>
        <wp:wrapNone/>
        <wp:docPr id="3" name="Group 3"/>
        <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
          <wpg:wgp>
            <wpg:cNvGrpSpPr/>
            <wpg:grpSpPr><a:xfrm><a:chOff x="0" y="0"/><a:chExt cx="2000000" cy="1000000"/></a:xfrm></wpg:grpSpPr>
            <wpg:graphicFrame>
              <wpg:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="1000000"/></wpg:xfrm>
            </wpg:graphicFrame>
          </wpg:wgp>
        </a:graphicData></a:graphic>
      </wp:anchor>
    </w:drawing>`;

  test("keeps the whole group as preservation-only raw XML", () => {
    const contents = runContentsOf(documentWith(unrenderableGroup));

    expect(contents).toHaveLength(1);
    const drawing = contents.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected a drawing");
    }
    expect(drawing.rawXmlMode).toBe("preserveOnly");
    expect(drawing.rawXml).toContain("wpg:graphicFrame");
    expect(drawing.rawXml).toContain('<wp:docPr id="3" name="Group 3"/>');
    expect(drawing.image.rId).toBeUndefined();
  });

  test("a group with an embedded picture is not re-modeled as that picture", () => {
    const contents = runContentsOf(
      documentWith(`
        <w:drawing>
          <wp:inline>
            <wp:extent cx="2000000" cy="1000000"/>
            <wp:docPr id="4" name="Group 4"/>
            <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
              <wpg:wgp>
                <wpg:graphicFrame>
                  <a:blip r:embed="rId9"/>
                </wpg:graphicFrame>
              </wpg:wgp>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>`),
    );

    expect(contents).toHaveLength(1);
    const drawing = contents.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected a drawing");
    }
    expect(drawing.rawXmlMode).toBe("preserveOnly");
    expect(drawing.image.rId).toBeUndefined();
  });
});

describe("shapes carrying properties the model drops", () => {
  test("keeps a shape with a 3-D scene as preservation-only raw XML", () => {
    const contents = runContentsOf(
      documentWith(`
        <w:drawing>
          <wp:inline>
            <wp:extent cx="914400" cy="457200"/>
            <wp:docPr id="5" name="Shape 5"/>
            <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
              <wps:wsp>
                <wps:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
                  <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                  <a:scene3d><a:camera prst="perspectiveFront"/></a:scene3d>
                  <a:sp3d extrusionH="57150"/>
                </wps:spPr>
                <wps:bodyPr/>
              </wps:wsp>
            </a:graphicData></a:graphic>
          </wp:inline>
        </w:drawing>`),
    );

    expect(contents).toHaveLength(1);
    const drawing = contents.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected a drawing");
    }
    expect(drawing.rawXmlMode).toBe("preserveOnly");
    expect(drawing.rawXml).toContain('<a:sp3d extrusionH="57150"/>');
    expect(drawing.rawXml).toContain('<a:camera prst="perspectiveFront"/>');
  });
});

describe("mc:AlternateContent with nothing the parser can model", () => {
  test("keeps the whole AlternateContent as preservation-only raw XML", () => {
    const contents = runContentsOf(
      documentWith(`
        <mc:AlternateContent>
          <mc:Choice Requires="v">
            <w:pict>
              <v:line id="Line 1" style="position:absolute" from="0,0" to="200,0" strokecolor="#123456"/>
            </w:pict>
          </mc:Choice>
          <mc:Fallback>
            <w:pict>
              <v:line id="Line 1 fallback" style="position:absolute" from="0,0" to="200,0"/>
            </w:pict>
          </mc:Fallback>
        </mc:AlternateContent>`),
    );

    expect(contents).toHaveLength(1);
    const drawing = contents.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected a drawing");
    }
    expect(drawing.rawXmlMode).toBe("preserveOnly");
    expect(drawing.rawXml).toContain("mc:AlternateContent");
    expect(drawing.rawXml).toContain('<v:line id="Line 1"');
    expect(drawing.rawXml).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
  });

  test("text beside an unmodelable branch keeps its order", () => {
    const contents = runContentsOf(
      documentWith(`<w:t>before</w:t>
        <mc:AlternateContent>
          <mc:Choice Requires="v">
            <w:pict><v:line id="Line 2" from="0,0" to="10,0"/></w:pict>
          </mc:Choice>
        </mc:AlternateContent>
        <w:t>after</w:t>`),
    );

    expect(contents.map((content) => content.type)).toEqual(["text", "drawing", "text"]);
  });
});

describe("w:pict VML shapes with no image relationship", () => {
  test("keeps a geometry-only v:shape as preservation-only raw XML", () => {
    const contents = runContentsOf(
      documentWith(`
        <w:pict>
          <v:shape id="Freeform 1" style="width:120pt;height:40pt" coordsize="2400,800"
            path="m0,0l2400,800e" fillcolor="#123456" strokecolor="#654321">
            <v:path arrowok="t"/>
          </v:shape>
        </w:pict>`),
    );

    expect(contents).toHaveLength(1);
    const drawing = contents.at(0);
    if (drawing?.type !== "drawing") {
      throw new Error("Expected a drawing");
    }
    expect(drawing.rawXmlMode).toBe("preserveOnly");
    expect(drawing.rawXml).toContain('<v:shape id="Freeform 1"');
    expect(drawing.rawXml).toContain('xmlns:v="urn:schemas-microsoft-com:vml"');
  });

  test("leaves a VML text box to the text-box pipeline", () => {
    const contents = runContentsOf(
      documentWith(`
        <w:pict>
          <v:shape id="TextBox 1" style="width:120pt;height:40pt">
            <v:textbox><w:txbxContent><w:p><w:r><w:t>boxed</w:t></w:r></w:p></w:txbxContent></v:textbox>
          </v:shape>
        </w:pict>`),
    );

    expect(contents.filter((content) => content.type === "drawing")).toHaveLength(0);
    const shape = contents.find((content) => content.type === "shape");
    expect(shape?.type === "shape" ? shape.shape.shapeType : undefined).toBe("textBox");
  });
});
