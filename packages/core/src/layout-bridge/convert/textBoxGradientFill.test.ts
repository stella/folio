/**
 * A text box's `a:gradFill` (ECMA-376 §20.1.8.33) reaches the layout, and the
 * editor saves it back. A text box used to carry only a solid `fillColor`, so a
 * gradient-filled box painted and saved with no fill at all.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { parseDocx } from "../../docx/parser";
import { createEmptyDocx, repackDocx } from "../../docx/rezip";
import type { TextBoxBlock } from "../../layout-engine/types";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import { toFlowBlocks } from "./toFlowBlocks";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const GRADIENT_FILL = `<a:gradFill rotWithShape="0">
  <a:gsLst>
    <a:gs pos="100000"><a:schemeClr val="accent1"/></a:gs>
    <a:gs pos="0"><a:srgbClr val="102030"/></a:gs>
    <a:gs pos="40000"><a:srgbClr val="405060"/></a:gs>
  </a:gsLst>
  <a:lin ang="18900000" scaled="1"/>
</a:gradFill>`;

const documentXml = (fill: string): string => `${XML_DECLARATION}
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:r>
        <w:drawing>
          <wp:anchor simplePos="0" relativeHeight="1" behindDoc="1" locked="0" layoutInCell="1" allowOverlap="1">
            <wp:simplePos x="0" y="0"/>
            <wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH>
            <wp:positionV relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionV>
            <wp:extent cx="1828800" cy="7315200"/>
            <wp:wrapNone/>
            <wp:docPr id="1" name="Band"/>
            <a:graphic>
              <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
                <wps:wsp>
                  <wps:spPr>
                    <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="7315200"/></a:xfrm>
                    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
                    ${fill}
                  </wps:spPr>
                  <wps:txbx><w:txbxContent><w:p/></w:txbxContent></wps:txbx>
                  <wps:bodyPr/>
                </wps:wsp>
              </a:graphicData>
            </a:graphic>
          </wp:anchor>
        </w:drawing>
      </w:r>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const buildDocx = async (fill: string): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml(fill));
  return zip.generateAsync({ type: "arraybuffer" });
};

const textBoxBlocks = (pmDoc: ReturnType<typeof toProseDoc>, theme: unknown): TextBoxBlock[] =>
  toFlowBlocks(pmDoc, { theme: theme as never }).filter(
    (block): block is TextBoxBlock => block.kind === "textBox",
  );

describe("text box gradient fill", () => {
  test("a linear gradient reaches the text box block with its stops resolved and ordered", async () => {
    const document = await parseDocx(await buildDocx(GRADIENT_FILL));
    const theme = document.package.theme ?? null;
    const [box] = textBoxBlocks(toProseDoc(document), theme);

    expect(box?.fillColor).toBeUndefined();
    expect(box?.fillGradient?.angle).toBe(315);
    expect(box?.fillGradient?.scaled).toBe(true);
    expect(box?.fillGradient?.stops.map(({ offset }) => offset)).toEqual([0, 0.4, 1]);
    expect(box?.fillGradient?.stops.slice(0, 2).map(({ color }) => color)).toEqual([
      "#102030",
      "#405060",
    ]);
    // The scheme color resolves through the theme rather than painting black.
    expect(box?.fillGradient?.stops[2]?.color).toMatch(/^#[0-9A-F]{6}$/iu);
    expect(box?.fillGradient?.stops[2]?.color).not.toBe("#000000");
  });

  test("an unstated a:lin@scaled paints the angle as stated", async () => {
    const document = await parseDocx(await buildDocx(GRADIENT_FILL.replace(' scaled="1"', "")));
    const [box] = textBoxBlocks(toProseDoc(document), document.package.theme ?? null);

    expect(box?.fillGradient?.scaled).toBe(false);
  });

  test("a path gradient is not painted as a linear one", async () => {
    const document = await parseDocx(
      await buildDocx(
        GRADIENT_FILL.replace(
          '<a:lin ang="18900000" scaled="1"/>',
          '<a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path>',
        ),
      ),
    );
    const [box] = textBoxBlocks(toProseDoc(document), document.package.theme ?? null);

    expect(box).toBeDefined();
    expect(box?.fillGradient).toBeUndefined();
  });

  test("the gradient survives an editor round trip and a save", async () => {
    const document = await parseDocx(await buildDocx(GRADIENT_FILL));
    const edited = fromProseDoc(toProseDoc(document), document);
    const saved = await repackDocx(edited, { updateModifiedDate: false });
    const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")!.async("text");

    expect(xml).toContain("<a:gradFill");
    expect(xml).toContain('<a:lin ang="18900000" scaled="1"/>');
    expect(xml).toContain('<a:srgbClr val="405060"/>');
  });
});
