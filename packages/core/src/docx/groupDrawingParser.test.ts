import { describe, expect, test } from "bun:test";

import { parseGroupDrawing } from "./groupDrawingParser";
import { parseXmlDocument } from "./xmlParser";

describe("parseGroupDrawing", () => {
  test("renders grouped geometry and text as an anchored SVG image", () => {
    const drawing = parseXmlDocument(`
      <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wp:anchor behindDoc="1">
          <wp:positionH relativeFrom="page"><wp:posOffset>914400</wp:posOffset></wp:positionH>
          <wp:positionV relativeFrom="paragraph"><wp:posOffset>252000</wp:posOffset></wp:positionV>
          <wp:extent cx="2000000" cy="1000000"/>
          <wp:wrapTopAndBottom/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>
              <wps:wsp>
                <wps:spPr>
                  <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="500000"/></a:xfrm>
                  <a:custGeom><a:pathLst><a:path w="2000000" h="500000"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="2000000" y="500000"/></a:lnTo></a:path></a:pathLst></a:custGeom>
                  <a:solidFill><a:srgbClr val="DBEDF3"/></a:solidFill>
                  <a:ln><a:solidFill><a:srgbClr val="123456"/></a:solidFill></a:ln>
                </wps:spPr>
              </wps:wsp>
              <wps:wsp>
                <wps:spPr><a:xfrm><a:off x="100000" y="600000"/><a:ext cx="900000" cy="200000"/></a:xfrm></wps:spPr>
                <wps:txbx><w:txbxContent><w:p><w:r><w:t>A &amp; B</w:t></w:r></w:p></w:txbxContent></wps:txbx>
              </wps:wsp>
            </wpg:wgp>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing>
    `);

    if (!drawing) {
      throw new Error("Expected drawing fixture");
    }
    const image = parseGroupDrawing(drawing);

    expect(image?.size).toEqual({ width: 2_000_000, height: 1_000_000 });
    expect(image?.wrap.type).toBe("topAndBottom");
    expect(image?.src).toStartWith("data:image/svg+xml");
    const svg = decodeURIComponent(image?.src?.split(",").at(1) ?? "");
    expect(svg).toContain('<path d="M 0 0 L 2000000 500000"');
    expect(svg).toContain('stroke="#123456" stroke-width="9525"');
    expect(svg).toContain("A &amp; B");
    expect(svg).not.toContain("A & B");
    expect(svg).not.toContain('<rect width="2000000" height="1000000" fill="#FFFFFF"');
  });
});
