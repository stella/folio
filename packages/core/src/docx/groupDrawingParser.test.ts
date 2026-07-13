import { describe, expect, test } from "bun:test";

import type { MediaFile, RelationshipMap } from "../types/document";
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

  test("uses grouped shape style outline colour when geometry omits a direct colour", () => {
    const drawing = parseXmlDocument(`
      <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup"
        xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
        <wp:anchor>
          <wp:extent cx="2000000" cy="100000"/>
          <wp:wrapTopAndBottom/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp><wps:wsp>
              <wps:spPr>
                <a:xfrm><a:off x="0" y="0"/><a:ext cx="2000000" cy="100000"/></a:xfrm>
                <a:custGeom><a:pathLst><a:path w="2000000" h="100000">
                  <a:moveTo><a:pt x="0" y="50000"/></a:moveTo>
                  <a:lnTo><a:pt x="2000000" y="50000"/></a:lnTo>
                </a:path></a:pathLst></a:custGeom>
                <a:ln w="20000"/>
              </wps:spPr>
              <wps:style>
                <a:lnRef idx="1"><a:srgbClr val="123456"/></a:lnRef>
                <a:fillRef idx="0"><a:srgbClr val="ABCDEF"/></a:fillRef>
              </wps:style>
            </wps:wsp></wpg:wgp>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing>
    `);

    if (!drawing) {
      throw new Error("Expected drawing fixture");
    }
    const image = parseGroupDrawing(drawing);
    const svg = decodeURIComponent(image?.src?.split(",").at(1) ?? "");

    expect(svg).toContain('fill="none" stroke="#123456" stroke-width="20000"');
  });

  test("composes grouped pictures within the authored group extent", () => {
    const drawing = parseXmlDocument(`
      <w:drawing xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
        xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
        xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
        xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"
        xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
        xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
        <wp:anchor behindDoc="1">
          <wp:extent cx="2000000" cy="1000000"/>
          <wp:wrapNone/>
          <a:graphic><a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup">
            <wpg:wgp>
              <wpg:grpSpPr><a:xfrm><a:chOff x="100" y="200"/><a:chExt cx="1800000" cy="800000"/></a:xfrm></wpg:grpSpPr>
              <pic:pic>
                <pic:blipFill><a:blip r:embed="rId1"/><a:srcRect l="25000"/></pic:blipFill>
                <pic:spPr><a:xfrm><a:off x="100" y="200"/><a:ext cx="900000" cy="800000"/></a:xfrm></pic:spPr>
              </pic:pic>
              <pic:pic>
                <pic:blipFill><a:blip r:embed="rId2"/></pic:blipFill>
                <pic:spPr><a:xfrm><a:off x="1000000" y="200"/><a:ext cx="900000" cy="800000"/></a:xfrm></pic:spPr>
              </pic:pic>
            </wpg:wgp>
          </a:graphicData></a:graphic>
        </wp:anchor>
      </w:drawing>
    `);
    if (!drawing) {
      throw new Error("Expected drawing fixture");
    }
    const rels: RelationshipMap = new Map([
      [
        "rId1",
        {
          id: "rId1",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/first.png",
        },
      ],
      [
        "rId2",
        {
          id: "rId2",
          type: "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
          target: "media/second.png",
        },
      ],
    ]);
    const media = new Map<string, MediaFile>([
      [
        "word/media/first.png",
        {
          path: "word/media/first.png",
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: "data:image/png;base64,Zmlyc3Q=",
        },
      ],
      [
        "word/media/second.png",
        {
          path: "word/media/second.png",
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: "data:image/png;base64,c2Vjb25k",
        },
      ],
    ]);

    const image = parseGroupDrawing(drawing, rels, media);
    expect(image?.size).toEqual({ width: 2_000_000, height: 1_000_000 });
    const svg = decodeURIComponent(image?.src?.split(",").at(1) ?? "");
    expect(svg).toContain('viewBox="100 200 1800000 800000"');
    expect(svg.match(/<image /gu)).toHaveLength(2);
    expect(svg).toContain('href="data:image/png;base64,Zmlyc3Q="');
    expect(svg).toContain('href="data:image/png;base64,c2Vjb25k"');
    expect(svg).toContain('<clipPath id="group-picture-1">');
  });
});
