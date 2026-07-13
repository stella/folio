import { describe, expect, test } from "bun:test";

import type { HeaderFooter, Paragraph } from "../types/document";
import { parseFooter, parseHeader } from "./headerFooterParser";

const NAMESPACES = [
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"',
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"',
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"',
].join(" ");

const textBoxDrawingXml = (text: string): string => `
  <w:drawing>
    <wp:anchor distT="45720" distB="45720" distL="114300" distR="114300"
      simplePos="0" relativeHeight="251695104" behindDoc="0" locked="0"
      layoutInCell="1" allowOverlap="1">
      <wp:simplePos x="0" y="0"/>
      <wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH>
      <wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>
      <wp:extent cx="1390650" cy="278130"/>
      <wp:effectExtent l="0" t="0" r="0" b="0"/>
      <wp:wrapNone/>
      <wp:docPr id="1" name="Text Box 1"/>
      <wp:cNvGraphicFramePr/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1390650" cy="278130"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>${text}</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:anchor>
  </w:drawing>`;

const textBoxRunXml = (text: string, alternateContent: boolean): string => {
  const drawing = textBoxDrawingXml(text);

  if (!alternateContent) {
    return `<w:r>${drawing}</w:r>`;
  }

  return `
    <w:r>
      <mc:AlternateContent>
        <mc:Choice Requires="wps">${drawing}</mc:Choice>
        <mc:Fallback><w:pict/></mc:Fallback>
      </mc:AlternateContent>
    </w:r>`;
};

const headerFooterXml = (
  rootName: "hdr" | "ftr",
  text: string,
  alternateContent: boolean,
): string => `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
  <w:${rootName} ${NAMESPACES}>
    <w:p>${textBoxRunXml(text, alternateContent)}</w:p>
  </w:${rootName}>`;

const getShapeTexts = (headerFooter: HeaderFooter): string[] => {
  const texts: string[] = [];

  for (const block of headerFooter.content) {
    if (block.type !== "paragraph") {
      continue;
    }
    collectParagraphShapeTexts(block, texts);
  }

  return texts;
};

const collectParagraphShapeTexts = (paragraph: Paragraph, texts: string[]) => {
  for (const paragraphContent of paragraph.content) {
    if (paragraphContent.type !== "run") {
      continue;
    }

    for (const runContent of paragraphContent.content) {
      if (runContent.type !== "shape") {
        continue;
      }

      const firstBlock = runContent.shape.textBody?.content.at(0);
      if (firstBlock?.type !== "paragraph") {
        continue;
      }

      const firstRun = firstBlock.content.at(0);
      if (firstRun?.type !== "run") {
        continue;
      }

      const firstContent = firstRun.content.at(0);
      if (firstContent?.type === "text") {
        texts.push(firstContent.text);
      }
    }
  }
};

describe("header/footer text boxes", () => {
  test("parses AlternateContent-wrapped text boxes in headers", () => {
    const header = parseHeader(headerFooterXml("hdr", "Header Box", true));

    expect(getShapeTexts(header)).toEqual(["Header Box"]);
  });

  test("parses bare drawing text boxes in headers", () => {
    const header = parseHeader(headerFooterXml("hdr", "Bare Header Box", false));

    expect(getShapeTexts(header)).toEqual(["Bare Header Box"]);
  });

  test("parses AlternateContent-wrapped text boxes in footers", () => {
    const footer = parseFooter(headerFooterXml("ftr", "Footer Box", true));

    expect(getShapeTexts(footer)).toEqual(["Footer Box"]);
  });

  test("parses positioned VML text boxes in footers", () => {
    const footer = parseFooter(`
      <w:ftr ${NAMESPACES}>
        <w:p><w:r><w:pict>
          <v:shape id="footer-box" style="position:absolute;margin-left:72pt;margin-top:720pt;width:216pt;height:18pt;z-index:-1;mso-position-horizontal-relative:page;mso-position-vertical-relative:page" filled="f" stroked="f">
            <v:textbox inset="0,0,0,0">
              <w:txbxContent><w:p><w:r><w:t>Positioned footer</w:t></w:r></w:p></w:txbxContent>
            </v:textbox>
          </v:shape>
        </w:pict></w:r></w:p>
      </w:ftr>`);

    expect(getShapeTexts(footer)).toEqual(["Positioned footer"]);
    const paragraph = footer.content.at(0);
    const run = paragraph?.type === "paragraph" ? paragraph.content.at(0) : undefined;
    const shapeContent = run?.type === "run" ? run.content.at(0) : undefined;
    if (shapeContent?.type !== "shape") {
      throw new Error("Expected positioned text box");
    }
    expect(shapeContent.shape).toMatchObject({
      id: "footer-box",
      shapeType: "textBox",
      size: { width: 2_743_200, height: 228_600 },
      position: {
        horizontal: { relativeTo: "page", posOffset: 914_400 },
        vertical: { relativeTo: "page", posOffset: 9_144_000 },
      },
      wrap: { type: "behind" },
      fill: { type: "none" },
      textBody: { margins: { left: 0, top: 0, right: 0, bottom: 0 } },
    });
  });
});
