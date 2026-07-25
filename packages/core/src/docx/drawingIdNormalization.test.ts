import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import type { Document } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const XML_DECLARATION = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';

const textBoxDrawing = (text: string, id?: string): string => `
  <w:drawing>
    <wp:inline>
      <wp:extent cx="914400" cy="457200"/>
      ${id === undefined ? "" : `<wp:docPr id="${id}" name="Text box ${id}"/>`}
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:cNvSpPr txBox="1"/>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="457200"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent><w:p><w:r><w:t>${text}</w:t></w:r></w:p></w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;

const buildDocx = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file(
    "word/document.xml",
    `${XML_DECLARATION}
    <w:document
      xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
      xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
      xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
      <w:body>
        <w:p><w:r>${textBoxDrawing("Generated")}</w:r></w:p>
        <w:p><w:r>${textBoxDrawing("Authored", "100000")}</w:r></w:p>
        <w:sectPr/>
      </w:body>
    </w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

const shapeIds = ({ package: { document } }: Document): (string | undefined)[] =>
  document.content.flatMap((block) => {
    if (block.type !== "paragraph") {
      return [];
    }
    return block.content.flatMap((content) => {
      if (content.type !== "run") {
        return [];
      }
      return content.content.flatMap((runContent) =>
        runContent.type === "shape" ? [runContent.shape.id] : [],
      );
    });
  });

describe("drawing ID normalization", () => {
  test("assigns missing shape IDs without colliding and remains stable after save", async () => {
    const parsed = await parseDocx(await buildDocx(), { preloadFonts: false });
    expect(shapeIds(parsed)).toEqual(["100001", "100000"]);

    const saved = await repackDocx(parsed, { updateModifiedDate: false });
    const reopened = await parseDocx(saved, { preloadFonts: false });
    expect(shapeIds(reopened)).toEqual(["100001", "100000"]);
  });
});
