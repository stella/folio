import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { fromProseDoc } from "../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { InlineSdt, Paragraph, Run } from "../types/document";
import { parseDocx } from "./parser";
import { createEmptyDocx, repackDocx } from "./rezip";

const textBoxDrawingXml = `
  <w:drawing>
    <wp:inline>
      <wp:extent cx="1828800" cy="914400"/>
      <wp:docPr id="11" name="Text Box 11"/>
      <a:graphic>
        <a:graphicData uri="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
          <wps:wsp>
            <wps:spPr>
              <a:xfrm><a:off x="0" y="0"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
              <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
            </wps:spPr>
            <wps:txbx>
              <w:txbxContent>
                <w:p><w:r><w:t>Controlled text box</w:t></w:r></w:p>
              </w:txbxContent>
            </wps:txbx>
            <wps:bodyPr/>
          </wps:wsp>
        </a:graphicData>
      </a:graphic>
    </wp:inline>
  </w:drawing>`;

const documentXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
  xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
  xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape">
  <w:body>
    <w:p>
      <w:sdt>
        <w:sdtPr><w:alias w:val="Outer"/><w:richText/></w:sdtPr>
        <w:sdtContent>
          <w:r><w:t>Before </w:t></w:r>
          <w:sdt>
            <w:sdtPr><w:alias w:val="Inner"/><w:tag w:val="inner"/><w:richText/></w:sdtPr>
            <w:sdtContent><w:r>${textBoxDrawingXml}</w:r></w:sdtContent>
          </w:sdt>
          <w:r><w:t> after</w:t></w:r>
        </w:sdtContent>
      </w:sdt>
    </w:p>
    <w:sectPr/>
  </w:body>
</w:document>`;

const createSource = async (): Promise<ArrayBuffer> => {
  const zip = await JSZip.loadAsync(await createEmptyDocx());
  zip.file("word/document.xml", documentXml);
  return zip.generateAsync({ type: "arraybuffer" });
};

const firstParagraph = (document: Awaited<ReturnType<typeof parseDocx>>): Paragraph => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph") {
    throw new Error("Expected paragraph");
  }
  return paragraph;
};

const firstInlineSdt = (content: Paragraph["content"] | InlineSdt["content"]): InlineSdt => {
  const sdt = content.find((item) => item.type === "inlineSdt");
  if (sdt?.type !== "inlineSdt") {
    throw new Error("Expected inline content control");
  }
  return sdt;
};

const textBoxRuns = (sdt: InlineSdt): Run[] =>
  sdt.content.filter(
    (item): item is Run =>
      item.type === "run" &&
      item.content.some((content) => content.type === "shape" && content.shape.textBody),
  );

describe("inline content-control text boxes", () => {
  test("preserves nested controls and drawings through the editor save path", async () => {
    const source = await createSource();
    const parsed = await parseDocx(source, { detectVariables: false, preloadFonts: false });
    const outer = firstInlineSdt(firstParagraph(parsed).content);
    const inner = firstInlineSdt(outer.content);

    expect(textBoxRuns(inner)).toHaveLength(1);

    const pmDocument = toProseDoc(parsed);
    expect(pmDocument.childCount).toBe(2);
    expect(pmDocument.child(1).type.name).toBe("textBox");
    expect(pmDocument.child(1).attrs["_docxInlineSdts"]).toMatchObject([
      { sdtType: "richText", alias: "Outer" },
      { sdtType: "richText", alias: "Inner", tag: "inner" },
    ]);

    const saved = await repackDocx(fromProseDoc(pmDocument, parsed), {
      updateModifiedDate: false,
    });
    const savedZip = await JSZip.loadAsync(saved);
    const savedXml = await savedZip.file("word/document.xml")?.async("text");
    expect(savedXml?.match(/<w:sdt>/gu)).toHaveLength(2);
    expect(savedXml).toContain("<wps:txbx>");

    const reparsed = await parseDocx(saved, {
      detectVariables: false,
      preloadFonts: false,
    });
    const reparsedOuter = firstInlineSdt(firstParagraph(reparsed).content);
    const reparsedInner = firstInlineSdt(reparsedOuter.content);
    expect(textBoxRuns(reparsedInner)).toHaveLength(1);
  });
});
