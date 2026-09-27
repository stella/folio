import { describe, expect, test } from "bun:test";
import { FolioDocxReviewer } from "@stll/folio-core/server";
import JSZip from "jszip";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCall } from "./execute";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const makeDocx = async (): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "word/document.xml",
    `<w:document xmlns:w="${W}"><w:body><w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="1000"/></w:tblGrid>` +
      `<w:customXml w:element="row"><w:customXmlPr><w:attr w:name="row-key" w:val="row-value"/></w:customXmlPr><w:bookmarkStart w:id="31" w:name="row-wrapper-start"/><w:tr><w:customXml w:element="cell"><w:customXmlPr><w:attr w:name="cell-key" w:val="cell-value"/></w:customXmlPr><w:bookmarkStart w:id="32" w:name="cell-wrapper-start"/><w:tc><w:p><w:r><w:t>Table custom XML sentinel.</w:t></w:r></w:p></w:tc><w:bookmarkEnd w:id="32"/></w:customXml></w:tr><w:bookmarkEnd w:id="31"/></w:customXml>` +
      `<w:sdt><w:sdtPr/><w:sdtContent><w:tr><w:tc><w:p><w:r><w:t>Table SDT sentinel.</w:t></w:r></w:p></w:tc></w:tr></w:sdtContent></w:sdt>` +
      `<w:tr><w:sdt><w:sdtPr/><w:sdtContent><w:tc><w:p><w:r><w:t>Row SDT sentinel.</w:t></w:r></w:p></w:tc></w:sdtContent></w:sdt></w:tr>` +
      `<w:customXml w:element="duplicate"><w:tr><w:tc><w:p/></w:tc></w:tr></w:customXml>` +
      `<w:customXml w:element="duplicate"><w:tr><w:tc><w:p/></w:tc></w:tr></w:customXml>` +
      `</w:tbl></w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("read_document inside table custom XML", () => {
  test("row text remains visible in the snapshot and read_document", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await makeDocx());
    const bridge = createReviewerBridge(reviewer);
    const read = executeFolioToolCall("read_document", {}, bridge);
    if (!read.ok) throw new Error(read.error);

    const expected = ["Table custom XML sentinel.", "Table SDT sentinel.", "Row SDT sentinel."];
    expect(
      reviewer
        .snapshot()
        .blocks.map(({ text }) => text)
        .filter(Boolean),
    ).toEqual(expected);
    expect(read.result.map(({ text }) => text).filter(Boolean)).toEqual(expected);

    const editableBlock = read.result.find(({ text }) => text === expected[0]);
    if (!editableBlock) throw new Error("Expected the custom XML row to be editable");
    const edit = executeFolioToolCall(
      "suggest_changes",
      {
        operations: [
          {
            type: "replaceInBlock",
            blockId: editableBlock.blockId,
            find: "Table custom XML",
            replace: "Edited custom XML",
          },
        ],
      },
      bridge,
    );
    if (!edit.ok) throw new Error(edit.error);

    const saved = await reviewer.toBuffer();
    const xml =
      (await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("text")) ?? "";
    expect(xml).toContain('<w:customXml w:element="row">');
    expect(xml).toContain('<w:customXml w:element="cell">');
    const rowXml = xml.split('<w:customXml w:element="row">')[1]?.split("</w:customXml>")[0];
    expect(rowXml).toContain("<w:customXmlPr>");
    expect(rowXml).toContain('w:name="row-wrapper-start"');
    const cellXml = xml.split('<w:customXml w:element="cell">')[1]?.split("</w:customXml>")[0];
    expect(cellXml).toContain("<w:customXmlPr>");
    expect(cellXml).toContain('w:name="cell-wrapper-start"');
    expect((xml.match(/<w:customXml w:element="duplicate">/gu) ?? []).length).toBe(2);
    expect(xml).toContain("Edited");
  });
});
