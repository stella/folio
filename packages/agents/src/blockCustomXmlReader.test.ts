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
    `<w:document xmlns:w="${W}"><w:body>` +
      `<w:customXml w:element="clause">` +
      `<w:p><w:r><w:t>First clause.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Second clause.</w:t></w:r></w:p>` +
      `</w:customXml>` +
      `<w:p><w:r><w:t>After.</w:t></w:r></w:p>` +
      `</w:body></w:document>`,
  );
  return zip.generateAsync({ type: "arraybuffer" });
};

describe("read_document inside block custom XML", () => {
  test("rows agree with the reviewer's snapshot", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await makeDocx());
    const bridge = createReviewerBridge(reviewer);
    const read = executeFolioToolCall("read_document", {}, bridge);
    if (!read.ok) {
      throw new Error(read.error);
    }
    const expected = ["First clause.", "Second clause.", "After."];
    expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(expected);
    expect(read.result.map(({ text }) => text)).toEqual(expected);
  });
});
