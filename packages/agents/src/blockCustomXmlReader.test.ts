import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";
import { FolioDocxReviewer } from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCall } from "./execute";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const SESSIONS = ["legacy", "canonical"] as const;

const makeDocx = async (): Promise<ArrayBuffer> => {
  const zip = new ZipArchive();
  zip.set(
    "[Content_Types].xml",
    '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>",
  );
  zip.set(
    "_rels/.rels",
    '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships>",
  );
  zip.set(
    "word/document.xml",
    `<w:document xmlns:w="${W}"><w:body>` +
      `<w:customXml w:element="clause">` +
      `<w:p><w:r><w:t>First clause.</w:t></w:r></w:p>` +
      `<w:p><w:r><w:t>Second clause.</w:t></w:r></w:p>` +
      `</w:customXml>` +
      `<w:p><w:r><w:t>After.</w:t></w:r></w:p>` +
      `</w:body></w:document>`,
  );
  return zip.generate({ compression: "store" }).then((bytes) => bytes.buffer);
};

describe("read_document inside block custom XML", () => {
  test.each(SESSIONS)(
    "rows agree with the reviewer's snapshot (%s)",
    async (session) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await makeDocx(), {
        session,
      });
      const bridge = createReviewerBridge(reviewer);
      const read = executeFolioToolCall("read_document", {}, bridge);
      if (!read.ok) {
        throw new Error(read.error);
      }
      const expected = ["First clause.", "Second clause.", "After."];
      expect(reviewer.snapshot().blocks.map(({ text }) => text)).toEqual(
        expected,
      );
      expect(read.result.map(({ text }) => text)).toEqual(expected);
    },
  );

  test.each(SESSIONS)(
    "an edit inside the wrapper saves the wrapper (%s)",
    async (session) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await makeDocx(), {
        session,
      });
      const bridge = createReviewerBridge(reviewer);
      const blockId = reviewer.snapshot().blocks.at(1)?.id;
      if (blockId === undefined) throw new Error("Expected a second clause");
      const edit = executeFolioToolCall(
        "suggest_changes",
        {
          operations: [
            {
              type: "replaceInBlock",
              blockId,
              find: "Second",
              replace: "Edited",
            },
          ],
        },
        bridge,
      );
      if (!edit.ok) throw new Error(edit.error);
      expect(edit.result.applied).toHaveLength(1);

      const saved = await ZipArchive.read(await reviewer.toBuffer());
      const xml = (await saved.entry("word/document.xml")?.text()) ?? "";
      const wrapped = xml
        .split('<w:customXml w:element="clause">')[1]
        ?.split("</w:customXml>")[0];
      expect(wrapped).toContain("First clause.");
      expect(wrapped).toContain("Edited");
      expect(xml.split("</w:customXml>")[1]).toContain("After.");
    },
  );
});
