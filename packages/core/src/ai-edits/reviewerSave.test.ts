import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14_NS = "http://schemas.microsoft.com/office/word/2010/wordml";

const buildDocx = async (paragraphs: readonly string[]): Promise<ArrayBuffer> => {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
      "</Types>",
  );
  zip.file(
    "_rels/.rels",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
      "</Relationships>",
  );
  const body = paragraphs
    .map(
      (text, index) =>
        `<w:p w14:paraId="${(0x1000_0000 + index).toString(16).toUpperCase()}" w14:textId="77777777">` +
        `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
    )
    .join("");
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="${W_NS}" xmlns:w14="${W14_NS}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const firstBlockId = (reviewer: FolioDocxReviewer): string => {
  const block = reviewer.snapshot().blocks.at(0);
  if (!block) throw new Error("fixture has no blocks");
  return block.id;
};

describe("FolioDocxReviewer.save", () => {
  test("reports a selective save for an in-paragraph edit", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx(["Pay $50 now."]), {
      author: "Reviewer",
    });
    reviewer.applyOperations([
      {
        id: "edit",
        type: "replaceInBlock",
        blockId: firstBlockId(reviewer),
        find: "$50",
        replace: "$500",
      },
    ]);

    const result = await reviewer.save({ repack: "refuse" });

    expect(result.type).toBe("selective");
  });

  test("refuses or performs a full repack for a structural edit", async () => {
    const source = await buildDocx(["First.", "Second."]);
    const insert = async (): Promise<FolioDocxReviewer> => {
      const reviewer = await FolioDocxReviewer.fromBuffer(source, { author: "Reviewer" });
      reviewer.applyOperations([
        { id: "insert", type: "insertAfterBlock", blockId: firstBlockId(reviewer), text: "New." },
      ]);
      return reviewer;
    };

    const refused = await (await insert()).save({ repack: "refuse" });
    const repacked = await (await insert()).save();

    expect(refused).toEqual({ type: "repackRefused", reason: "structuralChange" });
    expect(repacked.type).toBe("full-repack");
    if (repacked.type !== "full-repack") return;
    expect(repacked.reason).toBe("structuralChange");
    const reopened = await FolioDocxReviewer.fromBuffer(repacked.buffer);
    expect(reopened.getContent().map(({ text }) => text)).toContain("New.");
  });

  test("stamps created comments and replies with the supplied date", async () => {
    const date = "2026-01-02T03:04:05Z";
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx(["Clause."]), {
      author: "Reviewer",
    });
    reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        operations: [
          {
            id: "note",
            type: "commentOnBlock",
            blockId: firstBlockId(reviewer),
            comment: { text: "Check." },
          },
        ],
      },
      { revisionStamp: { date, idSeed: 100 } },
    );
    const thread = reviewer.getComments().at(0);
    if (!thread) throw new Error("comment was not created");
    const reply = reviewer.replyTo(thread, { text: "Checked.", date });

    expect(thread.date).toBe(date);
    expect(reply?.date).toBe(date);
  });
});
