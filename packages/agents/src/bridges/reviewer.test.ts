import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { executeFolioToolCall } from "../execute";
import { FOLIO_AGENT_TOOL_NAMES } from "../types";
import { createReviewerBridge } from "./reviewer";

const buildDocx = async (text: string): Promise<ArrayBuffer> => {
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
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
      'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body>' +
      `<w:p w14:paraId="1A2B3C4D" w14:textId="77777777"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>` +
      "<w:sectPr/></w:body></w:document>",
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

describe("createReviewerBridge revisionStamp", () => {
  test("stamps revisions, comments and replies from one fixed provenance", async () => {
    const revisionStamp = { date: "2026-01-02T03:04:05Z", idSeed: 7000 };
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx("Pay $50 now."), {
      author: "Reviewer",
    });
    const bridge = createReviewerBridge(reviewer, { revisionStamp });

    const suggested = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.suggestChanges,
      {
        operations: [{ type: "replaceInBlock", blockId: "1A2B3C4D", find: "$50", replace: "$500" }],
      },
      bridge,
    );
    const commented = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.addComment,
      { blockId: "1A2B3C4D", text: "Confirm the amount." },
      bridge,
    );
    const thread = bridge.getComments().at(0);
    if (!thread) throw new Error("comment was not created");
    const replied = executeFolioToolCall(
      FOLIO_AGENT_TOOL_NAMES.replyComment,
      { commentId: thread.id, text: "Confirmed." },
      bridge,
    );

    expect([suggested.ok, commented.ok, replied.ok]).toEqual([true, true, true]);
    const changes = reviewer.getChanges();
    expect(changes.length).toBeGreaterThan(0);
    for (const change of changes) {
      expect(change.date).toBe(revisionStamp.date);
      expect(change.id).toBeGreaterThanOrEqual(revisionStamp.idSeed);
    }
    const [comment] = reviewer.getComments();
    expect(comment?.date).toBe(revisionStamp.date);
    expect(comment?.replies.map(({ date }) => date)).toEqual([revisionStamp.date]);
  });
});
