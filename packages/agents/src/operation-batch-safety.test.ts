/**
 * `suggest_changes` refuses what the document-operation applier refuses, and
 * says why in words a model can act on: an operation whose target an earlier
 * operation of the same call already claims, and an offset inside a character.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FolioDocxReviewer } from "@stll/folio-core/server";

import { createReviewerBridge } from "./bridges/reviewer";
import { executeFolioToolCallUntyped } from "./execute";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const W14 = "http://schemas.microsoft.com/office/word/2010/wordml";

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
        `<w:p w14:paraId="1000000${String(index)}" w14:textId="77777777">` +
        `<w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`,
    )
    .join("");
  zip.file(
    "word/document.xml",
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      `<w:document xmlns:w="${W}" xmlns:w14="${W14}"><w:body>${body}<w:sectPr/></w:body></w:document>`,
  );
  return await zip.generateAsync({ type: "arraybuffer" });
};

const acceptedTexts = async (reviewer: FolioDocxReviewer): Promise<string[]> => {
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  reopened.acceptAll();
  const accepted = await FolioDocxReviewer.fromBuffer(await reopened.toBuffer());
  return accepted.snapshot().blocks.map((block) => block.text);
};

type SuggestSummary = {
  applied: { id: string }[];
  skipped: { id: string; reason: string }[];
  issues: { operationId: string; code: string; message?: string }[];
};

const suggest = (reviewer: FolioDocxReviewer, operations: unknown[]): SuggestSummary => {
  const call = executeFolioToolCallUntyped(
    "suggest_changes",
    { operations },
    createReviewerBridge(reviewer, { mode: "tracked-changes" }),
    {},
  );
  if (!call.ok) {
    throw new Error(`suggest_changes failed: ${JSON.stringify(call)}`);
  }
  // SAFETY: a successful suggest_changes call returns the apply summary.
  return call.result as SuggestSummary;
};

describe("suggest_changes with two operations on one block", () => {
  test("refuses the deletion that would have rewritten the next paragraph", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await buildDocx(["Alpha beta gamma.", "Second clause.", "Third clause."]),
      { author: "QA" },
    );
    const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
    const summary = suggest(reviewer, [
      { id: "replace", type: "replaceInBlock", blockId, find: "Alpha", replace: "First" },
      { id: "delete", type: "deleteBlock", blockId },
    ]);

    expect(summary.applied.map(({ id }) => id)).toEqual(["replace"]);
    expect(summary.skipped).toHaveLength(1);
    expect(summary.skipped[0]?.id).toBe("delete");
    expect(summary.skipped[0]?.reason).toContain("earlier operation in this batch");
    expect(summary.skipped[0]?.reason).toContain('"replace"');
    expect(summary.issues.map(({ code }) => code)).toEqual(["overlappingOperation"]);
    expect(await acceptedTexts(reviewer)).toEqual([
      "First beta gamma.",
      "Second clause.",
      "Third clause.",
    ]);
  });
});

describe("suggest_changes splitting inside an emoji", () => {
  test("refuses the offset and names the boundaries either side", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx(["Party 🧑 agrees."]), {
      author: "QA",
    });
    const blockId = reviewer.snapshot().blocks[0]?.id ?? "";
    const summary = suggest(reviewer, [{ id: "split", type: "splitBlock", blockId, offset: 7 }]);

    expect(summary.applied).toEqual([]);
    expect(summary.skipped[0]?.reason).toContain("inside a single character");
    expect(summary.skipped[0]?.reason).toContain("use 6 or 8");
    expect(await acceptedTexts(reviewer)).toEqual(["Party 🧑 agrees."]);
  });
});
