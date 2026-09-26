/**
 * `"suggested"` edits stay out of a saved package until accepted. Accepting
 * them all on the headless reviewer makes them content the save keeps,
 * whole inserted paragraphs included; rejecting them all removes them.
 */

import { describe, expect, test } from "bun:test";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const buildDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [
    paragraph("First clause."),
    paragraph("Signed in two copies."),
  ];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer;
};

const suggest = async (): Promise<FolioDocxReviewer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
  const anchor = reviewer.getContent().find(({ text }) => text === "Signed in two copies.");
  if (!anchor) {
    throw new Error("the fixture paragraph is missing");
  }
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "suggested",
    operations: [
      { id: "insert", type: "insertAfterBlock", blockId: anchor.id, text: "Suggested clause." },
      { id: "replace", type: "replaceInBlock", blockId: anchor.id, find: "two", replace: "three" },
    ],
  });
  expect(result.applied).toHaveLength(2);
  return reviewer;
};

const savedTexts = async (reviewer: FolioDocxReviewer): Promise<string[]> =>
  (await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer()))
    .getContent()
    .map(({ text }) => text);

describe("resolving every suggestion headlessly", () => {
  test("a save before resolving keeps the document as it was", async () => {
    expect(await savedTexts(await suggest())).toEqual(["First clause.", "Signed in two copies."]);
  });

  test("a suggested block deletion leaves no tracked change in the saved package", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), { author: "AI" });
    const target = reviewer.getContent().find(({ text }) => text === "First clause.");
    if (!target) {
      throw new Error("the fixture paragraph is missing");
    }
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "suggested",
      operations: [{ id: "delete", type: "deleteBlock", blockId: target.id }],
    });
    expect(result.applied).toHaveLength(1);
    const saved = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    expect(saved.getChanges()).toEqual([]);
    expect(saved.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
  });

  test("acceptAll keeps an inserted paragraph through the save", async () => {
    const reviewer = await suggest();
    reviewer.acceptAll();
    expect(reviewer.getChanges()).toEqual([]);
    expect(await savedTexts(reviewer)).toEqual([
      "First clause.",
      "Signed in three copies.",
      "Suggested clause.",
    ]);
  });

  test("acceptChange on a suggested paragraph insert keeps the paragraph through the save", async () => {
    const reviewer = await suggest();
    const insertion = reviewer
      .getChanges()
      .find(({ type, text }) => type === "insertion" && text === "Suggested clause.");
    if (!insertion) {
      throw new Error("the suggested insertion is not listed");
    }
    expect(reviewer.acceptChange(insertion)).toBe(true);
    expect(await savedTexts(reviewer)).toContain("Suggested clause.");
  });

  test("rejectChange on a suggested paragraph insert removes the paragraph", async () => {
    const reviewer = await suggest();
    const insertion = reviewer
      .getChanges()
      .find(({ type, text }) => type === "insertion" && text === "Suggested clause.");
    if (!insertion) {
      throw new Error("the suggested insertion is not listed");
    }
    expect(reviewer.rejectChange(insertion)).toBe(true);
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain("");
    expect(reviewer.getContent().map(({ text }) => text)).not.toContain("Suggested clause.");
  });

  test("rejectAll removes every suggestion", async () => {
    const reviewer = await suggest();
    reviewer.rejectAll();
    expect(reviewer.getContent().map(({ text }) => text)).toEqual([
      "First clause.",
      "Signed in two copies.",
    ]);
    expect(await savedTexts(reviewer)).toEqual(["First clause.", "Signed in two copies."]);
  });
});
