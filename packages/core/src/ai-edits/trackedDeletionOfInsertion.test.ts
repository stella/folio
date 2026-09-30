/**
 * A tracked deletion over text that is itself a pending tracked insertion (a
 * reviewer revising a suggestion) saves as `w:ins > w:del` and reads back as
 * it was left: the deleted inserted text stays deleted.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { ensureParaIds } from "../docx/ensureParaIds";
import { createDocx } from "../docx/rezip";
import { paragraph } from "../docx/server/build";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const TEXT = "Signed in two copies.";

const buildDocument = async (): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = [paragraph(TEXT)];
  const { docx } = await ensureParaIds(new Uint8Array(await createDocx(document)));
  return docx.buffer.slice(docx.byteOffset, docx.byteOffset + docx.byteLength) as ArrayBuffer;
};

const replace = (reviewer: FolioDocxReviewer, find: string, replacement: string): void => {
  const block = reviewer.getContent()[0];
  if (!block) {
    throw new Error("the fixture paragraph is missing");
  }
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "tracked-changes",
    operations: [
      { id: find, type: "replaceInBlock", blockId: block.id, find, replace: replacement },
    ],
  });
  expect(result.applied).toHaveLength(1);
};

const changesOf = (reviewer: FolioDocxReviewer) =>
  reviewer.getChanges().map(({ type, text, author }) => ({ type, text, author }));

describe("a tracked deletion of pending inserted text", () => {
  for (const secondAuthor of ["Reviewer", "Second Reviewer"]) {
    test(`survives a save (${secondAuthor === "Reviewer" ? "same" : "another"} author)`, async () => {
      const first = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
        author: "Reviewer",
      });
      replace(first, "two", "three");
      const second = await FolioDocxReviewer.fromBuffer(await first.toBuffer(), {
        author: secondAuthor,
      });
      replace(second, "three", "four");
      expect(second.getContent()[0]?.text).toBe("Signed in four copies.");

      const saved = await second.toBuffer();
      const xml = await (await JSZip.loadAsync(saved)).file("word/document.xml")?.async("string");
      expect(xml).toMatch(
        /<w:ins [^>]*><w:del [^>]*><w:r>(?:<w:rPr\/>)?<w:delText>three<\/w:delText>/u,
      );

      const reopened = await FolioDocxReviewer.fromBuffer(saved);
      expect(reopened.getContent()[0]?.text).toBe("Signed in four copies.");
      expect(changesOf(reopened)).toContainEqual({
        type: "deletion",
        text: "three",
        author: secondAuthor,
      });

      reopened.rejectAll();
      expect(reopened.getContent()[0]?.text).toBe(TEXT);
      const accepted = await FolioDocxReviewer.fromBuffer(saved);
      accepted.acceptAll();
      expect(accepted.getContent()[0]?.text).toBe("Signed in four copies.");
    });
  }
});

describe("revision wrapper enumeration", () => {
  test("keeps deletion fragments around a pending insertion across save", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocument(), {
      author: "Reviewer",
    });
    replace(reviewer, "two", "three");
    const block = reviewer.getContent().at(0);
    if (!block) throw new Error("the fixture paragraph is missing");
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [{ id: "delete", type: "deleteBlock", blockId: block.id }],
    });
    expect(result.applied).toHaveLength(1);
    const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
    const sorted = (target: FolioDocxReviewer) =>
      changesOf(target).toSorted((a, b) => JSON.stringify(a).localeCompare(JSON.stringify(b)));
    expect(sorted(reviewer)).toEqual(sorted(reopened));
  });
});
