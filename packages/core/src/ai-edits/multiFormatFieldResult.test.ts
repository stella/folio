/**
 * A field whose result is in more than one format is one thing to an edit, like
 * a field with a plain result: replaced or removed whole, never cut into.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import { hashFolioAIBlockText } from "./snapshot";

const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"';

const COMPLEX = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>20</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>14</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const PLAIN = `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve"> DATE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>2014</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>`;
const SIMPLE = `<w:fldSimple w:instr=" DATE "><w:r><w:rPr><w:b/></w:rPr><w:t>20</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>14</w:t></w:r></w:fldSimple>`;

const HOST = "10000001";
const SOURCE = "Date: 2014 end";

const open = async (field: string) => {
  const zip = await ZipArchive.read(await createDocx(createEmptyDocument()));
  zip.set(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document ${NS}><w:body><w:p w14:paraId="${HOST}"><w:r><w:t xml:space="preserve">Date: </w:t></w:r>${field}<w:r><w:t xml:space="preserve"> end</w:t></w:r></w:p><w:p w14:paraId="10000002"><w:r><w:t>Tail</w:t></w:r></w:p></w:body></w:document>`,
  );
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer);
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes",
  operation: Record<string, unknown>,
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );

const firstText = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().at(0)?.text;

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());

describe.each([
  ["complex", COMPLEX],
  ["simple", SIMPLE],
  ["plain-result", PLAIN],
] as const)("a %s field with a result in two formats", (_kind, field) => {
  test("reads as its result", async () => {
    expect(firstText(await open(field))).toBe(SOURCE);
  });

  describe.each(["direct", "tracked-changes"] as const)("(%s)", (mode) => {
    test("refuses an edit that cuts into its result", async () => {
      const reviewer = await open(field);
      const result = apply(reviewer, mode, {
        type: "replaceInBlock",
        blockId: HOST,
        find: "20",
        replace: "30",
      });
      expect(result.applied).toHaveLength(0);
      expect(firstText(await reopen(reviewer))).toBe(SOURCE);
    });

    test("is replaced whole by an edit of all of its result", async () => {
      const reviewer = await open(field);
      const result = apply(reviewer, mode, {
        type: "replaceInBlock",
        blockId: HOST,
        find: "2014",
        replace: "1999",
      });
      expect(result.applied).toHaveLength(1);
      const saved = await reopen(reviewer);
      saved.acceptAll();
      expect(firstText(saved)).toBe("Date: 1999 end");
    });

    test("goes whole with an edit across it", async () => {
      const reviewer = await open(field);
      const result = apply(reviewer, mode, {
        type: "replaceInBlock",
        blockId: HOST,
        find: ": 2014 e",
        replace: "X",
      });
      expect(result.applied).toHaveLength(1);
      const saved = await reopen(reviewer);
      saved.acceptAll();
      expect(firstText(saved)).toBe("DateXnd");
    });

    test("goes with its paragraph", async () => {
      const reviewer = await open(field);
      const result = apply(reviewer, mode, {
        type: "deleteBlock",
        blockId: HOST,
      });
      expect(result.applied).toHaveLength(1);
      const saved = await reopen(reviewer);
      saved.acceptAll();
      expect(saved.getContent().map(({ text }) => text)).toEqual(["Tail"]);
    });
  });

  test("comes back whole when its deletion is rejected", async () => {
    const reviewer = await open(field);
    apply(reviewer, "tracked-changes", { type: "deleteBlock", blockId: HOST });
    const saved = await reopen(reviewer);
    saved.rejectAll();
    expect(firstText(saved)).toBe(SOURCE);
    expect(firstText(await reopen(saved))).toBe(SOURCE);

    reviewer.rejectAll();
    expect(firstText(reviewer)).toBe(SOURCE);
    expect(firstText(await reopen(reviewer))).toBe(SOURCE);
  });

  test("keeps quoting its result in a comment over it after a reopen", async () => {
    const reviewer = await open(field);
    const text = firstText(reviewer) ?? "";
    apply(reviewer, "direct", {
      type: "commentOnRange",
      range: {
        type: "textRange",
        story: "main",
        blockId: HOST,
        startOffset: text.indexOf("2014"),
        endOffset: text.indexOf("2014") + 4,
        selectedTextHash: hashFolioAIBlockText("2014"),
      },
      comment: { text: "Check the date" },
    });
    expect(reviewer.getComments().at(0)?.anchoredText).toBe("2014");
    const saved = await reopen(reviewer);
    expect(saved.getComments().at(0)?.anchoredText).toBe("2014");
  });

  test("takes a comment over it", async () => {
    const reviewer = await open(field);
    const content = reviewer.getContent().at(0);
    if (!content) throw new Error("The paragraph is missing");
    const result = apply(reviewer, "direct", {
      type: "commentOnBlock",
      blockId: HOST,
      comment: { text: "Check the date" },
    });
    expect(result.applied).toHaveLength(1);
    const saved = await reopen(reviewer);
    expect(firstText(saved)).toBe(SOURCE);
    expect(saved.getComments()).toHaveLength(1);
  });
});
