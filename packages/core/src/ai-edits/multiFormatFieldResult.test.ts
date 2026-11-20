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
  return FolioDocxReviewer.fromBuffer((await zip.generate()).slice().buffer, {
    session: "legacy",
  });
};

const apply = (
  reviewer: FolioDocxReviewer,
  mode: "direct" | "tracked-changes" | "suggested",
  operation: Record<string, unknown>,
) =>
  reviewer.applyDocumentOperations(
    { version: 1, mode, operations: [{ id: "edit", ...operation }] } as never,
    { undefinedReferences: "refuse" },
  );

const firstText = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().at(0)?.text;

const reopen = async (reviewer: FolioDocxReviewer) =>
  FolioDocxReviewer.fromBuffer(await reviewer.toBuffer(), {
    session: "legacy",
  });

describe.each([
  ["complex", COMPLEX],
  ["simple", SIMPLE],
  ["plain-result", PLAIN],
] as const)("a %s field with a result in two formats", (_kind, field) => {
  test("reads as its result", async () => {
    expect(firstText(await open(field))).toBe(SOURCE);
  });

  describe.each(["direct", "tracked-changes", "suggested"] as const)(
    "(%s)",
    (mode) => {
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

      test.each([
        [
          "its entire result",
          { type: "replaceInBlock", find: "2014", replace: "1999" },
          ["Date: 1999 end", "Tail"],
        ],
        [
          "text across it",
          { type: "replaceInBlock", find: ": 2014 e", replace: "X" },
          ["DateXnd", "Tail"],
        ],
        ["its host paragraph", { type: "deleteBlock" }, ["Tail"]],
      ] as const)("edits %s as a unit", async (_name, operation, expected) => {
        const reviewer = await open(field);
        const original = reviewer.toDocument();
        const result = apply(reviewer, mode, { ...operation, blockId: HOST });
        if (mode === "tracked-changes") {
          expect(result.status).toBe("rejected");
          expect(result.applied).toEqual([]);
          expect(result.skipped).toEqual([
            {
              id: "edit",
              reason: "invalidResult",
              message: expect.stringMatching(/Accept All.*field/u),
            },
          ]);
          expect(reviewer.getChanges()).toEqual([]);
          expect(reviewer.toDocument()).toEqual(original);
          expect(firstText(await reopen(reviewer))).toBe(SOURCE);
        } else {
          expect(result.applied).toHaveLength(1);
          if (mode === "suggested") {
            const suggestionId = result.applied.at(0)?.suggestionId;
            if (suggestionId === undefined)
              throw new Error("Expected a field edit suggestion");
            expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
          }
          const saved = await reopen(reviewer);
          if (mode === "suggested")
            expect(saved.acceptAll().resolved).toBeGreaterThan(0);
          expect(saved.getContent().map(({ text }) => text)).toEqual(expected);
        }
      });
    },
  );

  test("comes back whole when its deletion is rejected", async () => {
    const reviewer = await open(field);
    const staged = apply(reviewer, "suggested", {
      type: "deleteBlock",
      blockId: HOST,
    });
    const suggestionId = staged.applied.at(0)?.suggestionId;
    if (suggestionId === undefined)
      throw new Error("Expected a deletion suggestion");
    expect(reviewer.acceptSuggestion(suggestionId)).toBe(true);
    const saved = await reopen(reviewer);
    expect(saved.rejectAll().resolved).toBeGreaterThan(0);
    expect(firstText(saved)).toBe(SOURCE);
    expect(firstText(await reopen(saved))).toBe(SOURCE);

    expect(reviewer.rejectAll().resolved).toBeGreaterThan(0);
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
