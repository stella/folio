/**
 * Deleting the paragraph that ends a section.
 *
 * A section's properties live on the mark of its last paragraph (ECMA-376
 * Part 1 §17.6.18). Deleting that paragraph deletes the break, and the
 * content before it becomes part of the following section, whose properties
 * it takes, which is also what accepting the same deletion tracked already
 * did. The direct edit used to report success and leave a reviewer
 * whose save threw, because the save saw a section vanish that no edit had
 * accounted for.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { buildSectionsDocx, SECTIONS_FIXTURE_TEXT } from "./__fixtures__/sections";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditApplyMode } from "./types";

const open = async (bytes: ArrayBuffer): Promise<FolioDocxReviewer> =>
  await FolioDocxReviewer.fromBuffer(bytes, { author: "Reviewer" });

const blockId = (reviewer: FolioDocxReviewer, text: string): string => {
  const block = reviewer.getContent().find((candidate) => candidate.text === text);
  if (!block) throw new Error(`no block "${text}"`);
  return block.id;
};

const deleteBlock = (reviewer: FolioDocxReviewer, text: string, mode: FolioAIEditApplyMode) =>
  reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: [{ id: "delete", type: "deleteBlock", blockId: blockId(reviewer, text) }],
  });

const documentXml = async (bytes: ArrayBuffer): Promise<string> => {
  const xml = await (await JSZip.loadAsync(bytes)).file("word/document.xml")?.async("string");
  if (xml === undefined) throw new Error("no document part");
  return xml;
};

/** Each top-level paragraph's text and whether it ends a section, in order. */
const sectionsOf = (reviewer: FolioDocxReviewer): { text: string; endsSection: boolean }[] =>
  reviewer.getContent().flatMap((block) => {
    if (block.table) return [];
    const paragraph = reviewer
      .toDocument()
      .package.document.content.find(
        (candidate) => candidate.type === "paragraph" && candidate.paraId === block.id,
      );
    return [
      {
        text: block.text,
        endsSection: paragraph?.type === "paragraph" && paragraph.sectionProperties !== undefined,
      },
    ];
  });

describe("deleting the paragraph that ends a section", () => {
  test("merges its section into the next one and saves, in direct mode", async () => {
    const reviewer = await open(await buildSectionsDocx());

    const result = deleteBlock(reviewer, SECTIONS_FIXTURE_TEXT.oneCloses, "direct");
    expect(result.applied).toHaveLength(1);
    expect(result.issues).toEqual([]);

    const saved = await reviewer.save();
    expect(saved.type).toBe("full-repack");
    if (saved.type === "repackRefused") throw new Error("refused");
    const xml = await documentXml(saved.buffer);
    // Two section records left: the landscape one, and the body's own.
    expect(xml.match(/<w:sectPr\b/gu)).toHaveLength(2);
    // The merged section's own header references went with its record.
    expect(xml).not.toContain("rIdHeaderOneDefault");

    const reopened = await open(saved.buffer);
    expect(sectionsOf(reopened)).toEqual([
      { text: SECTIONS_FIXTURE_TEXT.oneOpens, endsSection: false },
      { text: SECTIONS_FIXTURE_TEXT.twoOpens, endsSection: false },
      { text: SECTIONS_FIXTURE_TEXT.twoCloses, endsSection: true },
      { text: SECTIONS_FIXTURE_TEXT.threeOpens, endsSection: false },
      { text: SECTIONS_FIXTURE_TEXT.threeCloses, endsSection: false },
    ]);
  });

  test("reaches the same document as accepting the tracked and suggested deletion", async () => {
    const source = await buildSectionsDocx();
    const direct = await open(source);
    deleteBlock(direct, SECTIONS_FIXTURE_TEXT.oneCloses, "direct");
    const expected = sectionsOf(await open(await direct.toBuffer()));

    for (const mode of ["tracked-changes", "suggested"] as const) {
      const reviewer = await open(source);
      expect(deleteBlock(reviewer, SECTIONS_FIXTURE_TEXT.oneCloses, mode).applied).toHaveLength(1);
      // Pending, the section is still there, and the package still saves.
      expect(sectionsOf(await open(await reviewer.toBuffer()))).toContainEqual({
        text: mode === "suggested" ? SECTIONS_FIXTURE_TEXT.oneCloses : "",
        endsSection: true,
      });
      if (mode === "suggested") expect(reviewer.acceptSuggestion("delete")).toBe(true);
      reviewer.acceptAll();
      expect(sectionsOf(await open(await reviewer.toBuffer()))).toEqual(expected);
    }
  });

  test("keeps saving after a later edit moves the remaining section breaks", async () => {
    const reviewer = await open(await buildSectionsDocx());
    deleteBlock(reviewer, SECTIONS_FIXTURE_TEXT.oneCloses, "direct");

    const later = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [
        {
          id: "insert",
          type: "insertBeforeBlock",
          blockId: blockId(reviewer, SECTIONS_FIXTURE_TEXT.oneOpens),
          text: "Lead.",
        },
      ],
    });
    expect(later.applied).toHaveLength(1);

    const reopened = await open(await reviewer.toBuffer());
    expect(sectionsOf(reopened).map(({ text }) => text)).toContain("Lead.");
  });

  test("restores the section when the deletion is undone", async () => {
    const reviewer = await open(await buildSectionsDocx());
    const result = deleteBlock(reviewer, SECTIONS_FIXTURE_TEXT.oneCloses, "direct");
    expect(result.undoHandle).not.toBeNull();
    if (!result.undoHandle) return;

    expect(reviewer.undoDocumentOperations(result.undoHandle).status).toBe("undone");

    const xml = await documentXml(await reviewer.toBuffer());
    expect(xml.match(/<w:sectPr\b/gu)).toHaveLength(3);
  });
});

describe("merging across the paragraph that ends a section", () => {
  const merge = async (text: string) => {
    const reviewer = await open(await buildSectionsDocx());
    const result = reviewer.applyDocumentOperations({
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "direct",
      operations: [{ id: "merge", type: "mergeBlockWithNext", blockId: blockId(reviewer, text) }],
    });
    expect(result.applied).toHaveLength(1);
    return sectionsOf(await open(await reviewer.toBuffer()));
  };

  test("into it: the joined paragraph still ends the section", async () => {
    expect((await merge(SECTIONS_FIXTURE_TEXT.oneOpens)).slice(0, 2)).toEqual([
      {
        text: SECTIONS_FIXTURE_TEXT.oneOpens + SECTIONS_FIXTURE_TEXT.oneCloses,
        endsSection: true,
      },
      { text: SECTIONS_FIXTURE_TEXT.twoOpens, endsSection: false },
    ]);
  });

  test("out of it: the break goes, and the section runs on into the next", async () => {
    expect((await merge(SECTIONS_FIXTURE_TEXT.oneCloses)).slice(0, 2)).toEqual([
      { text: SECTIONS_FIXTURE_TEXT.oneOpens, endsSection: false },
      {
        text: SECTIONS_FIXTURE_TEXT.oneCloses + SECTIONS_FIXTURE_TEXT.twoOpens,
        endsSection: false,
      },
    ]);
  });
});
