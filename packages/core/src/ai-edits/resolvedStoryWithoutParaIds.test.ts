/**
 * Resolving a story to its final view and saving, in a package that carries no
 * `w14:paraId`.
 *
 * folio mints an id for every paragraph that arrives without one, and the save
 * deliberately writes no minted id back: that is what keeps a package from
 * LibreOffice, Google Docs or python-docx id-less, and an edit to one local.
 * A minted id is therefore a different value on every parse, so the check that
 * the resolved story reached the package cannot be a check on that id — it
 * reported a story that persisted perfectly as one that had not.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createDocx } from "../docx/rezip";
import { endnote, run } from "../docx/server/build";
import type { BlockContent, Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const partText = async (buffer: ArrayBuffer, part: string): Promise<string> =>
  (await (await JSZip.loadAsync(buffer)).file(part)?.async("string")) ?? "";

/** One paragraph holding a deletion, an insertion and a relocation. */
const reviewedParagraph = (paraId?: string): BlockContent => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId, textId: paraId }),
  content: [
    run("kept "),
    { type: "deletion", info: { id: 1, author: "Alice" }, content: [run("deleted ")] },
    { type: "insertion", info: { id: 2, author: "Alice" }, content: [run("inserted ")] },
    { type: "moveFrom", info: { id: 3, author: "Alice" }, content: [run("old-position ")] },
    { type: "moveTo", info: { id: 3, author: "Alice" }, content: [run("new-position")] },
  ],
});

const FINAL_TEXT = "kept inserted new-position";

const buildDocx = async (compose: (document: Document) => BlockContent[]): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = compose(document);
  return await createDocx(document);
};

const resolveEveryStoryToFinal = async (source: ArrayBuffer): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source);
  for (const { handle } of reviewer.listStories()) {
    expect(reviewer.resolveReviewedStory({ story: handle, view: "final" })).toBe(true);
  }
  return await reviewer.toBuffer();
};

describe("a resolved story in a package with no authored paragraph ids", () => {
  test("persists without asking the save to write an id it will not write", async () => {
    const source = await buildDocx(() => [reviewedParagraph()]);
    expect(await partText(source, "word/document.xml")).not.toContain("w14:paraId");

    const saved = await resolveEveryStoryToFinal(source);

    const xml = await partText(saved, "word/document.xml");
    expect(xml).toContain(FINAL_TEXT);
    expect(xml).not.toContain("deleted");
    expect(xml).not.toContain("old-position");
    expect(xml).not.toMatch(/<w:(?:ins|del|moveFrom|moveTo)\b/u);
    // The package stays the author's: a minted id is not an authored one, and
    // writing it here would upgrade an id-less document on every edit.
    expect(xml).not.toContain("w14:paraId");

    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    expect(reopened.readReviewedStory({ view: "current-markup" })?.changes).toEqual([]);
    expect(reopened.readStory({ type: "main" })?.text).toEndWith(FINAL_TEXT);
  });

  test("persists a secondary story the same way", async () => {
    const source = await buildDocx((document) => [
      { type: "paragraph", content: [run("Body"), endnote(document, [reviewedParagraph()])] },
    ]);
    expect(await partText(source, "word/endnotes.xml")).not.toContain("w14:paraId");

    const saved = await resolveEveryStoryToFinal(source);

    const reopened = await FolioDocxReviewer.fromBuffer(saved);
    const note = reopened.listStories().find(({ handle }) => handle.type === "endnote");
    expect(note).toMatchObject({ text: FINAL_TEXT });
    expect(
      note
        ? reopened.readReviewedStory({ story: note.handle, view: "current-markup" })?.changes
        : null,
    ).toEqual([]);
  });

  test("still compares an authored id, which the save does write", async () => {
    const source = await buildDocx(() => [reviewedParagraph("1A2B3C4D")]);

    const xml = await partText(await resolveEveryStoryToFinal(source), "word/document.xml");

    expect(xml).toContain(FINAL_TEXT);
    expect(xml).toContain('w14:paraId="1A2B3C4D"');
  });
});
