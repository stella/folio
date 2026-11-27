/**
 * Resolving a story to its final view and saving, in a package that carries no
 * `w14:paraId`.
 *
 * folio mints an id for every paragraph that arrives without one, keyed on
 * the package and the paragraph's place in it, so a reopen of the same bytes
 * mints it again and a save need not write it. A save writes them only into a
 * part it rewrites. Resolving the
 * story changes the package, so a reopened copy keys its ids on other content.
 * The check that the resolved story reached the package therefore cannot
 * depend on a minted id: it reported a story that persisted perfectly as one
 * that had not.
 */

import { describe, expect, test } from "bun:test";

import { ZipArchive } from "@stll/docx-core/zip";

import { createDocx } from "../docx/rezip";
import { endnote, run } from "../docx/server/build";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type { BlockContent, Document } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";

const partText = async (buffer: ArrayBuffer, part: string): Promise<string> =>
  (await (await ZipArchive.read(buffer)).entry(part)?.text()) ?? "";

/** One paragraph holding a deletion, an insertion and a relocation. */
const reviewedParagraph = (paraId?: string): BlockContent => ({
  type: "paragraph",
  ...(paraId === undefined ? {} : { paraId, textId: paraId }),
  content: [
    run("kept "),
    {
      type: "deletion",
      info: { id: 1, author: "Alice" },
      content: [run("deleted ")],
    },
    {
      type: "insertion",
      info: { id: 2, author: "Alice" },
      content: [run("inserted ")],
    },
    {
      type: "moveFrom",
      info: { id: 3, author: "Alice" },
      content: [run("old-position ")],
    },
    {
      type: "moveTo",
      info: { id: 3, author: "Alice" },
      content: [run("new-position")],
    },
  ],
});

const FINAL_TEXT = "kept inserted new-position";

const AUTHORED_PARA_ID = "1A2B3C4D";

/**
 * A package of `compose`'s blocks. A save gives every paragraph an id; the
 * ones it gave are taken off again, and an authored one stays.
 */
const buildDocx = async (
  compose: (document: Document) => BlockContent[],
): Promise<ArrayBuffer> => {
  const document = createEmptyDocument();
  document.package.document.content = compose(document);
  const zip = await ZipArchive.read(await createDocx(document));
  for (const path of zip
    .paths()
    .filter((name) => /^word\/[^/]+\.xml$/u.test(name))) {
    // oxlint-disable-next-line no-await-in-loop -- each part is read and rewritten in turn
    const xml = await zip.entry(path)?.text();
    if (xml !== undefined)
      zip.set(
        path,
        xml.replace(
          new RegExp(
            `\\s(?:w14:paraId|w14:textId)="(?!${AUTHORED_PARA_ID}")[^"]*"`,
            "gu",
          ),
          "",
        ),
      );
  }
  return (await zip.generate({ compression: "store" })).buffer as ArrayBuffer;
};

const resolveEveryStoryToFinal = async (
  source: ArrayBuffer,
  session: HeadlessReviewerSession,
): Promise<ArrayBuffer> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(source, { session });
  for (const { handle } of reviewer.listStories()) {
    expect(
      reviewer.resolveReviewedStory({ story: handle, view: "final" }),
    ).toBe(true);
  }
  return await reviewer.toBuffer();
};

describe.each([
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[])(
  "a resolved story in a package with no authored paragraph ids (%s)",
  (session) => {
    test("persists without asking the save to write an id it need not write", async () => {
      const source = await buildDocx(() => [reviewedParagraph()]);
      expect(await partText(source, "word/document.xml")).not.toContain(
        "w14:paraId",
      );

      const saved = await resolveEveryStoryToFinal(source, session);

      const xml = await partText(saved, "word/document.xml");
      expect(xml).toContain(FINAL_TEXT);
      expect(xml).not.toContain("deleted");
      expect(xml).not.toContain("old-position");
      expect(xml).not.toMatch(/<w:(?:ins|del|moveFrom|moveTo)\b/u);

      const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
      expect(
        reopened.readReviewedStory({ view: "current-markup" })?.changes,
      ).toEqual([]);
      expect(reopened.readStory({ type: "main" })?.text).toEndWith(FINAL_TEXT);
    });

    test("persists a secondary story the same way", async () => {
      const source = await buildDocx((document) => [
        {
          type: "paragraph",
          content: [run("Body"), endnote(document, [reviewedParagraph()])],
        },
      ]);
      expect(await partText(source, "word/endnotes.xml")).not.toContain(
        "w14:paraId",
      );

      const saved = await resolveEveryStoryToFinal(source, session);

      const reopened = await FolioDocxReviewer.fromBuffer(saved, { session });
      const note = reopened
        .listStories()
        .find(({ handle }) => handle.type === "endnote");
      expect(note).toMatchObject({ text: FINAL_TEXT });
      expect(
        note
          ? reopened.readReviewedStory({
              story: note.handle,
              view: "current-markup",
            })?.changes
          : null,
      ).toEqual([]);
    });

    test("still compares an authored id, which the save does write", async () => {
      const source = await buildDocx(() => [
        reviewedParagraph(AUTHORED_PARA_ID),
      ]);

      const xml = await partText(
        await resolveEveryStoryToFinal(source, session),
        "word/document.xml",
      );

      expect(xml).toContain(FINAL_TEXT);
      expect(xml).toContain('w14:paraId="1A2B3C4D"');
    });
  },
);
