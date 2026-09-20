/**
 * An edit that replaces a block's text keeps the comment ranges that block
 * anchored, and the link its text carried.
 *
 * The comment shape is the public corpus's `Comment021`: five ranges that all
 * start in the first paragraph and end in later ones. Replacing the first
 * paragraph's text used to drop every `comment` mark it carried, because the
 * mark is `inclusive: false` and the replacement reached the end of what it
 * covered. The ends and references in the later paragraphs stayed, so the save
 * wrote five `commentRangeEnd` elements with no `commentRangeStart`: invalid
 * OOXML, and five comments anchored to nothing. `hyperlink` is non-inclusive
 * for the same reason and lost the same way, leaving prose that had been a
 * link pointing nowhere.
 */

import { describe, expect, test } from "bun:test";
import JSZip from "jszip";

import { createDocx } from "../docx/rezip";
import type { Comment, Document, Paragraph, ParagraphContent } from "../types/document";
import { FolioDocxReviewer } from "./headless";

const COMMENT_IDS = [0, 1, 2, 3, 4] as const;
/** Paragraph each comment's range ends in; every range starts in paragraph 0. */
const RANGE_END_PARAGRAPH = [1, 1, 2, 3, 3] as const;
const FIRST_PARAGRAPH_TEXT = "First paragraph text.";

const run = (text: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text }],
});

const comment = (id: number): Comment => ({
  id,
  author: "Dana Lindqvist",
  initials: "DL",
  date: "2024-01-01T00:00:00Z",
  content: [{ type: "paragraph", content: [run(`Note ${id}`)] }],
});

const buildDocument = (): Document => {
  const opening: ParagraphContent[] = COMMENT_IDS.map((id) => ({
    type: "commentRangeStart",
    id,
  }));
  opening.push(run(FIRST_PARAGRAPH_TEXT));

  const content: Paragraph[] = [{ type: "paragraph", paraId: "10000001", content: opening }];
  for (let index = 1; index <= 3; index++) {
    const paragraph: ParagraphContent[] = [run(`Paragraph ${index} text.`)];
    for (const id of COMMENT_IDS) {
      if (RANGE_END_PARAGRAPH[id] === index) {
        paragraph.push({ type: "commentRangeEnd", id }, { type: "commentReference", id });
      }
    }
    content.push({ type: "paragraph", paraId: `1000000${index + 1}`, content: paragraph });
  }

  return { package: { document: { comments: COMMENT_IDS.map(comment), content } } };
};

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const zip = await JSZip.loadAsync(buffer);
  return (await zip.file("word/document.xml")?.async("text")) ?? "";
};

const markerIds = (xml: string, marker: string): number[] =>
  [...xml.matchAll(new RegExp(`<w:${marker}\\s[^>]*w:id="(\\d+)"`, "gu"))].map(([, id]) =>
    Number(id),
  );

const replaceFirstParagraph = async (mode: "direct" | "tracked-changes"): Promise<string> => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(buildDocument()), {
    author: "Editor",
  });
  const [target] = reviewer.snapshot().blocks;
  expect(target?.text).toBe(FIRST_PARAGRAPH_TEXT);
  reviewer.applyOperations(
    [
      {
        id: "r1",
        type: "replaceInBlock",
        blockId: target?.id ?? "",
        find: FIRST_PARAGRAPH_TEXT,
        replace: "Wholly different wording.",
      },
    ],
    { mode },
  );
  return documentXml(await reviewer.toBuffer());
};

describe("replaceInBlock over a block that opens comment ranges", () => {
  test.each(["direct", "tracked-changes"] as const)(
    "keeps every range's start in %s mode",
    async (mode) => {
      const xml = await replaceFirstParagraph(mode);

      expect(markerIds(xml, "commentRangeStart").toSorted()).toEqual([...COMMENT_IDS]);
      expect(markerIds(xml, "commentRangeEnd").toSorted()).toEqual([...COMMENT_IDS]);
      expect(markerIds(xml, "commentReference").toSorted()).toEqual([...COMMENT_IDS]);

      // The starts move to the beginning of the replacement, which is still the
      // first paragraph: every range covers the new text, as Word's does.
      const [firstParagraph] = xml.split("</w:p>");
      for (const id of COMMENT_IDS) {
        expect(firstParagraph).toContain(`<w:commentRangeStart w:id="${id}"/>`);
      }
      expect(xml).toContain("Wholly different wording.");
    },
  );
});

const LINK_TEXT = "the standard terms";
const LINK_TARGET = "https://example.invalid/standard-terms";
const REPLACED_LINK_TEXT = "the revised terms of engagement";

const buildLinkedDocument = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "30000001",
          content: [
            {
              type: "hyperlink",
              href: LINK_TARGET,
              tooltip: "Standard terms",
              children: [run(LINK_TEXT)],
            },
          ],
        },
      ],
    },
  },
});

describe("replaceInBlock over a block whose text is a hyperlink", () => {
  test.each(["direct", "tracked-changes"] as const)(
    "keeps the link on the new text in %s mode",
    async (mode) => {
      const reviewer = await FolioDocxReviewer.fromBuffer(await createDocx(buildLinkedDocument()), {
        author: "Editor",
      });
      const [target] = reviewer.snapshot().blocks;
      expect(target?.text).toBe(LINK_TEXT);
      reviewer.applyOperations(
        [
          {
            id: "r1",
            type: "replaceInBlock",
            blockId: target?.id ?? "",
            find: LINK_TEXT,
            replace: REPLACED_LINK_TEXT,
          },
        ],
        { mode },
      );

      const saved = await reviewer.toBuffer();
      const xml = await documentXml(saved);
      const relationships = await (async () => {
        const zip = await JSZip.loadAsync(saved);
        return (await zip.file("word/_rels/document.xml.rels")?.async("text")) ?? "";
      })();

      // The link still wraps the replacement, and still points where it did.
      const linked = [...xml.matchAll(/<w:hyperlink\b[^>]*>([\s\S]*?)<\/w:hyperlink>/gu)].map(
        ([, inside]) => inside ?? "",
      );
      expect(linked.some((inside) => inside.includes(REPLACED_LINK_TEXT))).toBe(true);
      expect(relationships).toContain(LINK_TARGET);
      // Nothing outside a link carries the new text: the replacement did not
      // escape the hyperlink it replaced.
      expect(xml.replaceAll(/<w:hyperlink\b[^>]*>[\s\S]*?<\/w:hyperlink>/gu, "")).not.toContain(
        REPLACED_LINK_TEXT,
      );
    },
  );
});
