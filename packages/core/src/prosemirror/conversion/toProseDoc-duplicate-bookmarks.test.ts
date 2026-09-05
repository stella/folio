/**
 * Regression: a `.docx` whose bookmark boundaries do not pair cleanly must
 * still load.
 *
 * Real contract templates carry untidy bookmarks — an id repeated across
 * paragraphs, a start Word never closed, an end whose start was edited away.
 * Word and LibreOffice open those files; conversion used to reject the whole
 * document, so nothing downstream (the editor, the reviewer, `compareDocx`)
 * could touch them.
 *
 * The document here is built from the typed model rather than copied from any
 * real file, so the fixture carries the shape and none of the content.
 */

import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../../ai-edits/headless";
import { compareDocx } from "../../compare/compare";
import { createDocx } from "../../docx/rezip";
import { parseDocx } from "../../docx/parser";
import type { Paragraph } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { toProseDoc } from "./toProseDoc";

type ParagraphSpec = {
  paraId: string;
  text: string;
  content?: Paragraph["content"];
};

const buildDocxBuffer = (paragraphs: readonly ParagraphSpec[]): Promise<ArrayBuffer> => {
  const template = createEmptyDocument();
  return createDocx({
    ...template,
    package: {
      ...template.package,
      document: {
        ...template.package.document,
        content: paragraphs.map(({ paraId, text, content }) => ({
          type: "paragraph",
          paraId,
          content: [...(content ?? []), { type: "run", content: [{ type: "text", text }] }],
        })),
      },
    },
  });
};

/**
 * The shape that used to fail: bookmark id 0 started in two different
 * paragraphs, plus a start with no end and an end with no start.
 */
const buildUntidyBookmarkDocx = (): Promise<ArrayBuffer> =>
  buildDocxBuffer([
    {
      paraId: "C0000001",
      text: "First clause.",
      content: [{ type: "bookmarkStart", id: 0, name: "_GoBack" }],
    },
    {
      paraId: "C0000002",
      text: "Second clause.",
      content: [{ type: "bookmarkStart", id: 0, name: "_GoBack" }],
    },
    {
      paraId: "C0000003",
      text: "Third clause.",
      content: [{ type: "bookmarkStart", id: 7, name: "_Unclosed" }],
    },
    {
      paraId: "C0000004",
      text: "Fourth clause.",
      content: [{ type: "bookmarkEnd", id: 99 }],
    },
  ]);

const BLOCK_TEXTS = ["First clause.", "Second clause.", "Third clause.", "Fourth clause."] as const;

describe("bookmark boundaries that do not pair", () => {
  test("toProseDoc converts a document with a duplicate bookmark id", async () => {
    const document = await parseDocx(await buildUntidyBookmarkDocx(), {
      detectVariables: false,
      preloadFonts: false,
    });

    const doc = toProseDoc(document);

    expect(doc.textContent).toContain("First clause.");
    expect(doc.textContent).toContain("Second clause.");
  });

  test("the reviewer loads it and keeps every block", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildUntidyBookmarkDocx());

    expect(reviewer.getContent().map((block) => block.text)).toEqual([...BLOCK_TEXTS]);
  });

  test("it survives a save and reload", async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(await buildUntidyBookmarkDocx());
    const reloaded = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());

    expect(reloaded.getContent().map((block) => block.text)).toEqual([...BLOCK_TEXTS]);
  });

  test("compareDocx reports the edit instead of failing to parse", async () => {
    const base = await buildUntidyBookmarkDocx();
    const target = await buildDocxBuffer([
      {
        paraId: "C0000001",
        text: "First clause, amended.",
        content: [{ type: "bookmarkStart", id: 0, name: "_GoBack" }],
      },
      {
        paraId: "C0000002",
        text: "Second clause.",
        content: [{ type: "bookmarkStart", id: 0, name: "_GoBack" }],
      },
      {
        paraId: "C0000003",
        text: "Third clause.",
        content: [{ type: "bookmarkStart", id: 7, name: "_Unclosed" }],
      },
      { paraId: "C0000004", text: "Fourth clause.", content: [{ type: "bookmarkEnd", id: 99 }] },
    ]);

    const result = await compareDocx(base, target, {
      author: "compare",
      timestamp: "2024-03-01T00:00:00.000Z",
    });

    expect(result.isOk()).toBe(true);
    if (!result.isOk()) {
      return;
    }
    expect(result.value.changes.map((change) => change.kind)).toEqual(["replace"]);
  });
});
