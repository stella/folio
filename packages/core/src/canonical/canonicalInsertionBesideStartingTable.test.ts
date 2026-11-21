/**
 * A paragraph inserted before a table that starts its container, or between
 * two tables, has no paragraph on that side to name: it goes right beside
 * the table, on both engines, direct and tracked.
 */
import { expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { createDocx } from "../docx/rezip";
import { fromMarkdown } from "../markdown/fromMarkdown";

const SESSIONS = ["legacy", "canonical"] as const;
const MODES = ["direct", "tracked-changes"] as const;

const source = async (markdown: string) =>
  await createDocx(fromMarkdown(markdown));

const texts = (reviewer: FolioDocxReviewer) =>
  reviewer.getContent().map(({ text }) => text);

const blockId = (reviewer: FolioDocxReviewer, text: string) => {
  const block = reviewer.snapshot().blocks.find((item) => item.text === text);
  if (block === undefined) throw new Error(`Missing the block "${text}"`);
  return block.id;
};

const CASES = [
  {
    name: "before a table that starts the body",
    markdown: "| H |\n|---|\n| A |\n\nAfter.",
    anchor: "A",
    type: "insertBeforeBlock",
    original: ["H", "A", "After."],
    inserted: ["Heading", "H", "A", "After."],
  },
  {
    name: "after a table another table follows",
    markdown: "| H |\n|---|\n| A |\n\n| I |\n|---|\n| B |\n\nAfter.",
    anchor: "H",
    type: "insertAfterBlock",
    original: ["H", "A", "I", "B", "After."],
    inserted: ["H", "A", "Heading", "I", "B", "After."],
  },
  {
    name: "before a table another table precedes",
    markdown: "| H |\n|---|\n| A |\n\n| I |\n|---|\n| B |\n\nAfter.",
    anchor: "B",
    type: "insertBeforeBlock",
    original: ["H", "A", "I", "B", "After."],
    inserted: ["H", "A", "Heading", "I", "B", "After."],
  },
] as const;

for (const session of SESSIONS)
  for (const mode of MODES)
    for (const item of CASES)
      test(`a paragraph inserted ${item.name} goes beside it (${session}, ${mode})`, async () => {
        const input = await source(item.markdown);
        const reviewer = await FolioDocxReviewer.fromBuffer(input, {
          session,
        });
        expect(texts(reviewer)).toEqual([...item.original]);
        const { skipped } = reviewer.applyOperations(
          [
            {
              id: "heading",
              type: item.type,
              blockId: blockId(reviewer, item.anchor),
              text: "Heading",
            },
          ],
          { undefinedReferences: "refuse", mode },
        );
        expect(skipped).toEqual([]);
        expect(texts(reviewer)).toEqual([...item.inserted]);
        const reopened = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        expect(texts(reopened)).toEqual([...item.inserted]);
        if (mode === "direct") return;
        const rejected = await FolioDocxReviewer.fromBuffer(
          await reviewer.toBuffer(),
          { session },
        );
        rejected.rejectAll();
        expect(texts(rejected)).toEqual([...item.original]);
        reopened.acceptAll();
        expect(texts(reopened)).toEqual([...item.inserted]);
        expect(reopened.snapshot().blocks.length).toBe(item.inserted.length);
      });

for (const session of SESSIONS)
  test(`a commented paragraph inserted before a table that starts the body carries its comment (${session})`, async () => {
    const reviewer = await FolioDocxReviewer.fromBuffer(
      await source("| H |\n|---|\n| A |\n\nAfter."),
      { session },
    );
    const { skipped } = reviewer.applyOperations(
      [
        {
          id: "heading",
          type: "insertBeforeBlock",
          blockId: blockId(reviewer, "H"),
          text: "Heading",
          comment: { text: "Why a heading?" },
        },
      ],
      { undefinedReferences: "refuse", mode: "direct" },
    );
    expect(skipped).toEqual([]);
    expect(texts(reviewer)).toEqual(["Heading", "H", "A", "After."]);
    const reopened = await FolioDocxReviewer.fromBuffer(
      await reviewer.toBuffer(),
      { session },
    );
    const comments = reopened.getComments();
    expect(comments.map(({ text }) => text)).toEqual(["Why a heading?"]);
  });
