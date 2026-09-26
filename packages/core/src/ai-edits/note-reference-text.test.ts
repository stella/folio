/**
 * A footnote or endnote reference is structure, not text.
 *
 * The editor holds a reference as a text node whose characters are the note's
 * package id. Read as text, a paragraph with footnotes 10 and 30 said
 * `Term A10 and Term B30.` to `getContent()`, the snapshot and the agent
 * tools, while Markdown said `Term A[^1] and Term B[^2].`; a search for `10`
 * found the reference, and replacing it deleted the reference in direct mode
 * (orphaning its note) while the same operation was silently undone in
 * tracked mode.
 *
 * Every reader now shows a reference as the marker Markdown writes — numbered
 * in reading order, `[^e1]` for an endnote — and a text operation that would
 * cut into, rewrite or drop a marker is refused as `protectedReference` in
 * every mode, before anything changes.
 */

import { describe, expect, test } from "bun:test";

import { fromMarkdown } from "../markdown";
import { createDocx, docxToMarkdown, ensureParaIds } from "../server";
import type { Endnote, Footnote, Paragraph } from "../types/document";
import { FolioDocxReviewer } from "./headless";
import { createFolioAITextRangeHandle } from "./snapshot";
import type { FolioAIEditApplyMode, FolioAIEditOperation } from "./types";

const noteBody = (text: string): Paragraph => ({
  type: "paragraph",
  formatting: {},
  content: [{ type: "run", formatting: {}, content: [{ type: "text", text }] }],
});

/**
 * One paragraph with two footnote references and an endnote reference, their
 * package ids deliberately out of reading order and far from 1, 2, 3.
 */
const buildDocx = async (): Promise<ArrayBuffer> => {
  const model = fromMarkdown("Placeholder.\n\nSecond paragraph.");
  const first = model.package.document.content[0];
  if (first?.type !== "paragraph") {
    throw new Error("fixture paragraph is missing");
  }
  first.content = [
    {
      type: "run",
      formatting: {},
      content: [
        { type: "text", text: "Term A" },
        { type: "footnoteRef", id: 30 },
        { type: "text", text: " and Term B" },
        { type: "footnoteRef", id: 10 },
        { type: "text", text: " see" },
        { type: "endnoteRef", id: 7 },
        { type: "text", text: "." },
      ],
    },
  ];
  model.package.footnotes = [30, 10].map(
    (id, index): Footnote => ({
      type: "footnote",
      id,
      noteType: "normal",
      content: [noteBody(`Footnote ${index + 1}.`)],
    }),
  );
  model.package.endnotes = [
    { type: "endnote", id: 7, noteType: "normal", content: [noteBody("Endnote one.")] } as Endnote,
  ];
  const bytes = (await ensureParaIds(new Uint8Array(await createDocx(model)))).docx;
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
};

const TEXT = "Term A[^1] and Term B[^2] see[^e1].";

const markdownOf = async (docx: ArrayBuffer): Promise<string> => {
  const result = await docxToMarkdown(docx, {
    annotations: "strip",
    trackedChanges: "clean",
    comments: "strip",
    footnotes: "keep",
  });
  return typeof result === "string" ? result : result.markdown;
};

const MARKDOWN = `${TEXT}\n\nSecond paragraph.\n\n[^1]: Footnote 1.\n[^2]: Footnote 2.\n[^e1]: Endnote one.`;

const MODES: readonly FolioAIEditApplyMode[] = ["direct", "tracked-changes", "suggested"];

const open = async () => {
  const reviewer = await FolioDocxReviewer.fromBuffer(await buildDocx(), { author: "Reviewer" });
  const block = reviewer.getContent()[0];
  if (!block) {
    throw new Error("fixture block is missing");
  }
  return { reviewer, block };
};

/** Save, reopen with every revision accepted, and read it as Markdown. */
const savedMarkdown = async (reviewer: FolioDocxReviewer): Promise<string> => {
  const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
  reopened.acceptAll();
  return markdownOf(await reopened.toBuffer());
};

/**
 * What the edit leaves: the block's text in the session and, where the mode
 * writes its edits into the package (a suggestion stays with the host until
 * accepted), the saved package accepted and read as Markdown.
 */
const expectEdited = async (
  reviewer: FolioDocxReviewer,
  mode: FolioAIEditApplyMode,
  text: string,
): Promise<void> => {
  expect(reviewer.getContent()[0]?.text).toBe(text);
  if (mode !== "suggested") {
    expect(await savedMarkdown(reviewer)).toBe(MARKDOWN.replace(TEXT, text));
  }
};

describe("note references read as their markers", () => {
  test("every reader shows the marker Markdown writes, never the package id", async () => {
    const { reviewer, block } = await open();

    expect(block.text).toBe(TEXT);
    expect(reviewer.snapshot().blocks[0]?.text).toBe(TEXT);
    expect(reviewer.getContentAsText()).toContain(TEXT);
    expect(await markdownOf(await buildDocx())).toBe(MARKDOWN);
    expect(block.structuralBoundaries).toEqual([
      { type: "noteReference", noteType: "footnote", offset: 6, length: 4 },
      { type: "noteReference", noteType: "footnote", offset: 21, length: 4 },
      { type: "noteReference", noteType: "endnote", offset: 29, length: 5 },
    ]);
    // The preview runs spell the same text.
    expect(block.previewRuns?.map((run) => run.text).join("")).toBe(TEXT);
  });
});

describe("a text operation cannot reach a note reference", () => {
  const refused = (id: string) => [{ id, reason: "protectedReference" }];

  test.each(MODES)("%s: the package id is not text to find", async (mode) => {
    const { reviewer, block } = await open();
    const result = reviewer.applyOperations(
      [{ id: "id", type: "replaceInBlock", blockId: block.id, find: "30", replace: "X" }],
      { mode },
    );
    expect(result.skipped).toEqual([{ id: "id", reason: "missingFind" }]);
  });

  test.each(MODES)("%s: rewriting, dropping or cutting into a marker is refused", async (mode) => {
    const { reviewer, block } = await open();
    const range = createFolioAITextRangeHandle({
      blockId: block.id,
      text: block.text,
      startOffset: 7,
      endOffset: 14,
    });
    const second = reviewer.getContent()[1];
    if (range === null || second === undefined) {
      throw new Error("fixture is invalid");
    }
    const operations: FolioAIEditOperation[] = [
      { id: "marker", type: "replaceInBlock", blockId: block.id, find: "[^1]", replace: "X" },
      { id: "cut", type: "replaceInBlock", blockId: block.id, find: "^2] see", replace: " see" },
      { id: "drop", type: "replaceInBlock", blockId: block.id, find: "A[^1] and", replace: "and" },
      {
        id: "reorder",
        type: "replaceInBlock",
        blockId: block.id,
        find: "[^1] and Term B[^2]",
        replace: "[^2] and Term B[^1]",
      },
      { id: "range", type: "replaceRange", range, replace: "and" },
      {
        id: "forged",
        type: "replaceInBlock",
        blockId: second.id,
        find: "Second",
        replace: "Second[^3]",
      },
      { id: "block", type: "replaceBlock", blockId: block.id, text: "Term A and Term B." },
      {
        id: "emphasis",
        type: "replaceBlock",
        blockId: block.id,
        text: "**Term** A[^1] and Term B[^2] see[^e1].",
      },
      {
        id: "unformatted",
        type: "replaceBlock",
        blockId: block.id,
        text: "Term C[^1] and Term B[^2] see[^e1].",
        preserveFormatting: false,
      },
    ];
    for (const operation of operations) {
      const result = reviewer.applyOperations([operation], { mode });
      expect({ id: operation.id, skipped: result.skipped }).toEqual({
        id: operation.id,
        skipped: refused(operation.id),
      });
    }
    expect(reviewer.getContent()[0]?.text).toBe(TEXT);
    expect(await savedMarkdown(reviewer)).toBe(MARKDOWN);
  });
});

describe("prose beside a note reference stays editable", () => {
  test.each(MODES)("%s: the edit applies and every reference survives a save", async (mode) => {
    const { reviewer, block } = await open();
    const result = reviewer.applyOperations(
      [
        {
          id: "prose",
          type: "replaceInBlock",
          blockId: block.id,
          find: "Term A[^1] and",
          replace: "Term C[^1], and",
        },
      ],
      { mode },
    );
    expect(result.skipped).toEqual([]);
    await expectEdited(reviewer, mode, "Term C[^1], and Term B[^2] see[^e1].");
  });

  test.each(MODES)("%s: a whole-block rewrite keeps the markers it repeats", async (mode) => {
    const { reviewer, block } = await open();
    const rewritten = "Terms A[^1] and B[^2] (see[^e1]).";
    const result = reviewer.applyOperations(
      [{ id: "block", type: "replaceBlock", blockId: block.id, text: rewritten }],
      { mode },
    );
    expect(result.skipped).toEqual([]);
    await expectEdited(reviewer, mode, rewritten);
  });

  test.each(["direct", "tracked-changes"] as const)(
    "%s: a comment may cover a reference",
    async (mode) => {
      const { reviewer, block } = await open();
      const result = reviewer.applyOperations(
        [
          {
            id: "comment",
            type: "commentOnBlock",
            blockId: block.id,
            quote: "Term B[^2]",
            comment: { text: "Check the note." },
          },
        ],
        { mode },
      );
      expect(result.skipped).toEqual([]);
      expect(reviewer.getComments().map((comment) => comment.anchoredText)).toEqual(["Term B[^2]"]);
      expect(await savedMarkdown(reviewer)).toBe(MARKDOWN);
    },
  );
});
