/**
 * Tracked edits against a reference suite of review shapes: what a reviewer's
 * Enter and Backspace leave in a package with tracking on, recorded from a
 * reference implementation. Each case builds the package's paragraphs (an
 * intro, the edited paragraphs, a tail), makes the edit with operations, and
 * compares the paragraphs with the reference: ids, properties, marks,
 * property changes and text.
 *
 * - Enter inside a paragraph moves the text before it to a new paragraph,
 *   whose mark is inserted; the paragraph keeps its id with the rest.
 * - Enter at the end adds a paragraph after it: the paragraph's mark is
 *   inserted, and the new paragraph takes the pending property change, or
 *   records one from the paragraph's properties when its own differ.
 * - Backspace at the start of a paragraph deletes the previous one's mark,
 *   and the paragraph takes the previous one's paragraph properties as a
 *   property change, unless the previous one holds no content; its mark run
 *   properties stay its own.
 */

import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  ParagraphFormatting,
  ParagraphPropertyChange,
} from "../../model/document";
import { applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES, type RevisionStamp } from "../types";

const DATE = "2026-01-01T00:00:00Z";
const REVIEWER: RevisionStamp = { id: 0, author: "Reviewer", date: DATE };

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const paragraph = (
  paraId: string,
  content: ParagraphContent[],
  formatting?: ParagraphFormatting,
  fields: Partial<Paragraph> = {},
): Paragraph =>
  formatting === undefined
    ? { ...fields, type: "paragraph", paraId, content }
    : { ...fields, type: "paragraph", paraId, formatting, content };

const INTRO = paragraph("1A000001", [text("Intro paragraph.")]);
const TAIL = paragraph("5E000005", [text("Tail paragraph.")]);

const documentOf = (...paragraphs: Paragraph[]): Document => ({
  package: { document: { content: [INTRO, ...paragraphs, TAIL] } },
});

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const edited = (document: Document, ops: readonly DocumentOp[]): BlockContent[] => {
  const applied = applyDocumentOps(document, ops);
  if (applied.isErr()) throw applied.error;
  // Every edit is undone exactly.
  const undone = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return applied.value.document.package.document.content;
};

const change = (id: number, previous?: ParagraphFormatting): ParagraphPropertyChange =>
  previous === undefined
    ? { type: "paragraphPropertyChange", info: { id, author: "Reviewer", date: DATE } }
    : {
        type: "paragraphPropertyChange",
        info: { id, author: "Reviewer", date: DATE },
        previousFormatting: previous,
      };

const mark = (kind: "ins" | "del", id: number) => ({
  kind,
  info: { id, author: "Reviewer", date: DATE },
});

/** The properties the cases vary: one at a time, as the reference suite does. */
const PROPERTIES: readonly [string, ParagraphFormatting, ParagraphFormatting][] = [
  ["alignment", { alignment: "center" }, { alignment: "right" }],
  ["indentation", { indentLeft: 720 }, { indentLeft: 1440 }],
  ["spacing", { spaceBefore: 240, spaceAfter: 120 }, { spaceAfter: 360 }],
  ["style", { styleId: "Heading1" }, { styleId: "BlockQuote" }],
];

describe("Enter inside a paragraph", () => {
  for (const [name, own] of PROPERTIES) {
    test(`moves the text before it to a new paragraph (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own);
      const content = edited(documentOf(alpha), [
        {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: at("2A000002", 6),
          newBlockId: "5791F953",
          revision: REVIEWER,
        },
      ]);
      expect(content).toEqual([
        INTRO,
        paragraph("5791F953", [text("Alpha ")], own, { pPrMark: mark("ins", 0) }),
        paragraph("2A000002", [text("text.")], own),
        TAIL,
      ]);
    });
  }
});

describe("Enter at the end of a paragraph", () => {
  const typeAfter = (newBlockId: string, fields?: Omit<Paragraph, "type" | "paraId" | "content">) =>
    [
      {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: at("2A000002", 11),
        newBlockId,
        ...(fields === undefined ? {} : { newParagraph: fields }),
        revision: REVIEWER,
        newIds: { revision: [1] },
      },
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at(newBlockId, 0),
        text: "New text.",
        runProps: "inherit",
        revision: { ...REVIEWER, id: 2 },
      },
    ] satisfies DocumentOp[];

  test("adds a paragraph after it with the same properties", () => {
    const alpha = paragraph("2A000002", [text("Alpha text.")], { alignment: "center" });
    expect(edited(documentOf(alpha), typeAfter("06279B03"))).toEqual([
      INTRO,
      { ...alpha, pPrMark: mark("ins", 0) },
      paragraph(
        "06279B03",
        [
          {
            type: "insertion",
            info: { id: 2, author: "Reviewer", date: DATE },
            content: [text("New text.")],
          },
        ],
        { alignment: "center" },
      ),
      TAIL,
    ]);
  });

  test("records the paragraph's properties on a new paragraph of another style", () => {
    const alpha = paragraph("2A000002", [text("Alpha text.")], { styleId: "Heading1" });
    expect(edited(documentOf(alpha), typeAfter("57B0C177", {}))).toEqual([
      INTRO,
      { ...alpha, pPrMark: mark("ins", 0) },
      paragraph(
        "57B0C177",
        [
          {
            type: "insertion",
            info: { id: 2, author: "Reviewer", date: DATE },
            content: [text("New text.")],
          },
        ],
        undefined,
        { propertyChanges: [change(1, { styleId: "Heading1" })] },
      ),
      TAIL,
    ]);
  });

  test("hands the paragraph's pending property change to the new paragraph", () => {
    const alpha = paragraph(
      "2A000002",
      [text("Alpha text.")],
      { alignment: "center" },
      {
        propertyChanges: [change(501, { alignment: "right" })],
      },
    );
    expect(edited(documentOf(alpha), typeAfter("4E9F542F"))).toEqual([
      INTRO,
      paragraph(
        "2A000002",
        [text("Alpha text.")],
        { alignment: "center" },
        {
          pPrMark: mark("ins", 0),
        },
      ),
      paragraph(
        "4E9F542F",
        [
          {
            type: "insertion",
            info: { id: 2, author: "Reviewer", date: DATE },
            content: [text("New text.")],
          },
        ],
        { alignment: "center" },
        { propertyChanges: [change(501, { alignment: "right" })] },
      ),
      TAIL,
    ]);
  });
});

describe("Backspace at the start of a paragraph", () => {
  const backspace: DocumentOp = {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "2A000002",
    nextBlockId: "3B000003",
    revision: REVIEWER,
    newIds: { revision: [1] },
  };

  for (const [name, own, next] of PROPERTIES) {
    test(`gives the paragraph the previous one's properties as a change (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own);
      const bravo = paragraph("3B000003", [text("Bravo text.")], next);
      expect(edited(documentOf(alpha, bravo), [backspace])).toEqual([
        INTRO,
        { ...alpha, pPrMark: mark("del", 0) },
        paragraph("3B000003", [text("Bravo text.")], own, { propertyChanges: [change(1, next)] }),
        TAIL,
      ]);
    });
  }

  test("leaves the paragraph's mark run properties its own", () => {
    const alpha = paragraph("2A000002", [text("Alpha text.")], {
      runProperties: { bold: true, color: { rgb: "C00000" } },
    });
    const bravo = paragraph("3B000003", [text("Bravo text.")], {
      runProperties: { italic: true, fontSize: 28 },
    });
    expect(edited(documentOf(alpha, bravo), [backspace])).toEqual([
      INTRO,
      { ...alpha, pPrMark: mark("del", 0) },
      bravo,
      TAIL,
    ]);
  });

  for (const [name, own, next] of PROPERTIES) {
    test(`deletes the mark of an empty paragraph and leaves the next one as it is (${name})`, () => {
      const alpha = paragraph("2A000002", [], own);
      const bravo = paragraph("3B000003", [text("Bravo text.")], next);
      expect(edited(documentOf(alpha, bravo), [backspace])).toEqual([
        INTRO,
        { ...alpha, pPrMark: mark("del", 0) },
        bravo,
        TAIL,
      ]);
    });
  }
});
