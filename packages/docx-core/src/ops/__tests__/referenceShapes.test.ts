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
 * - Accepting or rejecting every change: a paragraph mark that goes takes
 *   its paragraph's properties with it, and leaves the paragraph after it
 *   (its id, properties, mark run properties and pending property change),
 *   with the first's content before its own, in whatever order.
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
import { asParagraphContent, childNodes, rebuildNode, type InlineNode } from "../leaves";
import { isTrackedWrapper } from "../review";
import { applyDocumentOps } from "../apply";
import { IDENTITY_SPACES, identityKeysIn } from "../ids";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
  type RevisionStamp,
} from "../types";

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
  // Reference OOXML shapes omit operation-private cut provenance. The exact
  // model (including provenance) is still checked by the inverse above and
  // the resolution-provenance properties.
  const referenceInline = (node: InlineNode): InlineNode => {
    const children = childNodes(node);
    const mapped = children === undefined ? node : rebuildNode(node, children.map(referenceInline));
    if (!isTrackedWrapper(mapped)) return mapped;
    const { resolutionJoins: _resolutionJoins, ...reference } = mapped;
    return reference;
  };
  return applied.value.document.package.document.content.map((block) => {
    if (block.type !== "paragraph") return block;
    const content = asParagraphContent(block.content.map(referenceInline));
    if (block.pPrMark === undefined) return Object.assign({}, block, { content });
    const { resolutionJoin: _resolutionJoin, ...pPrMark } = block.pPrMark;
    return Object.assign({}, block, { content, pPrMark });
  });
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

// ---------------------------------------------------------------------------
// Resolving every change: the reference's accept-all and reject-all results
// ---------------------------------------------------------------------------

const deleted = (id: number, content: ParagraphContent[]): ParagraphContent => ({
  type: "deletion",
  info: { id, author: "Reviewer", date: DATE },
  content: content.flatMap((item) => (item.type === "run" ? [item] : [])),
});

const inserted = (id: number, content: ParagraphContent[]): ParagraphContent => ({
  type: "insertion",
  info: { id, author: "Reviewer", date: DATE },
  content: content.flatMap((item) => (item.type === "run" ? [item] : [])),
});

const storyRevisionIds = (document: Document): number[] => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return identityKeysIn(document.package.document.content).flatMap((key) =>
    key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : [],
  );
};

/** Every change in the story accepted or rejected, the way a review's accept all does. */
const resolvedAll = (
  document: Document,
  decision: RevisionDecision,
  ids = storyRevisionIds(document),
): BlockContent[] => {
  const applied = applyDocumentOps(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: ids,
      decision,
    },
  ]);
  if (applied.isErr()) throw applied.error;
  const undone = applyDocumentOps(applied.value.document, applied.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  // Reference OOXML shapes omit operation-private cut provenance. The exact
  // model (including provenance) is still checked by the inverse above and
  // the resolution-provenance properties.
  return applied.value.document.package.document.content.map((block) => {
    if (block.type !== "paragraph" || block.pPrMark === undefined) return block;
    const { resolutionJoin: _resolutionJoin, ...pPrMark } = block.pPrMark;
    return Object.assign({}, block, { pPrMark });
  });
};

const ACCEPT = REVISION_DECISIONS.ACCEPT;
const REJECT = REVISION_DECISIONS.REJECT;

describe("Removing a paragraph mark leaves the paragraph after it", () => {
  for (const [name, own, next] of PROPERTIES) {
    const bravo = paragraph("3B000003", [text("Bravo text.")], next);

    test(`accepting a deleted mark (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own, {
        pPrMark: mark("del", 101),
      });
      const document = documentOf(alpha, bravo);
      expect(resolvedAll(document, ACCEPT)).toEqual([
        INTRO,
        paragraph("3B000003", [text("Alpha text.Bravo text.")], next),
        TAIL,
      ]);
      expect(resolvedAll(document, REJECT)).toEqual([
        INTRO,
        paragraph("2A000002", [text("Alpha text.")], own),
        bravo,
        TAIL,
      ]);
    });

    test(`accepting a deleted mark of an empty paragraph (${name})`, () => {
      const alpha = paragraph("2A000002", [], own, { pPrMark: mark("del", 101) });
      expect(resolvedAll(documentOf(alpha, bravo), ACCEPT)).toEqual([INTRO, bravo, TAIL]);
    });

    test(`accepting a deleted mark whose paragraph's text is deleted too (${name})`, () => {
      const alpha = paragraph("2A000002", [deleted(102, [text("Alpha text.")])], own, {
        pPrMark: mark("del", 101),
      });
      const document = documentOf(alpha, bravo);
      expect(resolvedAll(document, ACCEPT)).toEqual([INTRO, bravo, TAIL]);
      expect(resolvedAll(document, REJECT)).toEqual([
        INTRO,
        paragraph("2A000002", [text("Alpha text.")], own),
        bravo,
        TAIL,
      ]);
    });

    test(`accepting a deleted mark before deleted text (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own, {
        pPrMark: mark("del", 101),
      });
      const deletedBravo = paragraph("3B000003", [deleted(102, [text("Bravo text.")])], next);
      expect(resolvedAll(documentOf(alpha, deletedBravo), ACCEPT)).toEqual([
        INTRO,
        paragraph("3B000003", [text("Alpha text.")], next),
        TAIL,
      ]);
    });

    test(`rejecting an inserted mark (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own, {
        pPrMark: mark("ins", 101),
      });
      const document = documentOf(alpha, bravo);
      expect(resolvedAll(document, REJECT)).toEqual([
        INTRO,
        paragraph("3B000003", [text("Alpha text.Bravo text.")], next),
        TAIL,
      ]);
      expect(resolvedAll(document, ACCEPT)).toEqual([
        INTRO,
        paragraph("2A000002", [text("Alpha text.")], own),
        bravo,
        TAIL,
      ]);
    });

    test(`rejecting an inserted mark before inserted text (${name})`, () => {
      const alpha = paragraph("2A000002", [text("Alpha text.")], own, {
        pPrMark: mark("ins", 101),
      });
      const insertedBravo = paragraph("3B000003", [inserted(102, [text("Bravo text.")])], next);
      expect(resolvedAll(documentOf(alpha, insertedBravo), REJECT)).toEqual([
        INTRO,
        paragraph("3B000003", [text("Alpha text.")], next),
        TAIL,
      ]);
    });
  }

  test("the paragraph whose mark goes drops its pending property change", () => {
    const alpha = paragraph(
      "2A000002",
      [text("Alpha text.")],
      { alignment: "center" },
      {
        pPrMark: mark("del", 101),
        propertyChanges: [change(102)],
      },
    );
    const bravo = paragraph("3B000003", [text("Bravo text.")], { alignment: "right" });
    const document = documentOf(alpha, bravo);
    expect(resolvedAll(document, ACCEPT)).toEqual([
      INTRO,
      paragraph("3B000003", [text("Alpha text.Bravo text.")], { alignment: "right" }),
      TAIL,
    ]);
    expect(resolvedAll(document, REJECT)).toEqual([
      INTRO,
      paragraph("2A000002", [text("Alpha text.")]),
      bravo,
      TAIL,
    ]);
  });

  test("the paragraph left keeps its own pending property change until it is resolved", () => {
    const alpha = paragraph(
      "2A000002",
      [text("Alpha text.")],
      { styleId: "Heading1" },
      {
        pPrMark: mark("del", 101),
      },
    );
    const bravo = paragraph(
      "3B000003",
      [text("Bravo text.")],
      { styleId: "BlockQuote" },
      {
        propertyChanges: [change(102)],
      },
    );
    const document = documentOf(alpha, bravo);
    expect(resolvedAll(document, ACCEPT, [101])).toEqual([
      INTRO,
      { ...bravo, content: [text("Alpha text.Bravo text.")] },
      TAIL,
    ]);
    expect(resolvedAll(document, REJECT)).toEqual([
      INTRO,
      paragraph("2A000002", [text("Alpha text.")], { styleId: "Heading1" }),
      paragraph("3B000003", [text("Bravo text.")]),
      TAIL,
    ]);
  });

  test("the paragraph left keeps its mark run properties", () => {
    const alpha = paragraph(
      "2A000002",
      [text("Alpha text.")],
      {
        runProperties: { bold: true, color: { rgb: "C00000" } },
      },
      { pPrMark: mark("del", 101) },
    );
    const bravo = paragraph("3B000003", [text("Bravo text.")], {
      runProperties: { italic: true, fontSize: 28 },
    });
    expect(resolvedAll(documentOf(alpha, bravo), ACCEPT)).toEqual([
      INTRO,
      { ...bravo, content: [text("Alpha text.Bravo text.")] },
      TAIL,
    ]);
  });

  describe("a chain of removed marks ends in the paragraph after the last", () => {
    const alpha = paragraph(
      "2A000002",
      [text("Alpha text.")],
      { alignment: "center" },
      {
        pPrMark: mark("del", 101),
      },
    );
    const bravo = paragraph(
      "3B000003",
      [text("Bravo text.")],
      { indentLeft: 720, alignment: "right" },
      {
        pPrMark: mark("del", 102),
      },
    );
    const charlie = paragraph("4C000004", [text("Charlie text.")], { styleId: "BlockQuote" });
    const document = documentOf(alpha, bravo, charlie);
    const joined = paragraph("4C000004", [text("Alpha text.Bravo text.Charlie text.")], {
      styleId: "BlockQuote",
    });

    test("at once", () => {
      expect(resolvedAll(document, ACCEPT)).toEqual([INTRO, joined, TAIL]);
    });

    test("in either order", () => {
      const firstThenSecond = applyDocumentOps(document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [101],
          decision: ACCEPT,
        },
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [102],
          decision: ACCEPT,
        },
      ]);
      const secondThenFirst = applyDocumentOps(document, [
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [102],
          decision: ACCEPT,
        },
        {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [101],
          decision: ACCEPT,
        },
      ]);
      if (firstThenSecond.isErr()) throw firstThenSecond.error;
      if (secondThenFirst.isErr()) throw secondThenFirst.error;
      expect(firstThenSecond.value.document.package.document.content).toEqual([
        INTRO,
        joined,
        TAIL,
      ]);
      expect(secondThenFirst.value.document.package.document.content).toEqual([
        INTRO,
        joined,
        TAIL,
      ]);
    });
  });
});

describe("Resolving inline changes", () => {
  test("a deletion nested in another author's insertion", () => {
    const keep = paragraph("2A000002", [
      text("Keep "),
      {
        type: "insertion",
        info: { id: 201, author: "Author One", date: DATE },
        content: [text("added ")].flatMap((item) => (item.type === "run" ? [item] : [])),
      },
      {
        type: "insertion",
        info: { id: 202, author: "Author One", date: DATE },
        content: [
          {
            type: "deletion",
            info: { id: 203, author: "Author Two", date: "2026-01-02T00:00:00Z" },
            content: [text("retracted ")].flatMap((item) => (item.type === "run" ? [item] : [])),
          },
        ],
      },
      text("end."),
    ]);
    const document = documentOf(keep);
    expect(resolvedAll(document, ACCEPT)).toEqual([
      INTRO,
      paragraph("2A000002", [text("Keep added end.")]),
      TAIL,
    ]);
    expect(resolvedAll(document, REJECT)).toEqual([
      INTRO,
      paragraph("2A000002", [text("Keep end.")]),
      TAIL,
    ]);
  });

  test("paragraph and run property changes", () => {
    const changed = paragraph(
      "2A000002",
      [
        {
          type: "run",
          formatting: { bold: true },
          propertyChanges: [
            {
              type: "runPropertyChange",
              info: { id: 302, author: "Reviewer", date: DATE },
              previousFormatting: { italic: true },
            },
          ],
          content: [{ type: "text", text: "Bold now " }],
        },
        text("plain."),
      ],
      { alignment: "center" },
      { propertyChanges: [change(301, { alignment: "right" })] },
    );
    const document = documentOf(changed);
    expect(resolvedAll(document, ACCEPT)).toEqual([
      INTRO,
      paragraph(
        "2A000002",
        [
          {
            type: "run",
            formatting: { bold: true },
            content: [{ type: "text", text: "Bold now " }],
          },
          text("plain."),
        ],
        { alignment: "center" },
      ),
      TAIL,
    ]);
    expect(resolvedAll(document, REJECT)).toEqual([
      INTRO,
      paragraph(
        "2A000002",
        [
          {
            type: "run",
            formatting: { italic: true },
            content: [{ type: "text", text: "Bold now " }],
          },
          text("plain."),
        ],
        { alignment: "right" },
      ),
      TAIL,
    ]);
  });
});

describe("Tracked Enter and Backspace, then every change resolved", () => {
  const alphaOf = (formatting: ParagraphFormatting, fields: Partial<Paragraph> = {}) =>
    paragraph("2A000002", [text("Alpha text.")], formatting, fields);
  const bravo = paragraph("3B000003", [text("Bravo text.")], { alignment: "right" });

  const tracked = (document: Document, ops: readonly DocumentOp[]): Document => {
    const applied = applyDocumentOps(document, ops);
    if (applied.isErr()) throw applied.error;
    return applied.value.document;
  };

  test("Enter inside a paragraph: rejecting gives the paragraph back under its own id", () => {
    const document = documentOf(alphaOf({ alignment: "center" }), bravo);
    const split = tracked(document, [
      {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: at("2A000002", 6),
        newBlockId: "5791F953",
        revision: REVIEWER,
      },
    ]);
    expect(resolvedAll(split, REJECT)).toEqual(document.package.document.content);
    expect(resolvedAll(split, ACCEPT)).toEqual([
      INTRO,
      paragraph("5791F953", [text("Alpha ")], { alignment: "center" }),
      paragraph("2A000002", [text("text.")], { alignment: "center" }),
      bravo,
      TAIL,
    ]);
  });

  for (const [name, formatting, next, fields] of [
    ["same properties", { alignment: "center" }, undefined, {}],
    ["another style", { styleId: "Heading1" }, {}, {}],
    [
      "a pending change",
      { alignment: "center" },
      undefined,
      { propertyChanges: [change(501, { alignment: "right" })] },
    ],
  ] as const) {
    test(`Enter at the end (${name}): rejecting leaves the new paragraph with the old properties`, () => {
      const alpha = alphaOf(formatting, fields);
      const split = tracked(documentOf(alpha, bravo), [
        {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: at("2A000002", 11),
          newBlockId: "06279B03",
          ...(next === undefined ? {} : { newParagraph: next }),
          revision: REVIEWER,
          newIds: { revision: [1] },
        },
        {
          type: DOCUMENT_OP_TYPES.INSERT_TEXT,
          at: at("06279B03", 0),
          text: "New text.",
          runProps: "inherit",
          revision: { ...REVIEWER, id: 2 },
        },
      ]);
      // The pending change is rejected with the rest, which gives the paragraph its old properties.
      const old =
        fields.propertyChanges === undefined ? formatting : { alignment: "right" as const };
      expect(resolvedAll(split, REJECT)).toEqual([
        INTRO,
        paragraph("06279B03", [text("Alpha text.")], old),
        bravo,
        TAIL,
      ]);
    });
  }

  test("Backspace: accepting gives the text the previous paragraph's look under the next one's id", () => {
    const document = documentOf(alphaOf({ alignment: "center" }), bravo);
    const joined = tracked(document, [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "2A000002",
        nextBlockId: "3B000003",
        revision: REVIEWER,
        newIds: { revision: [1] },
      },
    ]);
    expect(resolvedAll(joined, ACCEPT)).toEqual([
      INTRO,
      paragraph("3B000003", [text("Alpha text.Bravo text.")], { alignment: "center" }),
      TAIL,
    ]);
    expect(resolvedAll(joined, REJECT)).toEqual(document.package.document.content);
  });

  test("Backspace from an empty paragraph: accepting leaves the next paragraph as it was", () => {
    const document = documentOf(paragraph("2A000002", [], { alignment: "center" }), bravo);
    const joined = tracked(document, [
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "2A000002",
        nextBlockId: "3B000003",
        revision: REVIEWER,
      },
    ]);
    expect(resolvedAll(joined, ACCEPT)).toEqual([INTRO, bravo, TAIL]);
  });
});
