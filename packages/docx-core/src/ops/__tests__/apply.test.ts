import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  EMPTY_PROPERTY_SETS,
  INHERIT_RUN_PROPS,
  OP_STORIES,
} from "../types";

const RSID = [{ name: "rsidR", value: "00A1B2C3" }];

const run = (text: string, bold?: boolean): Run =>
  bold === undefined
    ? { type: "run", content: [{ type: "text", text }], preservedAttributes: RSID }
    : {
        type: "run",
        formatting: { bold },
        content: [{ type: "text", text }],
        preservedAttributes: RSID,
      };

const documentOf = (...content: Paragraph[]): Document => ({
  package: { document: { content } },
});

const paragraphs = (document: Document): Paragraph[] =>
  document.package.document.content.flatMap((block) => (block.type === "paragraph" ? [block] : []));

const at = (blockId: string, offset: number, zeroWidthBefore?: number) =>
  zeroWidthBefore === undefined
    ? { story: OP_STORIES.MAIN, blockId, offset }
    : { story: OP_STORIES.MAIN, blockId, offset, zeroWidthBefore };

const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

const refusalOf = (document: Document, op: DocumentOp): string | undefined => {
  const result = applyDocumentOp(document, op);
  return result.isErr() ? result.error.reason : undefined;
};

const undone = (document: Document, inverse: readonly DocumentOp[]): Document => {
  const result = applyDocumentOps(document, inverse);
  if (result.isErr()) {
    throw result.error;
  }
  return result.value.document;
};

describe("insertText", () => {
  test("inherited text joins the run before it and is undone by a deletion", () => {
    const first = run("ab", true);
    const second = run("cd");
    const document = documentOf({
      type: "paragraph",
      paraId: "00000001",
      content: [first, second],
    });
    const { document: next, inverse } = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 2),
      text: "X",
      runProps: INHERIT_RUN_PROPS,
    });
    const [paragraph] = paragraphs(next);
    expect(paragraph?.content).toEqual([
      { ...first, content: [{ type: "text", text: "abX" }] },
      second,
    ]);
    expect(paragraph?.content[1]).toBe(second);
    expect(inverse.map(({ type }) => type)).toEqual([DOCUMENT_OP_TYPES.DELETE_RANGE]);
    expect(undone(next, inverse)).toEqual(document);
  });

  test("text with other properties splits its host, and the undo merges it back", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("abcd")] });
    const { document: next, inverse } = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000001", 2),
      text: "X",
      runProps: { italic: true },
    });
    expect(paragraphs(next)[0]?.content).toEqual([
      run("ab"),
      { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "X" }] },
      run("cd"),
    ]);
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: at("00000001", 2, 0),
        to: at("00000001", 3, 0),
        expected: {
          content: [
            { type: "run", formatting: { italic: true }, content: [{ type: "text", text: "X" }] },
          ],
          openStart: 0,
          openEnd: 0,
        },
      },
      { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at: at("00000001", 2, 0), depth: 2 },
    ]);
    expect(undone(next, inverse)).toEqual(document);
  });
});

describe("deleteRange", () => {
  const document = documentOf({
    type: "paragraph",
    paraId: "00000001",
    content: [
      run("a"),
      { type: "bookmarkStart", id: 1, name: "_Ref1" },
      run("bc"),
      { type: "commentRangeStart", id: 4 },
      run("d"),
      { type: "bookmarkEnd", id: 1 },
      run("e"),
    ],
  });

  test("removes what lies strictly inside and keeps the markers at its ends", () => {
    const { document: next } = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 4),
    });
    expect(paragraphs(next)[0]?.content).toEqual([
      run("a"),
      { type: "bookmarkStart", id: 1, name: "_Ref1" },
      { type: "bookmarkEnd", id: 1 },
      run("e"),
    ]);
  });

  test("is undone by inserting the removed slice, markers and all", () => {
    const { document: next, inverse } = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 4),
    });
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at: at("00000001", 1, 1),
        slice: {
          content: [run("bc"), { type: "commentRangeStart", id: 4 }, run("d")],
          openStart: 0,
          openEnd: 0,
        },
      },
    ]);
    expect(undone(next, inverse)).toEqual(document);
  });

  test("a deletion through a run keeps its ends as one run and puts the middle back", () => {
    const original = documentOf({ type: "paragraph", paraId: "00000001", content: [run("abcd")] });
    const { document: next, inverse } = applied(original, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 3),
    });
    expect(paragraphs(next)[0]?.content).toEqual([run("ad")]);
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
        at: at("00000001", 1, 0),
        slice: { content: [run("bc")], openStart: 2, openEnd: 2 },
      },
    ]);
    expect(undone(next, inverse)).toEqual(original);
  });
});

describe("setRunProps", () => {
  test("is undone by restoring the prior values over the range and merging the cut run", () => {
    const original = documentOf({
      type: "paragraph",
      paraId: "00000001",
      content: [{ ...run("abcd"), formatting: {} }],
    });
    const { document: next, inverse } = applied(original, {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: at("00000001", 1),
      to: at("00000001", 3),
      patch: { bold: true },
    });
    expect(paragraphs(next)[0]?.content).toEqual([
      { ...run("a"), formatting: {} },
      { ...run("bc"), formatting: { bold: true } },
      { ...run("d"), formatting: {} },
    ]);
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
        from: at("00000001", 1, 0),
        to: at("00000001", 3, 0),
        patch: { bold: null },
        whenEmpty: EMPTY_PROPERTY_SETS.KEEP,
      },
      { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at: at("00000001", 1, 0), depth: 2 },
      { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at: at("00000001", 3, 0), depth: 2 },
    ]);
    expect(undone(next, inverse)).toStrictEqual(original);
  });
});

describe("splitBlock and joinBlocks", () => {
  const sectionProperties = { pageWidth: 11906 };
  const pPrMark = { kind: "ins" as const, info: { id: 5, author: "A" } };
  const original: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    textId: "77777777",
    formatting: { alignment: "center" },
    preservedAttributes: RSID,
    sectionProperties,
    pPrMark,
    content: [run("abcd")],
  };

  test("the second half takes the paragraph mark, and a join undoes the split", () => {
    const { document: next, inverse } = applied(documentOf(original), {
      type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
      at: at("00000001", 1),
      newBlockId: "0000ABCD",
    });
    expect(paragraphs(next)).toEqual([
      {
        type: "paragraph",
        paraId: "00000001",
        textId: "77777777",
        formatting: { alignment: "center" },
        preservedAttributes: RSID,
        content: [run("a")],
      },
      {
        type: "paragraph",
        paraId: "0000ABCD",
        formatting: { alignment: "center" },
        sectionProperties,
        pPrMark,
        content: [run("bcd")],
      },
    ]);
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "0000ABCD",
        depth: 2,
      },
    ]);
    expect(undone(next, inverse)).toEqual(documentOf(original));
  });

  test("a join is undone by a split that gives the second paragraph its own fields back", () => {
    const leading: Paragraph = {
      type: "paragraph",
      paraId: "00000001",
      pPrMark,
      content: [run("ab")],
    };
    const trailing: Paragraph = {
      type: "paragraph",
      paraId: "00000002",
      textId: "12345678",
      formatting: { alignment: "end" },
      preservedAttributes: RSID,
      content: [run("cd")],
    };
    const document = documentOf(leading, trailing);
    const { document: next, inverse } = applied(document, {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
      depth: 2,
    });
    expect(paragraphs(next)).toEqual([
      { type: "paragraph", paraId: "00000001", content: [run("abcd")] },
    ]);
    expect(inverse).toEqual([
      {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: at("00000001", 2, 0),
        newBlockId: "00000002",
        newParagraph: {
          textId: "12345678",
          formatting: { alignment: "end" },
          preservedAttributes: RSID,
        },
        firstMark: pPrMark,
      },
    ]);
    expect(undone(next, inverse)).toEqual(document);
  });

  test("a join does not remove a section break", () => {
    const document = documentOf(original, {
      type: "paragraph",
      paraId: "00000002",
      content: [run("e")],
    });
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
        story: OP_STORIES.MAIN,
        blockId: "00000001",
        nextBlockId: "00000002",
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY);
  });
});

describe("refusals", () => {
  const document = documentOf(
    { type: "paragraph", paraId: "0000ABCD", content: [run("a😀b")] },
    { type: "paragraph", paraId: "00000002", content: [run("c")] },
  );

  test.each([
    [
      "a new id that differs from a used one only in case",
      {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: at("0000ABCD", 1),
        newBlockId: "0000abcd",
      },
      DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION,
    ],
    [
      "the reserved zero id",
      { type: DOCUMENT_OP_TYPES.SPLIT_BLOCK, at: at("0000ABCD", 1), newBlockId: "00000000" },
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID,
    ],
    [
      "a position between the halves of a surrogate pair",
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at("0000ABCD", 2),
        text: "x",
        runProps: INHERIT_RUN_PROPS,
      },
      DOCUMENT_OP_REFUSAL_REASONS.SPLITS_SURROGATE_PAIR,
    ],
    [
      "a range across paragraphs",
      {
        type: DOCUMENT_OP_TYPES.DELETE_RANGE,
        from: at("0000ABCD", 0),
        to: at("00000002", 1),
      },
      DOCUMENT_OP_REFUSAL_REASONS.CROSS_BLOCK_RANGE,
    ],
    [
      "a tab in inserted text",
      {
        type: DOCUMENT_OP_TYPES.INSERT_TEXT,
        at: at("0000ABCD", 0),
        text: "\t",
        runProps: INHERIT_RUN_PROPS,
      },
      DOCUMENT_OP_REFUSAL_REASONS.INVALID_TEXT,
    ],
    [
      "a join of records that differ",
      { type: DOCUMENT_OP_TYPES.JOIN_INLINE, at: at("0000ABCD", 4), depth: 1 },
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    ],
  ] satisfies [string, DocumentOp, string][])("refuses %s", (_name, op, reason) => {
    expect(refusalOf(document, op)).toBe(reason);
  });

  test.each([
    ["past the paraId range", "FFFFFFFF"],
    ["not hex", "hello123"],
    ["too short", "12"],
  ])("refuses a new paragraph id %s", (_name, newBlockId) => {
    expect(
      refusalOf(document, {
        type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
        at: at("0000ABCD", 1),
        newBlockId,
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.INVALID_BLOCK_ID);
  });

  test("a replacement's ids are compared as hex, whatever their case", () => {
    const second = paragraphs(document)[1];
    const replace = (paraId: string): DocumentOp => ({
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: OP_STORIES.MAIN,
      expected: second === undefined ? [] : [second],
      blocks: [{ type: "paragraph", paraId, content: [] }],
    });
    expect(refusalOf(document, replace("0000abcd"))).toBe(DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION);
    expect(refusalOf(document, replace("0000abce"))).toBeUndefined();
  });

  test("a replacement does not move a section break", () => {
    const sectioned = documentOf(
      {
        type: "paragraph",
        paraId: "00000001",
        sectionProperties: { pageWidth: 11906 },
        content: [],
      },
      { type: "paragraph", paraId: "00000002", content: [run("b")] },
    );
    expect(
      refusalOf(sectioned, {
        type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
        story: OP_STORIES.MAIN,
        expected: paragraphs(sectioned),
        blocks: [{ type: "paragraph", paraId: "00000003", content: [run("b")] }],
      }),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.SECTION_BOUNDARY);
  });

  test("a new paragraph id is checked against every story in the package", () => {
    const withStories: Document = {
      package: {
        document: {
          content: [{ type: "paragraph", paraId: "00000001", content: [run("ab")] }],
          comments: [
            {
              id: 1,
              author: "A",
              content: [{ type: "paragraph", paraId: "0C0C0C0C", content: [] }],
            },
          ],
        },
        headers: new Map([
          [
            "rId7",
            {
              type: "header",
              hdrFtrType: "default",
              content: [{ type: "paragraph", paraId: "0E0E0E0E", content: [] }],
            },
          ],
        ]),
        footnotes: [
          {
            type: "footnote",
            id: 2,
            content: [{ type: "paragraph", paraId: "0F0F0F0F", content: [] }],
          },
        ],
      },
    };
    for (const id of ["0C0C0C0C", "0E0E0E0E", "0F0F0F0F"]) {
      expect(
        refusalOf(withStories, {
          type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
          at: at("00000001", 1),
          newBlockId: id,
        }),
      ).toBe(DOCUMENT_OP_REFUSAL_REASONS.ID_COLLISION);
    }
  });

  test("an undo authored against another state is refused as stale", () => {
    const edit = applied(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000002", 1),
      text: "d",
      runProps: INHERIT_RUN_PROPS,
    });
    const later = applied(edit.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000002", 0),
      text: "e",
      runProps: INHERIT_RUN_PROPS,
    });
    const undo = applyDocumentOps(later.document, edit.inverse);
    expect(undo.isErr() ? undo.error.reason : undefined).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });
});
