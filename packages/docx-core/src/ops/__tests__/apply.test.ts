import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, Run } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, type DocumentOp, INHERIT_RUN_PROPS, OP_STORIES } from "../types";

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

const at = (blockId: string, offset: number) => ({ story: OP_STORIES.MAIN, blockId, offset });

const apply = (document: Document, op: DocumentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) {
    throw applied.error;
  }
  return applied.value.document;
};

const refusalOf = (document: Document, op: DocumentOp): string | undefined => {
  const applied = applyDocumentOp(document, op);
  return applied.isErr() ? applied.error.reason : undefined;
};

describe("insertText", () => {
  test("inherited text joins the run before it and leaves its neighbours as they were", () => {
    const first = run("ab", true);
    const second = run("cd");
    const document = documentOf({
      type: "paragraph",
      paraId: "00000001",
      content: [first, second],
    });
    const next = apply(document, {
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
  });

  test("text with other properties becomes its own run, splitting the host", () => {
    const document = documentOf({ type: "paragraph", paraId: "00000001", content: [run("abcd")] });
    const next = apply(document, {
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
  });
});

describe("deleteRange", () => {
  test("keeps the range markers and drops a run it empties", () => {
    const document = documentOf({
      type: "paragraph",
      paraId: "00000001",
      content: [
        run("a"),
        { type: "bookmarkStart", id: 1, name: "_Ref1" },
        run("bc"),
        { type: "bookmarkEnd", id: 1 },
        run("d"),
      ],
    });
    const next = apply(document, {
      type: DOCUMENT_OP_TYPES.DELETE_RANGE,
      from: at("00000001", 1),
      to: at("00000001", 3),
    });
    expect(paragraphs(next)[0]?.content).toEqual([
      run("a"),
      { type: "bookmarkStart", id: 1, name: "_Ref1" },
      { type: "bookmarkEnd", id: 1 },
      run("d"),
    ]);
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

  test("the second half takes the paragraph mark; the first keeps the identity", () => {
    const next = apply(documentOf(original), {
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
  ] satisfies [string, DocumentOp, string][])("refuses %s", (_name, op, reason) => {
    expect(refusalOf(document, op)).toBe(reason);
  });

  test("an inverse authored against another state is refused as stale", () => {
    const edit = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000002", 1),
      text: "d",
      runProps: INHERIT_RUN_PROPS,
    });
    if (edit.isErr()) throw edit.error;
    const later = apply(edit.value.document, {
      type: DOCUMENT_OP_TYPES.INSERT_TEXT,
      at: at("00000002", 0),
      text: "e",
      runProps: INHERIT_RUN_PROPS,
    });
    const undo = applyDocumentOps(later, edit.value.inverse);
    expect(undo.isErr() ? undo.error.reason : undefined).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
  });
});
