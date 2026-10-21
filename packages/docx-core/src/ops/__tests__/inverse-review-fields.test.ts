import { describe, expect, test } from "bun:test";
import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  EMPTY_PROPERTY_SETS,
  OP_STORIES,
  SPLIT_HALVES,
  type DocumentOp,
} from "../types";

const originalDocument = (): Document =>
  normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
          },
        ],
        sections: [
          {
            properties: { pageWidth: 12_240, pageHeight: 15_840 },
            content: [
              {
                type: "paragraph",
                paraId: "00000001",
                content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
              },
              {
                type: "paragraph",
                paraId: "00000002",
                content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
              },
            ],
          },
        ],
        finalSectionProperties: { pageWidth: 12_240, pageHeight: 15_840 },
        comments: [
          { id: 7, author: "A", content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }] },
        ],
      },
      footnotes: [
        {
          type: "footnote",
          id: 2,
          content: [{ type: "paragraph", paraId: "7FFFFFF1", content: [] }],
        },
      ],
      settings: { defaultTabStop: 720 },
      properties: { title: "synthetic" },
    },
    warnings: ["kept"],
  } satisfies Document);

const operations = [
  {
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
    newBlockId: "10000000",
    newParagraph: { formatting: { runProperties: {} } },
  },
  {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "10000000",
    depth: 0,
    survivor: SPLIT_HALVES.FIRST,
  },
  {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "00000002",
    depth: 1,
    survivor: SPLIT_HALVES.FIRST,
  },
  {
    type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    patch: { styleId: "Heading2" },
  },
  {
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    text: "x",
    runProps: {},
  },
] as const satisfies readonly DocumentOp[];

const appliedSequence = () => {
  let document = originalDocument();
  const inverses: (readonly DocumentOp[])[] = [];
  for (const op of operations) {
    const applied = applyDocumentOp(document, op).unwrap();
    document = applied.document;
    inverses.push(applied.inverse);
  }
  return { document, inverse: inverses.toReversed().flat() };
};

describe("inverse paragraph review fields", () => {
  test("a paragraph property inverse preserves an explicit empty property set", () => {
    let document = originalDocument();
    for (const op of operations.slice(0, 3))
      document = applyDocumentOp(document, op).unwrap().document;
    const beforePatch = document;
    const patched = applyDocumentOp(document, operations[3]).unwrap();
    const inverse = patched.inverse.at(0);
    expect(inverse).toMatchObject({
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_PROPS,
      whenEmpty: EMPTY_PROPERTY_SETS.KEEP,
    });
    if (inverse === undefined) throw new Error("A paragraph property edit must have an inverse.");
    expect(applyDocumentOp(patched.document, inverse).unwrap().document).toStrictEqual(beforePatch);
  });

  test("split, join, paragraph patch, and content inverses restore review fields in reverse", () => {
    const original = originalDocument();
    const sequence = appliedSequence();
    expect(applyDocumentOps(sequence.document, sequence.inverse).unwrap().document).toStrictEqual(
      original,
    );
    const batch = applyDocumentOps(original, operations).unwrap();
    expect(batch.document).toStrictEqual(sequence.document);
    expect(applyDocumentOps(batch.document, batch.inverse).unwrap().document).toStrictEqual(
      original,
    );
  });
});
