import { describe, expect, test } from "bun:test";

import type { Document } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps } from "../contract";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  SPLIT_HALVES,
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
                content: [
                  { type: "run", content: [{ type: "text", text: "a" }] },
                ],
              },
              {
                type: "paragraph",
                paraId: "00000002",
                content: [
                  { type: "run", content: [{ type: "text", text: "a" }] },
                ],
              },
            ],
          },
        ],
        finalSectionProperties: { pageWidth: 12_240, pageHeight: 15_840 },
        comments: [
          {
            id: 7,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
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
  test("an absent and an empty paragraph property set are the same review state", () => {
    const document = originalDocument();
    const result = applyDocumentOp(document, {
      type: DOCUMENT_OP_TYPES.SET_PARAGRAPH_REVIEW,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      expected: { formatting: {} },
      review: { formatting: {} },
    }).unwrap();

    expect(result.document).toBe(document);
    expect(result.inverse).toEqual([]);
    expect(result.document.package.document.content[0]).not.toHaveProperty(
      "formatting",
    );
  });

  test("split, join, paragraph patch, and content inverses restore review fields in reverse", () => {
    const original = originalDocument();
    const sequence = appliedSequence();
    expect(
      applyDocumentOps(sequence.document, sequence.inverse).unwrap().document,
    ).toStrictEqual(original);
    const batch = applyDocumentOps(original, operations).unwrap();
    expect(batch.document).toStrictEqual(sequence.document);
    expect(
      applyDocumentOps(batch.document, batch.inverse).unwrap().document,
    ).toStrictEqual(original);
  });
});
