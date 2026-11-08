/**
 * A point comment reference (`w:commentReference` in its own run) can stand
 * inside a tracked insertion. Splitting, inserting into or deleting from that
 * insertion keeps the reference inside a tracked insertion, and each edit's
 * inverse restores the paragraph.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, type DocumentOp, OP_STORIES } from "../types";

const run = (text: string): Run => ({
  type: "run",
  content: [{ type: "text", text }],
});
const info = { id: 8, author: "A", date: "2024-01-01T00:00:00Z" };
const reference = { type: "commentReference", id: 5 } as const;

/** "ab", the reference, then "cd": offsets 0..2, 2..3 and 3..5. */
const paragraph: Paragraph = {
  type: "paragraph",
  paraId: "00000001",
  content: [
    {
      type: "insertion",
      info,
      content: [run("ab"), reference, run("cd")],
    },
  ],
};
const documentOf = (): Document => ({
  package: { document: { content: [structuredClone(paragraph)] } },
});
const at = (offset: number) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
});

/** Where the reference sits: the types of the containers that hold it. */
const referenceHolders = (document: Document): string[][] => {
  const found: string[][] = [];
  const visit = (items: readonly ParagraphContent[], path: string[]) => {
    for (const item of items) {
      if (item.type === "commentReference") found.push(path);
      else if (
        item.type === "insertion" ||
        item.type === "deletion" ||
        item.type === "moveFrom" ||
        item.type === "moveTo"
      )
        visit(item.content, [...path, item.type]);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") visit(block.content, []);
  return found;
};

const OPS: Record<string, DocumentOp[]> = {
  split: [0, 1, 2, 3, 4, 5].map((offset) => ({
    type: DOCUMENT_OP_TYPES.SPLIT_BLOCK,
    at: at(offset),
    newBlockId: "0000ABCD",
    newIds: { revision: [90] },
  })),
  insert: [0, 1, 2, 3, 4, 5].map((offset) => ({
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: at(offset),
    slice: { content: [run("X")], openStart: 0, openEnd: 0 },
    newIds: { revision: [90] },
  })),
  delete: [
    [0, 1],
    [1, 2],
    [3, 4],
    [4, 5],
  ].map(([from = 0, to = 0]) => ({
    type: DOCUMENT_OP_TYPES.DELETE_RANGE,
    from: at(from),
    to: at(to),
  })),
};

describe("a point comment reference inside a tracked insertion", () => {
  for (const [name, ops] of Object.entries(OPS))
    test(`${name} keeps it inside the insertion and undoes exactly`, () => {
      for (const op of ops) {
        const result = applyDocumentOp(documentOf(), op);
        if (result.isErr()) throw result.error;
        const { document, inverse } = result.value;
        expect(referenceHolders(document)).toEqual([["insertion"]]);
        expect(
          document.package.document.content.some(
            (block) =>
              block.type === "paragraph" &&
              JSON.stringify(block.content).includes('"text":"X"'),
          ) ===
            (op.type === DOCUMENT_OP_TYPES.INSERT_CONTENT),
        ).toBe(true);
        const undone = applyDocumentOps(document, inverse);
        if (undone.isErr()) throw undone.error;
        expect(undone.value.document).toEqual(documentOf());
      }
    });
});
