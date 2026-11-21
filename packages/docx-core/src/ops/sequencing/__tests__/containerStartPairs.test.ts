import { expect, test } from "bun:test";

import type { Document, Paragraph } from "../../../model/document";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
} from "../../types";
import { applyBatch } from "../applyBatch";
import { BATCH_REJECTION_REASONS, type DocumentBatch } from "../envelope";
import { transformBatch } from "../transform";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A body that starts with an empty content control. */
const base = {
  package: {
    document: {
      content: [
        { type: "blockSdt", properties: { sdtType: "richText" }, content: [] },
        paragraph("10000005", "Kept"),
        paragraph("10000006", "Elsewhere"),
      ],
    },
  },
} satisfies Document;

const batch = (opId: string, op: InsertBlocksOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops: [op],
});

const atStart = (opId: string, paraId: string) =>
  batch(opId, {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "containerStart", blockId: "10000005" },
    blocks: [paragraph(paraId, opId)],
  });

const rebase = (
  first: DocumentBatch,
  second: DocumentBatch,
  order: "before" | "after",
) => {
  const applied = applyBatch(base, first.ops).unwrap();
  return transformBatch(
    second,
    [{ ...first, revision: 1, effects: applied.effects }],
    { order },
  );
};

test.each(["before", "after"] as const)(
  "two concurrent insertions at one list's start are refused as a block conflict (%s)",
  (order) => {
    const rebased = rebase(
      atStart("first", "00001001"),
      atStart("second", "00001002"),
      order,
    );
    expect(rebased.isErr() && rebased.error.reason).toBe(
      BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
    );
  },
);

test.each(["before", "after"] as const)(
  "an insertion at a list's start commutes with one beside another paragraph (%s)",
  (order) => {
    const elsewhere = batch("elsewhere", {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "10000006" },
      blocks: [paragraph("00001003", "elsewhere")],
    });
    for (const [first, second] of [
      [atStart("start", "00001001"), elsewhere],
      [elsewhere, atStart("start", "00001001")],
    ] as const) {
      const rebased = rebase(first, second, order).unwrap();
      expect(rebased.ops).toEqual(second.ops);
    }
  },
);
