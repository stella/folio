import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../../model/document";
import { equalForStaleness } from "../../equality";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../../types";
import { applyBatch } from "../applyBatch";
import type { DocumentBatch, SequencedBatch } from "../envelope";
import { transformBatch } from "../transform";

const ANCHOR = "10000005";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/** A body whose last paragraph is the anchor, with another before it. */
const base: Document = {
  package: {
    document: {
      content: [paragraph("10000004", "Before"), paragraph(ANCHOR, "Anchor")],
    },
  },
};

const batchOf = (opId: string, op: DocumentOp): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops: [op],
});

/** The direct insertion, its paragraph formatted unlike the others or not. */
const directOf = (formatted: boolean) =>
  batchOf("direct", {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: "main",
    at: { type: "after", blockId: ANCHOR },
    blocks: [
      formatted
        ? { ...paragraph("00001001", "Direct"), formatting: { keepNext: true } }
        : paragraph("00001001", "Direct"),
    ],
  });
const direct = directOf(false);

const tracked = batchOf("tracked", {
  type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
  story: "main",
  at: { type: "after", blockId: ANCHOR },
  blocks: [paragraph("00000001", "Tracked")],
  revision: { id: 11, author: "actor-0", date: "2026-01-01T00:00:10.000Z" },
  newIds: {
    revision: [12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25],
    control: [],
  },
});

/**
 * `first` sequenced, then `second` rebased over it; `anchored` is the
 * test-only break, which keeps the rebased insertion at the shared anchor
 * as a transform by sequence alone would for the later batch.
 */
const inOrder = (
  first: DocumentBatch,
  second: DocumentBatch,
  anchored = false,
): Document => {
  const applied = applyBatch(base, first.ops);
  if (applied.isErr()) throw applied.error;
  const sequenced: SequencedBatch = {
    ...first,
    revision: 1,
    effects: applied.value.effects,
  };
  const rebased = transformBatch(second, [sequenced]);
  if (rebased.isErr()) throw rebased.error;
  const ops = anchored ? second.ops : rebased.value.ops;
  const after = applyBatch(applied.value.document, ops);
  if (after.isErr()) throw after.error;
  return after.value.document;
};

const summary = (document: Document) =>
  document.package.document.content.map((block) =>
    block.type === "paragraph"
      ? [block.paraId, block.pPrMark?.kind ?? null, block.pPrMark?.survivor]
      : [block.type],
  );

describe("a direct and a tracked paragraph insertion after the last paragraph", () => {
  test("with the direct paragraph formatted unlike the others, the tracked one records its change against it in either order", () => {
    const formatted = directOf(true);
    const directFirst = inOrder(formatted, tracked);
    const trackedFirst = inOrder(tracked, formatted);
    const recorded = (document: Document) =>
      document.package.document.content.map((block) =>
        block.type === "paragraph" ? (block.propertyChanges ?? null) : null,
      );
    expect(recorded(trackedFirst)).toEqual(recorded(directFirst));
    expect(equalForStaleness(trackedFirst, directFirst)).toBe(true);
    // The tracked paragraph holds the last mark: its properties before were
    // the direct paragraph's, the holder of the inserted mark.
    expect(recorded(directFirst).at(-1)).toEqual([
      expect.objectContaining({
        type: "paragraphPropertyChange",
        previousFormatting: { keepNext: true },
      }),
    ]);
  });

  test("leave the same document in either order, the tracked paragraph last", () => {
    const directFirst = inOrder(direct, tracked);
    const trackedFirst = inOrder(tracked, direct);
    expect(summary(trackedFirst)).toEqual(summary(directFirst));
    expect(equalForStaleness(trackedFirst, directFirst)).toBe(true);
    // The inserted break stays right before the tracked paragraph: on the
    // direct paragraph, not on the anchor.
    expect(summary(directFirst)).toEqual([
      ["10000004", null, undefined],
      [ANCHOR, null, undefined],
      ["00001001", "ins", "first"],
      ["00000001", null, undefined],
    ]);
  });

  test("a rebase that keeps the later insertion at the anchor is caught", () => {
    expect(
      equalForStaleness(
        inOrder(direct, tracked, true),
        inOrder(tracked, direct, true),
      ),
    ).toBe(false);
  });
});
