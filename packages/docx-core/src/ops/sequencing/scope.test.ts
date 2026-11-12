import { describe, expect, test } from "bun:test";

import type { Document, Paragraph, TableRow } from "../../model/document";
import { envelopes } from "../__tests__/wireFixtures";
import { applyDocumentOps } from "../apply";
import { equalForStaleness } from "../equality";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
} from "../types";
import { applyBatch } from "./applyBatch";
import {
  BATCH_REJECTION_REASONS,
  type DocumentBatch,
  type SequencedBatch,
  sequencedEffectType,
} from "./envelope";
import { isBlockScoped } from "./scope";
import { transformBatch } from "./transform";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const row = (ids: readonly string[]): TableRow => ({
  type: "tableRow",
  cells: ids.map((id) => ({ type: "tableCell", content: [paragraph(id, id)] })),
});
const documentOf = (): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "abc"),
        {
          type: "table",
          columnWidths: [2000, 2000],
          rows: [row(["00000011", "00000012"]), row(["00000021", "00000022"])],
        },
        paragraph("00000002", "def"),
      ],
    },
  },
});
const batch = (opId: string, ...ops: DocumentOp[]): DocumentBatch => ({
  schema: DOCUMENT_OP_SCHEMA_VERSION,
  opId,
  actor: opId,
  baseRev: 0,
  ops,
});
const sequenced = (
  document: Document,
  local: DocumentBatch,
): SequencedBatch => {
  const applied = applyBatch(document, local.ops);
  if (applied.isErr()) throw applied.error;
  return { ...local, revision: 1, effects: applied.value.effects };
};
const apply = (document: Document, ops: readonly DocumentOp[]) => {
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  return result.value.document;
};
/** Both orders of a concurrent pair, or the refusal of the second. */
const pair = (first: DocumentBatch, second: DocumentBatch) => {
  const document = documentOf();
  const after = transformBatch(second, [sequenced(document, first)]);
  const before = transformBatch(first, [sequenced(document, second)], {
    order: "before",
  });
  if (after.isErr() || before.isErr())
    return { refused: after.isErr() ? after.error : before.error };
  const left = apply(apply(document, first.ops), after.value.ops);
  const right = apply(apply(document, second.ops), before.value.ops);
  return { commutes: equalForStaleness(left, right) };
};
const insertRow = (blockId: string, at: number, ids: readonly string[]) =>
  ({
    type: DOCUMENT_OP_TYPES.INSERT_ROW,
    story: "main",
    blockId,
    at,
    row: row(ids),
  }) as const satisfies DocumentOp;
const typing = (blockId: string, offset = 1) =>
  ({
    type: DOCUMENT_OP_TYPES.INSERT_TEXT,
    at: { story: "main", blockId, offset },
    text: "X",
    runProps: "inherit",
  }) as const satisfies DocumentOp;

describe("block-scoped sequencing", () => {
  test("every sendable kind is block-scoped exactly when it records touched paragraphs", () => {
    const seen = new Set<string>();
    for (const { op } of envelopes()) {
      const effect = sequencedEffectType(op.type);
      if (effect === undefined) continue;
      seen.add(op.type);
      expect(isBlockScoped(op)).toBe(
        effect === "touchedBlocks" &&
          op.type !== DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      );
    }
    expect(seen.size).toBeGreaterThan(30);
  });

  test("two structure edits of one table refuse; edits of other paragraphs pass and commute", () => {
    const first = batch(
      "a",
      insertRow("00000011", 1, ["00000031", "00000032"]),
    );
    const second = batch(
      "b",
      insertRow("00000022", 2, ["00000041", "00000042"]),
    );
    expect(pair(first, second).refused?.reason).toBe(
      BATCH_REJECTION_REASONS.TABLE_REQUIRES_EXCLUSIVE_EDIT,
    );
    for (const text of ["00000012", "00000001"]) {
      expect(pair(first, batch("t", typing(text)))).toEqual({ commutes: true });
      expect(pair(batch("t", typing(text)), first)).toEqual({ commutes: true });
    }
  });

  test("an edit of a paragraph a concurrent operation replaced refuses as a block conflict", () => {
    const replace = batch("r", {
      type: DOCUMENT_OP_TYPES.REPLACE_BLOCKS,
      story: "main",
      expected: [paragraph("00000002", "def")],
      blocks: [paragraph("00000002", "pasted")],
    });
    expect(pair(replace, batch("t", typing("00000002"))).refused?.reason).toBe(
      BATCH_REJECTION_REASONS.BLOCK_CONFLICT,
    );
    expect(pair(replace, batch("t", typing("00000001")))).toEqual({
      commutes: true,
    });
  });

  test("numbering creations conflict only on the same instance", () => {
    const create = (numId: number) =>
      batch(`n${String(numId)}`, {
        type: DOCUMENT_OP_TYPES.CREATE_NUMBERING_INSTANCE,
        num: { numId, abstractNumId: numId },
        abstractNum: { abstractNumId: numId, levels: [] },
      });
    const same = transformBatch(create(5), [
      {
        ...create(5),
        revision: 1,
        effects: [{ type: "touchedBlocks", blockIds: [] }],
      },
    ]);
    expect(same.isErr() && same.error.reason).toBe(
      BATCH_REJECTION_REASONS.PART_CONFLICT,
    );
    expect(same.isErr() && same.error.part).toBe("numbering");
    expect(
      transformBatch(create(6), [
        {
          ...create(5),
          revision: 1,
          effects: [{ type: "touchedBlocks", blockIds: [] }],
        },
      ]).isOk(),
    ).toBe(true);
  });

  test("text inherited at a paragraph's start and formatting from the start commute", () => {
    const format = batch("f", {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: { story: "main", blockId: "00000001", offset: 0 },
      to: { story: "main", blockId: "00000001", offset: 2 },
      patch: { bold: true },
    });
    expect(pair(format, batch("t", typing("00000001", 0)))).toEqual({
      commutes: true,
    });
    expect(pair(batch("t", typing("00000001", 0)), format)).toEqual({
      commutes: true,
    });
  });
});
