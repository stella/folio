/**
 * Edits that join paragraphs report each removed paragraph with the one its
 * content went into: a join keeps the second paragraph's id.
 */
import { expect, spyOn, test } from "bun:test";

import { assertExactModel } from "../../../../../test/exactModel";
import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { combinedAbsorptions } from "../edits";
import * as identities from "../ids";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const deleted = (paraId: string, text: string, id: number): Paragraph => ({
  ...paragraph(paraId, text),
  pPrMark: { kind: "del", info: { id, author: "A" } },
});
const documentOf = (content: Paragraph[]): Document => ({
  package: { document: { content } },
});
const apply = (document: Document, op: DocumentOp) => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value;
};
const resolve = (revisionIds: number[]): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
  story: OP_STORIES.MAIN,
  revisionIds,
  decision: REVISION_DECISIONS.ACCEPT,
});

test("accepting a deleted mark reports its paragraph joined into the next one", () => {
  const applied = apply(
    documentOf([deleted("00000001", "One ", 1), paragraph("00000002", "Two")]),
    resolve([1]),
  );
  expect(applied.touched.removed).toEqual(["00000001"]);
  expect(applied.touched.absorbed).toEqual([
    { absorbed: "00000001", survivor: "00000002" },
  ]);
});

test("a chain of deleted marks reports every removed paragraph against the paragraph left", () => {
  const applied = apply(
    documentOf([
      deleted("00000001", "One ", 1),
      deleted("00000002", "Two ", 2),
      paragraph("00000003", "Three"),
    ]),
    resolve([1, 2]),
  );
  expect(
    applied.touched.absorbed?.toSorted((left, right) =>
      left.absorbed.localeCompare(right.absorbed),
    ),
  ).toEqual([
    { absorbed: "00000001", survivor: "00000003" },
    { absorbed: "00000002", survivor: "00000003" },
  ]);
});

test("a direct join reports the paragraph it removed", () => {
  const applied = apply(
    documentOf([paragraph("00000001", "One "), paragraph("00000002", "Two")]),
    {
      type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      nextBlockId: "00000002",
    },
  );
  const [gone] = applied.touched.removed;
  expect(gone).toBeDefined();
  expect(applied.touched.absorbed).toEqual([
    {
      absorbed: gone ?? "",
      survivor: gone === "00000001" ? "00000002" : "00000001",
    },
  ]);
});

test("deleting one's own inserted paragraph reports its join into the next one", () => {
  const own: Paragraph = {
    ...paragraph("00000001", "Added"),
    pPrMark: { kind: "ins", info: { id: 9, author: "Reviewer" } },
  };
  const applied = apply(documentOf([own, paragraph("00000002", "Kept")]), {
    type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
    story: OP_STORIES.MAIN,
    blockIds: ["00000001"],
    revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
    newIds: { revision: [101, 102, 103] },
  });
  expect(applied.touched.absorbed).toEqual([
    { absorbed: "00000001", survivor: "00000002" },
  ]);
});

test("absorptions made in turn chain to the paragraph holding them now", () => {
  expect(
    combinedAbsorptions(
      [
        [{ absorbed: "A", survivor: "B" }],
        [{ absorbed: "B", survivor: "C" }],
        [{ absorbed: "D", survivor: "E" }],
      ],
      new Set(["A", "B", "D", "E"]),
      new Set(["C", "F"]),
    ),
  ).toEqual([
    { absorbed: "A", survivor: "C" },
    { absorbed: "B", survivor: "C" },
  ]);
});

test("absorptions chain whatever case an id is written in, and name only a survivor that stands", () => {
  expect(
    combinedAbsorptions(
      [
        [{ absorbed: "0000000A", survivor: "0000000b" }],
        [{ absorbed: "0000000B", survivor: "0000000C" }],
        [{ absorbed: "0000000D", survivor: "0000000E" }],
      ],
      new Set(["0000000A", "0000000B", "0000000D"]),
      new Set(["0000000C"]),
    ),
  ).toEqual([
    { absorbed: "0000000A", survivor: "0000000C" },
    { absorbed: "0000000B", survivor: "0000000C" },
  ]);
});

test("a retired ID can be reused after its earlier joins have moved on", () => {
  expect(
    combinedAbsorptions(
      [
        [
          { absorbed: "AA", survivor: "BB" },
          { absorbed: "bb", survivor: "CC" },
          { absorbed: "BB", survivor: "DD" },
        ],
      ],
      new Set(["aa", "bb"]),
      new Set(["cc", "dd"]),
    ),
  ).toEqual([
    { absorbed: "AA", survivor: "CC" },
    { absorbed: "BB", survivor: "DD" },
  ]);
});

test("resolving four times as many chained marks takes linear ID comparisons", () => {
  const keys = spyOn(identities, "idKey");
  const work = (count: number) => {
    const id = (index: number) => index.toString(16).padStart(8, "0");
    const list = Array.from({ length: count }, (_, index) => ({
      absorbed: id(index),
      survivor: id(index + 1),
    }));
    keys.mockClear();
    const result = combinedAbsorptions(
      [list],
      new Set(list.map(({ absorbed }) => absorbed)),
      new Set([id(count)]),
    );
    expect(result).toHaveLength(count);
    expect(result.every(({ survivor }) => survivor === id(count))).toBe(true);
    return keys.mock.calls.length;
  };
  try {
    const small = work(16);
    expect(small).toBeGreaterThan(0);
    expect(work(64)).toBeLessThanOrEqual(5 * small);
  } finally {
    keys.mockRestore();
  }
});
const join = (blockId: string, nextBlockId: string): DocumentOp => ({
  type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
  story: OP_STORIES.MAIN,
  blockId,
  nextBlockId,
});

test.each([16, 64])(
  "combining joins lists no package paragraphs at width %i",
  (width) => {
    const id = (index: number) => identities.formatParaId(index + 1);
    let document = documentOf(
      Array.from({ length: width }, (_, index) => paragraph(id(index), "Text")),
    );
    const packageIds = identities.packageParagraphIds;
    let enumerated = 0;
    const census = spyOn(identities, "packageParagraphIds").mockImplementation(
      (pkg) => {
        const listed = packageIds(pkg);
        enumerated += listed.length;
        return listed;
      },
    );
    try {
      for (let index = 0; index < 4; index += 1) {
        const result = applyDocumentOps(document, [
          join(id(index), id(index + 1)),
        ]);
        if (result.isErr()) throw result.error;
        expect(result.value.touched.absorbed).toEqual([
          { absorbed: id(index), survivor: id(index + 1) },
        ]);
        document = result.value.document;
      }
      expect({ calls: census.mock.calls.length, enumerated }).toEqual({
        calls: 0,
        enumerated: 0,
      });
    } finally {
      census.mockRestore();
    }
  },
);

test("combined joins keep the first destination across mixed-case ID reuse and undo exactly", () => {
  const ids = {
    first: "0000000a",
    reused: "0000000B",
    firstSurvivor: "0000000C",
    lastSurvivor: "0000000d",
    retained: "0000000e",
    final: "0000000f",
  };
  const document = documentOf(
    Object.entries(ids).map(([name, id]) => paragraph(id, name)),
  );
  const insertBefore = (blockId: string, inserted: string): DocumentOp => ({
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "before", blockId },
    blocks: [paragraph(inserted, "reused")],
  });
  const ops = [
    join(ids.first.toUpperCase(), ids.reused.toLowerCase()),
    join(ids.reused.toLowerCase(), ids.firstSurvivor.toLowerCase()),
    insertBefore(ids.lastSurvivor.toUpperCase(), ids.reused.toLowerCase()),
    join(ids.reused, ids.lastSurvivor.toUpperCase()),
    join(ids.retained.toUpperCase(), ids.final.toUpperCase()),
    insertBefore(ids.final, ids.retained.toUpperCase()),
  ];
  let current = document;
  const steps = ops.map((op) => {
    const edit = apply(current, op);
    current = edit.document;
    return edit;
  });
  const result = applyDocumentOps(document, ops);
  if (result.isErr()) throw result.error;
  assertExactModel(result.value.document, current);
  assertExactModel(
    result.value.inverse,
    steps.toReversed().flatMap(({ inverse }) => inverse),
  );
  expect(result.value.touched.absorbed).toEqual([
    { absorbed: ids.first, survivor: ids.firstSurvivor.toLowerCase() },
    {
      absorbed: ids.reused.toLowerCase(),
      survivor: ids.lastSurvivor.toUpperCase(),
    },
  ]);
  expect(result.value.touched.absorbed).toEqual(
    combinedAbsorptions(
      steps.map(({ touched }) => touched.absorbed ?? []),
      new Set(result.value.touched.removed),
      new Set(identities.packageParagraphIds(current.package)),
    ),
  );
  const restored = applyDocumentOps(
    result.value.document,
    result.value.inverse,
  );
  if (restored.isErr()) throw restored.error;
  assertExactModel(restored.value.document, document);
});
