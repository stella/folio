import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { normalizeForOps, validateOpsDocument } from "../contract";
import { rangeBoundaryDefinitions } from "../rangeBoundaries";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import type { DocumentBatch, SequencedBatch } from "../sequencing/envelope";
import { transformBatch } from "../sequencing/transform";
import {
  DOCUMENT_OP_SCHEMA_VERSION,
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
} from "../types";

type BoundaryDefinitions = ReturnType<typeof rangeBoundaryDefinitions>;
type BoundaryFamily = BoundaryDefinitions[keyof BoundaryDefinitions]["family"];
const ranges = {
  bookmark: {
    start: { type: "bookmarkStart", id: 22, name: "_Ref22" },
    end: { type: "bookmarkEnd", id: 22 },
  },
  comment: {
    start: { type: "commentRangeStart", id: 22 },
    end: { type: "commentRangeEnd", id: 22 },
  },
  moveFrom: {
    start: { type: "moveFromRangeStart", id: 22, name: "move22", author: "A" },
    end: { type: "moveFromRangeEnd", id: 22 },
  },
  moveTo: {
    start: { type: "moveToRangeStart", id: 22, name: "move22", author: "A" },
    end: { type: "moveToRangeEnd", id: 22 },
  },
} as const satisfies Record<
  BoundaryFamily,
  { start: ParagraphContent; end: ParagraphContent }
>;

type RangePart = "start" | "end" | "pair";
type FixtureOptions = {
  range: (typeof ranges)[BoundaryFamily];
  source: "empty" | RangePart;
  inserted: RangePart;
  container: "paragraph" | "wrapper";
};

const fixture = ({ range, source, inserted, container }: FixtureOptions) => {
  const parts = {
    empty: [],
    start: [range.start],
    end: [range.end],
    pair: [range.start, range.end],
  };
  const standing = parts[source];
  const content = [
    { type: "run", content: [{ type: "text", text: "a" }] },
    ...standing,
    { type: "run", content: [{ type: "text", text: "b" }] },
  ] satisfies ParagraphContent[];
  const document = normalizeForOps({
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              { type: "run", content: [{ type: "text", text: "other" }] },
            ],
          },
          {
            type: "paragraph",
            paraId: "00000002",
            content:
              container === "paragraph"
                ? content
                : [
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      direction: "rtl",
                      content,
                    },
                  ],
          },
        ],
        comments: [
          {
            id: 22,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
        ],
      },
    },
  } satisfies Document);
  const zeroWidthBefore =
    source === "start" || (source === "pair" && inserted !== "start")
      ? standing.length
      : 0;
  const operation = {
    type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
    at: {
      story: OP_STORIES.MAIN,
      blockId: "00000002",
      offset: 1,
      zeroWidthBefore,
    },
    slice: { content: parts[inserted], openStart: 0, openEnd: 0 },
  } as const satisfies DocumentOp;
  return { document, operation };
};

const elsewhere = {
  type: DOCUMENT_OP_TYPES.INSERT_TEXT,
  at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
  text: "x",
} as const satisfies DocumentOp;

test("inserted range members restore source ownership through exact and JSON inverses and batches", () => {
  const exercised = new Set<string>();
  for (const [family, range] of Object.entries(ranges)) {
    exercised.add(family);
    for (const source of ["empty", "start", "end", "pair"] as const)
      for (const inserted of ["start", "end", "pair"] as const)
        for (const container of ["paragraph", "wrapper"] as const) {
          const { document, operation } = fixture({
            range,
            source,
            inserted,
            container,
          });
          validateOpsDocument(document).unwrap();
          const original = structuredClone(document);
          const edit = applyDocumentOp(document, operation).unwrap();
          expect(edit.inverse.map(({ type }) => type)).toEqual([
            DOCUMENT_OP_TYPES.REPLACE_INLINE,
          ]);
          const undo = applyDocumentOps(edit.document, edit.inverse).unwrap();
          expect(undo.document).toStrictEqual(original);
          expect(
            applyDocumentOps(undo.document, undo.inverse).unwrap().document,
          ).toStrictEqual(edit.document);
          expect(
            applyDocumentOps(
              edit.document,
              JSON.parse(JSON.stringify(edit.inverse)),
            ).unwrap().document,
          ).toStrictEqual(original);
          for (const operations of [
            [operation, elsewhere],
            [elsewhere, operation],
          ]) {
            let current = document;
            const inverses: (readonly DocumentOp[])[] = [];
            for (const op of operations) {
              const applied = applyDocumentOp(current, op).unwrap();
              current = applied.document;
              inverses.push(applied.inverse);
            }
            expect(
              applyDocumentOps(current, inverses.toReversed().flat()).unwrap()
                .document,
            ).toStrictEqual(original);
            const batch = applyDocumentOps(document, operations).unwrap();
            expect(batch.document).toStrictEqual(current);
            expect(
              applyDocumentOps(batch.document, batch.inverse).unwrap().document,
            ).toStrictEqual(original);
            expect(
              applyDocumentOps(
                batch.document,
                JSON.parse(JSON.stringify(batch.inverse)),
              ).unwrap().document,
            ).toStrictEqual(original);
          }
          expect(document).toStrictEqual(original);
        }
  }
  expect([...exercised].sort()).toEqual(
    [
      ...new Set(
        Object.values(rangeBoundaryDefinitions()).map(({ family }) => family),
      ),
    ].sort(),
  );
});

test("an inserted endpoint's exact inverse refuses changed source content and commutes with another paragraph", () => {
  for (const range of Object.values(ranges)) {
    const { document, operation } = fixture({
      range,
      source: "start",
      inserted: "end",
      container: "paragraph",
    });
    const edit = applyDocumentOp(document, operation).unwrap();
    const sameParagraph = {
      ...elsewhere,
      at: { ...elsewhere.at, blockId: "00000002" },
    };
    const changed = applyDocumentOp(edit.document, sameParagraph).unwrap()
      .document;
    const snapshot = structuredClone(changed);
    const refused = applyDocumentOps(changed, edit.inverse);
    expect(refused.isErr()).toBe(true);
    if (refused.isOk())
      throw new TypeError("An exact inverse overwrote concurrent content");
    expect(refused.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.STALE);
    expect(changed).toStrictEqual(snapshot);
    const undoBatch = {
      schema: DOCUMENT_OP_SCHEMA_VERSION,
      actor: "actor",
      opId: "undo",
      baseRev: 0,
      ops: edit.inverse,
    } as const satisfies DocumentBatch;
    const overlapping = {
      type: DOCUMENT_OP_TYPES.SET_RUN_PROPS,
      from: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 2 },
      patch: { bold: true },
    } as const satisfies DocumentOp;
    const formatted = applyDocumentOp(edit.document, overlapping).unwrap()
      .document;
    expect(applyDocumentOps(formatted, edit.inverse).isErr()).toBe(true);
    const concurrent = {
      ...undoBatch,
      actor: "other",
      opId: "concurrent",
      revision: 1,
      ops: [overlapping],
    } as const satisfies SequencedBatch;
    expect(transformBatch(undoBatch, [concurrent]).isErr()).toBe(true);
    const sameParagraphUndo = transformBatch(undoBatch, [
      { ...concurrent, ops: [sameParagraph] },
    ]).unwrap();
    expect(
      applyDocumentOps(changed, sameParagraphUndo.ops).unwrap().document,
    ).toStrictEqual(applyDocumentOp(document, sameParagraph).unwrap().document);
    const otherEdit = applyDocumentOp(edit.document, elsewhere).unwrap();
    const transformed = transformBatch(undoBatch, [
      { ...concurrent, ops: [elsewhere] },
    ]).unwrap();
    expect(transformed.ops).toStrictEqual(edit.inverse);
    expect(
      applyDocumentOps(otherEdit.document, transformed.ops).unwrap().document,
    ).toStrictEqual(applyDocumentOp(document, elsewhere).unwrap().document);
  }
});

test("text-only insertions beside source boundaries keep compact deletion inverses", () => {
  for (const range of Object.values(ranges)) {
    const { document, operation } = fixture({
      range,
      source: "start",
      inserted: "end",
      container: "wrapper",
    });
    const content = [
      { type: "run", content: [{ type: "text", text: "x" }] },
    ] satisfies ParagraphContent[];
    for (const op of [
      { ...operation, slice: { content, openStart: 0, openEnd: 0 } },
      { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at: operation.at, text: "x" },
    ] as const satisfies readonly DocumentOp[]) {
      const edit = applyDocumentOp(document, op).unwrap();
      expect(edit.inverse.map(({ type }) => type)).toEqual([
        DOCUMENT_OP_TYPES.DELETE_RANGE,
      ]);
      expect(
        applyDocumentOps(edit.document, edit.inverse).unwrap().document,
      ).toStrictEqual(document);
    }
  }
});

test("source-cut provenance already produces exact restoration inverses for both insertion kinds", () => {
  const source = normalizeForOps({
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content: [] }],
      },
    },
  } satisfies Document);
  const pending = applyDocumentOp(source, {
    ...elsewhere,
    text: "ab",
    revision: { id: 7, author: "A" },
  }).unwrap().document;
  const at = { ...elsewhere.at, offset: 1 };
  const newIds = { revision: [8, 9, 10, 11] };
  for (const operation of [
    { type: DOCUMENT_OP_TYPES.INSERT_TEXT, at, text: "x", newIds },
    {
      type: DOCUMENT_OP_TYPES.INSERT_CONTENT,
      at,
      slice: {
        content: [{ type: "run", content: [{ type: "text", text: "x" }] }],
        openStart: 0,
        openEnd: 0,
      },
      newIds,
    },
  ] as const satisfies readonly DocumentOp[]) {
    const edit = applyDocumentOp(pending, operation).unwrap();
    expect(edit.inverse.map(({ type }) => type)).toEqual([
      DOCUMENT_OP_TYPES.REPLACE_INLINE,
    ]);
    expect(
      applyDocumentOps(edit.document, edit.inverse).unwrap().document,
    ).toStrictEqual(pending);
  }
});
