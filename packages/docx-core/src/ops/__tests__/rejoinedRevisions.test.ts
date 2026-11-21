/**
 * A resolution that joins two pieces of one change into one reports the id
 * that no record carries any more with the change that holds its content
 * now, so a host holding the old id decides the survivor.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionAbsorption,
  type RevisionDecision,
} from "../types";

const makeParagraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const SHAPES = {
  body: () =>
    documentOf([
      makeParagraph("00000001", "Head"),
      makeParagraph("00000005", "Tail"),
    ]),
  cell: () =>
    documentOf([
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                type: "tableCell",
                content: [makeParagraph("00000001", "Head")],
              },
            ],
          },
        ],
      },
      makeParagraph("00000005", "Tail"),
    ]),
} as const;

const inserted = (shape: keyof typeof SHAPES, count: number) =>
  applyDocumentOp(SHAPES[shape](), {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    blocks: Array.from({ length: count }, (_, index) =>
      makeParagraph((16 + index).toString(16).padStart(8, "0"), `New ${index}`),
    ),
    revision: { id: 100, author: "R", date: "2026-02-03T04:05:06Z" },
    newIds: { revision: Array.from({ length: 16 }, (_, index) => 101 + index) },
    at: { type: "after", blockId: "00000001" },
  } satisfies InsertBlocksOp).unwrap();

/** Every revision id a record of the document carries. */
const revisionIdsOf = (document: Document): Set<number> => {
  const ids = new Set<number>();
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (value === null || typeof value !== "object") return;
    for (const [key, field] of Object.entries(value)) {
      if (
        key === "info" &&
        field !== null &&
        typeof field === "object" &&
        "id" in field &&
        typeof field.id === "number"
      )
        ids.add(field.id);
      visit(field);
    }
  };
  visit(document.package.document.content);
  return ids;
};

const resolve = (document: Document, id: number, decision: RevisionDecision) =>
  applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [id],
    decision,
  }).unwrap();

test("rejecting the break between two inserted paragraphs reports the words that rejoin", () => {
  const applied = inserted("body", 2);
  const [, first, second] = applied.document.package.document.content;
  if (first?.type !== "paragraph" || second?.type !== "paragraph")
    throw new Error("Two inserted paragraphs follow the head.");
  const firstWords = first.content[0];
  const secondWords = second.content[0];
  if (firstWords?.type !== "insertion" || secondWords?.type !== "insertion")
    throw new Error("Each inserted paragraph holds its words as an insertion.");
  const breakId = first.pPrMark?.info.id ?? -1;
  const rejected = resolve(
    applied.document,
    breakId,
    REVISION_DECISIONS.REJECT,
  );
  expect(rejected.touched.absorbedRevisions).toEqual([
    { absorbed: secondWords.info.id, survivor: firstWords.info.id },
  ] satisfies RevisionAbsorption[]);
  const left = revisionIdsOf(rejected.document);
  expect(left.has(secondWords.info.id)).toBe(false);
  expect(left.has(firstWords.info.id)).toBe(true);
  // The absorbed id names nothing; its survivor decides the rejoined words.
  expect(
    resolve(rejected.document, secondWords.info.id, REVISION_DECISIONS.REJECT)
      .document,
  ).toStrictEqual(rejected.document);
  expect(
    resolve(rejected.document, firstWords.info.id, REVISION_DECISIONS.ACCEPT)
      .document.package.document.content,
  ).toMatchObject([
    { paraId: "00000001" },
    {
      paraId: "00000011",
      content: [{ type: "run", content: [{ text: "New 0New 1" }] }],
    },
    { paraId: "00000005" },
  ]);
});

const orderings = (ids: readonly number[]): number[][] =>
  ids.length <= 1
    ? [[...ids]]
    : ids.flatMap((id, index) =>
        orderings(ids.filter((_, other) => other !== index)).map((rest) =>
          [id].concat(rest),
        ),
      );

test.each(
  (["body", "cell"] as const).flatMap((shape) =>
    [1, 2, 3].map((count) => ({ shape, count })),
  ),
)(
  "every change a single resolution retires without deciding it is reported with its survivor ($shape, $count paragraphs)",
  ({ shape, count }) => {
    const applied = inserted(shape, count);
    const all = [...revisionIdsOf(applied.document)].toSorted((a, b) => a - b);
    expect(all).toHaveLength(2 * count);
    // Every decision assignment up to two paragraphs; all-accept and
    // all-reject for three.
    const full = (1 << all.length) - 1;
    const assignments =
      count < 3
        ? Array.from({ length: full + 1 }, (_, index) => index)
        : [0, full];
    for (const accepting of assignments)
      for (const order of orderings(all)) {
        const survivorOf = new Map<number, number>();
        const target = (id: number): number => {
          let current = id;
          for (let next = survivorOf.get(current); next !== undefined;) {
            current = next;
            next = survivorOf.get(current);
          }
          return current;
        };
        let current = applied.document;
        for (const id of order) {
          const decided = target(id);
          const before = revisionIdsOf(current);
          if (!before.has(decided)) continue;
          const step = resolve(
            current,
            decided,
            accepting & (1 << all.indexOf(id))
              ? REVISION_DECISIONS.ACCEPT
              : REVISION_DECISIONS.REJECT,
          );
          const after = revisionIdsOf(step.document);
          const retired = [...before].filter(
            (other) => other !== decided && !after.has(other),
          );
          const reported = step.touched.absorbedRevisions ?? [];
          expect(reported.map(({ absorbed }) => absorbed).toSorted()).toEqual(
            retired.toSorted(),
          );
          for (const { absorbed, survivor } of reported) {
            expect(after.has(survivor)).toBe(true);
            survivorOf.set(absorbed, survivor);
          }
          current = step.document;
        }
        expect(revisionIdsOf(current).size).toBe(0);
      }
  },
);
