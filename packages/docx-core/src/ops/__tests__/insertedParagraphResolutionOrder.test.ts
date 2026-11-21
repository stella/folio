/**
 * The changes a tracked insertion of whole paragraphs records (each
 * paragraph's words and its break) resolve one at a time, in any order and
 * with any mix of accepts and rejects, to what resolving the accepted ones
 * and the rejected ones each at once gives.
 */
import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type InsertBlocksOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const orderings = (ids: readonly number[]): number[][] =>
  ids.length <= 1
    ? [[...ids]]
    : ids.flatMap((id, index) =>
        orderings(ids.filter((_, other) => other !== index)).map((rest) =>
          [id].concat(rest),
        ),
      );

type Step = { ids: readonly number[]; decision: RevisionDecision };

const resolveSteps = (document: Document, steps: readonly Step[]) => {
  let current = document;
  for (const { ids, decision } of steps) {
    if (ids.length === 0) continue;
    const resolved = applyDocumentOp(current, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: ids,
      decision,
    });
    if (resolved.isErr()) throw resolved.error;
    current = resolved.value.document;
  }
  return current;
};

const SHAPES = [
  {
    name: "between two paragraphs",
    document: documentOf([
      paragraph("00000001", "head"),
      paragraph("00000005", "tail"),
    ]),
  },
  {
    name: "after a cell's last paragraph",
    document: documentOf([
      {
        type: "table",
        rows: [
          {
            type: "tableRow",
            cells: [
              { type: "tableCell", content: [paragraph("00000001", "cell")] },
            ],
          },
        ],
      },
      paragraph("00000005", "tail"),
    ]),
  },
] as const;

for (const { name, document } of SHAPES)
  test.each([1, 2])(
    `%i paragraphs inserted ${name} resolve alike in any order and all at once`,
    (count) => {
      const op = {
        type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
        story: OP_STORIES.MAIN,
        at: { type: "after", blockId: "00000001" },
        blocks: Array.from({ length: count }, (_, index) =>
          paragraph((16 + index).toString(16).padStart(8, "0"), `new ${index}`),
        ),
        revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
        newIds: {
          revision: Array.from({ length: 16 }, (_, index) => 101 + index),
        },
      } satisfies InsertBlocksOp;
      const applied = applyDocumentOp(document, op);
      if (applied.isErr()) throw applied.error;
      const { revisions } = applied.value;
      expect(revisions).toHaveLength(2 * count);
      for (
        let accepting = 0;
        accepting < 1 << revisions.length;
        accepting += 1
      ) {
        const decide = (id: number) =>
          accepting & (1 << revisions.indexOf(id))
            ? REVISION_DECISIONS.ACCEPT
            : REVISION_DECISIONS.REJECT;
        const accepted: Step = {
          ids: revisions.filter(
            (id) => decide(id) === REVISION_DECISIONS.ACCEPT,
          ),
          decision: REVISION_DECISIONS.ACCEPT,
        };
        const rejected: Step = {
          ids: revisions.filter(
            (id) => decide(id) === REVISION_DECISIONS.REJECT,
          ),
          decision: REVISION_DECISIONS.REJECT,
        };
        const atOnce = resolveSteps(applied.value.document, [
          accepted,
          rejected,
        ]);
        expect(
          resolveSteps(applied.value.document, [rejected, accepted]),
        ).toStrictEqual(atOnce);
        for (const order of orderings(revisions))
          expect(
            resolveSteps(
              applied.value.document,
              order.map((id) => ({ ids: [id], decision: decide(id) })),
            ),
          ).toStrictEqual(atOnce);
      }
    },
  );

test("rejecting a break first keeps the next paragraph's words a change of their own", () => {
  const applied = applyDocumentOp(
    documentOf([paragraph("00000001", "head"), paragraph("00000005", "tail")]),
    {
      type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
      story: OP_STORIES.MAIN,
      at: { type: "after", blockId: "00000001" },
      blocks: [paragraph("00000010", "A"), paragraph("00000011", "B")],
      revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
      newIds: { revision: [101, 102, 103] },
    },
  );
  if (applied.isErr()) throw applied.error;
  // The first paragraph's break, then its words and the second's break are
  // rejected; the second paragraph's words are accepted and stay.
  const [firstBreak, firstWords, secondBreak, secondWords] =
    applied.value.revisions;
  if (secondWords === undefined) throw new Error("Four changes are recorded");
  const resolved = resolveSteps(applied.value.document, [
    { ids: [firstBreak ?? 0], decision: REVISION_DECISIONS.REJECT },
    { ids: [firstWords ?? 0], decision: REVISION_DECISIONS.REJECT },
    { ids: [secondBreak ?? 0], decision: REVISION_DECISIONS.REJECT },
    { ids: [secondWords], decision: REVISION_DECISIONS.ACCEPT },
  ]);
  expect(resolved.package.document.content).toStrictEqual([
    paragraph("00000001", "head"),
    paragraph("00000005", "Btail"),
  ]);
});
