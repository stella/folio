/**
 * A break inserted inside a tracked deletion leaves a piece of it either
 * side of the new mark, each carrying the deletion's joins. The join a piece
 * holds on the side facing the mark copies the deletion's far edge, so
 * rejecting the deletion and the break at once joins only what rejecting
 * them one after another, in any order, does: the deletion's own seam.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const NEW_IDS = {
  revision: [700, 701, 702, 703, 704, 705, 706, 707],
  control: [800, 801, 802],
};

const apply = (document: Document, op: DocumentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const resolve = (
  document: Document,
  ids: readonly number[],
  decision: RevisionDecision,
): Document =>
  apply(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  });

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const link = (children: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});

const revision = (id: number, second: number) => ({
  id,
  author: "C",
  date: `2026-03-04T05:00:0${second}Z`,
});

/** The ids of the tracked records and marks dated `second`, pieces included. */
const idsDated = (document: Document, second: number): number[] => {
  const ids: number[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const info: unknown = Reflect.get(value, "info");
    if (
      typeof info === "object" &&
      info !== null &&
      Reflect.get(info, "date") === revision(0, second).date
    )
      ids.push(Number(Reflect.get(info, "id")));
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return [...new Set(ids)];
};

test("a break inside a rejected deletion joins only the deletion's own seam, at once or in any order", () => {
  const original: ParagraphContent[] = [
    link([text("a")]),
    { type: "bookmarkEnd", id: 1 },
    link([text("bc")]),
  ];
  const source: Document = {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content: original }],
      },
    },
  };
  // The deletion cuts the second link; the break then cuts the deletion.
  const deleted = apply(source, {
    type: "deleteRange",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
    newIds: NEW_IDS,
    revision: revision(5001, 0),
  });
  const tracked = apply(deleted, {
    type: "splitBlock",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    newBlockId: "10000000",
    newIds: NEW_IDS,
    revision: revision(5002, 1),
  });
  const perOp = [0, 1].map((second) => idsDated(tracked, second));
  expect(perOp.map((ids) => ids.length > 0)).toEqual([true, true]);
  expect(tracked.package.document.content).toHaveLength(2);
  const expected = [
    { type: "paragraph", paraId: "00000001", content: original },
  ];
  expect(
    resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT).package.document
      .content,
  ).toEqual(expected);
  // The deletion first, then the break; and the other way round.
  for (const order of [perOp, perOp.toReversed()]) {
    let inTurn = tracked;
    for (const ids of order)
      inTurn = resolve(inTurn, ids, REVISION_DECISIONS.REJECT);
    expect(inTurn.package.document.content).toEqual(expected);
  }
});
