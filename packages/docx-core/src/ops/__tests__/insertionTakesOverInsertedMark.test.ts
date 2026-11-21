import { expect, test } from "bun:test";

import type { Document, Paragraph, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

// A body of one paragraph S. A tracked table after S, as the editor puts
// one at a container's end (an empty paragraph after S, then the table
// before it), makes S's mark a pending insertion. A tracked paragraph P
// after S then takes over that pending mark, and S gets the new operation's
// inserted mark. Each operation reviews on its own: accepting some and
// rejecting the others, in either order or at once, gives the accepted ones
// made directly, and each operation's inverse is exact.

const SOURCE = "00000001";
const FINAL = "00000002";
const ADDED = "00000003";

const run = (text: string) => ({
  type: "run" as const,
  content: [{ type: "text" as const, text }],
});

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: text === "" ? [] : [run(text)],
});

const table: Table = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        { type: "tableCell", content: [paragraph("00000010", "Term")] },
        { type: "tableCell", content: [paragraph("00000011", "Value")] },
      ],
    },
  ],
};

const original: Document = {
  package: { document: { content: [paragraph(SOURCE, "Source")] } },
};

type Tracking = "tracked" | "direct";

/** The review stamp and spare ids of an operation's `index`th op. */
const tracking = (mode: Tracking, index: number) =>
  mode === "direct"
    ? {}
    : {
        revision: {
          id: 100 * (index + 1),
          author: "Reviewer",
          date: `2026-02-03T04:05:0${index}Z`,
        },
        newIds: {
          revision: Array.from(
            { length: 32 },
            (_, offset) => 100 * (index + 1) + offset + 1,
          ),
        },
      };

const TABLE = (mode: Tracking): DocumentOp[] => [
  {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: SOURCE },
    blocks: [paragraph(FINAL, "")],
    ...tracking(mode, 0),
  },
  {
    type: DOCUMENT_OP_TYPES.INSERT_TABLE,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: SOURCE },
    table,
    ...tracking(mode, 1),
  },
];

const PARAGRAPH = (mode: Tracking): DocumentOp[] => [
  {
    type: DOCUMENT_OP_TYPES.INSERT_BLOCKS,
    story: OP_STORIES.MAIN,
    at: { type: "after", blockId: SOURCE },
    blocks: [paragraph(ADDED, "Added")],
    ...tracking(mode, 2),
  },
];

const STEPS = [TABLE, PARAGRAPH] as const;

type Applied = {
  document: Document;
  inverse: readonly DocumentOp[];
  revisions: readonly number[];
};

/** One operation's ops, one at a time, as the editor applies them. */
const apply = (document: Document, ops: readonly DocumentOp[]): Applied => {
  let current = document;
  const inverse: DocumentOp[][] = [];
  const revisions: number[] = [];
  for (const op of ops) {
    const result = applyDocumentOp(current, op);
    if (result.isErr()) throw result.error;
    current = result.value.document;
    inverse.unshift([...result.value.inverse]);
    revisions.push(...result.value.revisions);
  }
  return { document: current, inverse: inverse.flat(), revisions };
};

const undo = (applied: Applied) => {
  const restored = applyDocumentOps(applied.document, applied.inverse);
  if (restored.isErr()) throw restored.error;
  return restored.value.document;
};

const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  apply(document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds,
      decision,
    },
  ]).document;

const afterTable = apply(original, TABLE("tracked"));
const afterParagraph = apply(afterTable.document, PARAGRAPH("tracked"));
const OWN = [afterTable.revisions, afterParagraph.revisions] as const;

const paragraphs = (document: Document): Paragraph[] =>
  document.package.document.content.flatMap((block) =>
    block.type === "paragraph" ? [block] : [],
  );

/** The text under a node, pending or not. */
const textOf = (node: unknown): string => {
  if (typeof node !== "object" || node === null) return "";
  if ("text" in node && typeof node.text === "string") return node.text;
  return "content" in node && Array.isArray(node.content)
    ? node.content.map(textOf).join("")
    : "";
};

/** Each body block's text, or `table`, without ids or review records. */
const read = (document: Document): string[] =>
  document.package.document.content.map((block) =>
    block.type === "table" ? "table" : textOf(block),
  );

test("the paragraph takes over the source's pending inserted mark and the source gets the new one", () => {
  expect(read(afterTable.document)).toEqual(["Source", "table", ""]);
  const mark = paragraphs(afterTable.document).at(0)?.pPrMark;
  expect(mark).toMatchObject({ kind: "ins", survivor: "first" });
  expect(OWN[0]).toContain(mark?.info.id ?? -1);
  expect(read(afterParagraph.document)).toEqual([
    "Source",
    "Added",
    "table",
    "",
  ]);
  const [source, added] = paragraphs(afterParagraph.document);
  // The pending mark, its id and its rejection survivor, moved to P.
  expect(added?.paraId).toBe(ADDED);
  expect(added?.pPrMark).toEqual(mark);
  expect(source?.paraId).toBe(SOURCE);
  const own = source?.pPrMark?.info.id ?? -1;
  expect(source?.pPrMark).toEqual({
    kind: "ins",
    info: { id: own, author: "Reviewer", date: "2026-02-03T04:05:02Z" },
    survivor: "first",
  });
  expect(OWN[1]).toContain(own);
  expect(OWN[0]).not.toContain(own);
});

test("each operation's inverse is exact", () => {
  expect(undo(afterParagraph)).toEqual(afterTable.document);
  expect(undo(afterTable)).toEqual(original);
});

const DECISIONS = Object.values(REVISION_DECISIONS);

const directOf = (accepted: readonly boolean[]) => {
  let document = original;
  for (const [index, step] of STEPS.entries())
    if (accepted[index] === true)
      document = apply(document, step("direct")).document;
  return document;
};

for (const first of DECISIONS)
  for (const second of DECISIONS)
    test(`to ${first} the table and ${second} the paragraph, in either order or at once, gives the accepted ones made directly`, () => {
      const accepted = [first, second].map(
        (decision) => decision === REVISION_DECISIONS.ACCEPT,
      );
      const expected = read(directOf(accepted));
      const inOrder = resolve(
        resolve(afterParagraph.document, OWN[0], first),
        OWN[1],
        second,
      );
      const reversed = resolve(
        resolve(afterParagraph.document, OWN[1], second),
        OWN[0],
        first,
      );
      expect(read(inOrder)).toEqual(expected);
      expect(reversed).toEqual(inOrder);
      if (first === second)
        expect(
          resolve(afterParagraph.document, [...OWN[0], ...OWN[1]], first),
        ).toEqual(inOrder);
    });

test("rejecting the paragraph alone gives the table operation's state", () => {
  expect(
    resolve(afterParagraph.document, OWN[1], REVISION_DECISIONS.REJECT),
  ).toEqual(afterTable.document);
});

test("rejecting the table alone joins the final paragraph into the new one, which keeps its id", () => {
  const rejected = resolve(
    afterParagraph.document,
    OWN[0],
    REVISION_DECISIONS.REJECT,
  );
  expect(read(rejected)).toEqual(["Source", "Added"]);
  expect(paragraphs(rejected).map(({ paraId }) => paraId)).toEqual([
    SOURCE,
    ADDED,
  ]);
});

test("rejecting everything gives the original back", () => {
  for (const order of [
    [...OWN[0], ...OWN[1]],
    [...OWN[1], ...OWN[0]],
  ])
    expect(
      resolve(afterParagraph.document, order, REVISION_DECISIONS.REJECT),
    ).toEqual(original);
});
