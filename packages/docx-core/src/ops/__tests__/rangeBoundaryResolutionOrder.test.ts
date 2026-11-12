/**
 * A range boundary that rejected content leaves stays where that content was,
 * whether the run's revisions are rejected at once or one operation's at a
 * time: the seams the content recorded are read against the records that
 * remain, never against records a resolution removes or that it never cut.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { childNodes, type InlineNode } from "../leaves";
import { isTrackedWrapper } from "../review";
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

const DATES = [
  "2026-03-04T05:00:00Z",
  "2026-03-04T05:00:01Z",
  "2026-03-04T05:00:02Z",
  "2026-03-04T05:00:03Z",
] as const;

const date = (index: number): string => DATES[index] ?? "2026-03-04T05:00:04Z";

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
  ids.length === 0
    ? document
    : apply(document, {
        type: "resolveRevision",
        story: OP_STORIES.MAIN,
        revisionIds: ids,
        decision,
      });

/** The ids of the tracked records an operation dated `when` left, pieces included. */
const idsDated = (document: Document, when: string): number[] => {
  const ids = new Set<number>();
  const visitInline = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node) && node.info.date === when)
        ids.add(node.info.id);
      visitInline(childNodes(node) ?? []);
    }
  };
  const visitBlocks = (blocks: readonly BlockContent[]): void => {
    for (const block of blocks) {
      if (block.type === "paragraph") visitInline(block.content);
      else if (block.type === "blockSdt" || block.type === "blockCustomXml")
        visitBlocks(block.content);
    }
  };
  visitBlocks(document.package.document.content);
  return [...ids];
};

const dated = (document: Document, count: number): number[][] =>
  Array.from({ length: count }, (_, index) => idsDated(document, date(index)));

const body = (document: Document) => document.package.document.content;

const opsOn = (
  document: Document,
  ops: readonly ((index: number) => DocumentOp)[],
): Document => {
  let current = document;
  for (const [index, op] of ops.entries()) current = apply(current, op(index));
  return current;
};

/** Each operation's revisions rejected in turn, from the last. */
const rejectedInTurn = (
  document: Document,
  perOp: readonly (readonly number[])[],
): Document => {
  let current = document;
  for (const ids of perOp.toReversed())
    current = resolve(current, ids, REVISION_DECISIONS.REJECT);
  return current;
};

/** Every order of `items`. */
const orders = <T>(items: readonly T[]): T[][] =>
  items.length <= 1
    ? [[...items]]
    : items.flatMap((item, index) =>
        orders(items.toSpliced(index, 1)).map((rest) => [item].concat(rest)),
      );

/** Each operation's revisions rejected one after another, in `order`. */
const rejectedInOrder = (
  document: Document,
  order: readonly (readonly number[])[],
): Document => {
  let current = document;
  for (const ids of order)
    current = resolve(current, ids, REVISION_DECISIONS.REJECT);
  return current;
};

const revision = (index: number) => ({
  id: index + 1,
  author: "C",
  date: date(index),
});

const tab = (): Run => ({ type: "run", content: [{ type: "tab" }] });

test("rejecting content inserted inside a pending insertion joins the insertion's pieces again", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 1, tag: "clause" },
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "text", text: "😀" },
                      { type: "text", text: "aa" },
                    ],
                  },
                ],
              },
              {
                type: "run",
                content: [
                  { type: "text", text: "aaaa" },
                  { type: "text", text: "aa" },
                ],
              },
              tab(),
            ],
          },
        ],
      },
    },
  };
  const insertX =
    (offset: number) =>
    (index: number): DocumentOp => ({
      type: "insertText",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset },
      text: "x",
      runProps: {},
      newIds: NEW_IDS,
      revision: revision(index),
    });
  const link: ParagraphContent = {
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children: [
      {
        type: "inlineWrapper",
        kind: "bidi",
        control: "embedding",
        content: [{ type: "run", content: [{ type: "text", text: "aaa" }] }],
      },
      { type: "bookmarkStart", id: 1, name: "_Link" },
    ],
  };
  const tracked = opsOn(source, [
    insertX(0),
    (index) => ({
      type: "insertContent",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 5 },
      slice: { content: [link], openStart: 0, openEnd: 0 },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
    insertX(0),
    // Inside the inserted link's bidi wrapper: it cuts the insertion in two.
    insertX(7),
  ]);
  const perOp = dated(tracked, 4);
  expect(perOp[1]?.length).toBe(2);

  // The insertion the last operation cut is one record once that is rejected.
  const cutRejected = resolve(
    tracked,
    perOp[3] ?? [],
    REVISION_DECISIONS.REJECT,
  );
  expect(idsDated(cutRejected, date(1))).toHaveLength(1);

  const atOnce = resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT);
  const inTurn = rejectedInTurn(tracked, perOp);
  expect(body(inTurn)).toEqual(body(atOnce));
  // The bookmark start stays where the rejected link was: after the control.
  const paragraph = body(atOnce).at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("A paragraph.");
  expect(paragraph.content.map((node) => node.type)).toEqual([
    "inlineSdt",
    "bookmarkStart",
    "run",
    "run",
  ]);
});

test("a boundary kept beside a cut control goes where it was once the cut deletion is rejected too", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [] },
          {
            type: "blockSdt",
            properties: { sdtType: "group", id: 1 },
            content: [
              {
                type: "paragraph",
                paraId: "00000002",
                content: [
                  {
                    type: "inlineSdt",
                    properties: { sdtType: "richText", id: 2, tag: "clause" },
                    content: [
                      tab(),
                      {
                        type: "hyperlink",
                        rId: "rId9",
                        href: "https://example.org",
                        children: [
                          tab(),
                          { type: "bookmarkStart", id: 1, name: "_Link" },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  const tracked = opsOn(source, [
    (index) => ({
      type: "deleteRange",
      from: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 2 },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
    // Between the two tabs: outside the link, inside the deletion.
    (index) => ({
      type: "insertContent",
      at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 1 },
      slice: {
        content: [{ type: "bookmarkEnd", id: 2 }],
        openStart: 0,
        openEnd: 0,
      },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
    (index) => ({
      type: "insertText",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      text: "x",
      runProps: {},
      newIds: NEW_IDS,
      revision: revision(index),
    }),
  ]);
  const perOp = dated(tracked, 3);
  const atOnce = resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT);
  const inTurn = rejectedInTurn(tracked, perOp);
  expect(body(inTurn)).toEqual(body(atOnce));
  const control = body(atOnce).at(1);
  if (control?.type !== "blockSdt") throw new TypeError("A block control.");
  const paragraph = control.content.at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("A paragraph.");
  const inline = paragraph.content.at(0);
  if (inline?.type !== "inlineSdt") throw new TypeError("An inline control.");
  expect(inline.content).toEqual([
    tab(),
    { type: "bookmarkEnd", id: 2 },
    {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [tab(), { type: "bookmarkStart", id: 1, name: "_Link" }],
    },
  ]);
});

test("a boundary kept from content later edits wrapped goes where it was, in every rejection order", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "run",
                content: [
                  { type: "text", text: "aa" },
                  { type: "text", text: "😀😀a" },
                ],
              },
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 1, tag: "clause" },
                content: [
                  { type: "run", content: [{ type: "text", text: "aa" }] },
                  {
                    type: "hyperlink",
                    rId: "rId9",
                    href: "https://example.org",
                    children: [
                      {
                        type: "inlineWrapper",
                        kind: "bidi",
                        control: "embedding",
                        content: [
                          {
                            type: "run",
                            content: [
                              { type: "tab" },
                              { type: "text", text: "aa" },
                            ],
                          },
                          {
                            type: "run",
                            content: [
                              { type: "text", text: "aa" },
                              { type: "text", text: "aaaaa" },
                            ],
                          },
                        ],
                      },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  const tracked = opsOn(source, [
    // Between the two runs of the bidi wrapper: it cuts the wrapper, the link
    // and the control.
    (index) => ({
      type: "insertContent",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 12 },
      slice: {
        content: [
          { type: "moveToRangeStart", id: 1, name: "move1", author: "A" },
        ],
        openStart: 0,
        openEnd: 0,
      },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
    // Over the boundary: the piece before goes whole, the piece after only
    // at its start, so a deletion wraps the link inside that piece alone.
    (index) => ({
      type: "deleteRange",
      from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 13 },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
  ]);
  const perOp = dated(tracked, 2);
  const atOnce = resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT);
  for (const order of orders(perOp))
    expect(body(rejectedInOrder(tracked, order))).toEqual(body(atOnce));
  const paragraph = body(atOnce).at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("A paragraph.");
  const control = paragraph.content.at(1);
  if (control?.type !== "inlineSdt") throw new TypeError("An inline control.");
  const link = control.content.at(1);
  if (link?.type !== "hyperlink") throw new TypeError("A link.");
  expect(link.children).toHaveLength(1);
  const bidi = link.children.at(0);
  if (bidi?.type !== "inlineWrapper") throw new TypeError("A bidi wrapper.");
  expect(bidi.content.map((node) => node.type)).toEqual([
    "run",
    "moveToRangeStart",
    "run",
  ]);
});

test("a boundary kept beside a cut deletion goes where it was when the deletion is rejected first", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000002",
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 2, tag: "clause" },
                content: [
                  tab(),
                  {
                    type: "hyperlink",
                    rId: "rId9",
                    href: "https://example.org",
                    children: [
                      tab(),
                      { type: "bookmarkStart", id: 1, name: "_Link" },
                    ],
                  },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  const tracked = opsOn(source, [
    (index) => ({
      type: "deleteRange",
      from: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 0 },
      to: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 2 },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
    // Between the two tabs: it cuts the deletion and the control.
    (index) => ({
      type: "insertContent",
      at: { story: OP_STORIES.MAIN, blockId: "00000002", offset: 1 },
      slice: {
        content: [{ type: "bookmarkEnd", id: 2 }],
        openStart: 0,
        openEnd: 0,
      },
      newIds: NEW_IDS,
      revision: revision(index),
    }),
  ]);
  const perOp = dated(tracked, 2);
  const atOnce = resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT);
  // The deletion first: the cut it no longer holds is one level shallower.
  expect(body(rejectedInOrder(tracked, perOp))).toEqual(body(atOnce));
  expect(body(rejectedInTurn(tracked, perOp))).toEqual(body(atOnce));
  const paragraph = body(atOnce).at(0);
  if (paragraph?.type !== "paragraph") throw new TypeError("A paragraph.");
  const inline = paragraph.content.at(0);
  if (inline?.type !== "inlineSdt") throw new TypeError("An inline control.");
  expect(inline.content).toEqual([
    tab(),
    { type: "bookmarkEnd", id: 2 },
    {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [tab(), { type: "bookmarkStart", id: 1, name: "_Link" }],
    },
  ]);
});
