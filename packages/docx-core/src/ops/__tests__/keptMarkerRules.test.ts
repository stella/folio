import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { childNodes, type InlineNode } from "../leaves";
import { isTrackedWrapper } from "../review";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const { ACCEPT, REJECT } = REVISION_DECISIONS;

const documentOf = (content: BlockContent[]): Document => ({
  package: { document: { content } },
});

const resolve = (
  document: Document,
  ids: number[],
  decision: RevisionDecision,
): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

type Step = { ids: number[]; decision: RevisionDecision };

const resolvedInTurn = (document: Document, steps: readonly Step[]) => {
  let current = document;
  for (const { ids, decision } of steps)
    current = resolve(current, ids, decision);
  return current;
};

const link = (children: ParagraphContent[]) =>
  ({
    type: "hyperlink",
    rId: "rId9",
    href: "https://example.org",
    children,
  }) as const;

const text = (value: string) =>
  ({ type: "run", content: [{ type: "text", text: value }] }) as const;

const stamp = (id: number, second: number) => ({
  id,
  author: "C",
  date: `2026-03-04T05:00:0${String(second)}Z`,
});

/** Every range marker in `nodes` held inside a pending tracked change. */
const markersInsidePending = (
  nodes: readonly InlineNode[],
  inside = false,
): InlineNode[] =>
  nodes.flatMap((node) => {
    if (
      inside &&
      (node.type === "bookmarkStart" || node.type === "bookmarkEnd")
    )
      return [node];
    return markersInsidePending(
      childNodes(node) ?? [],
      inside || isTrackedWrapper(node),
    );
  });

describe("rejecting the insertion a pending insertion's join faces cancels that join", () => {
  // Insertion 2's link is joined after (3 levels) to the piece insertion 3
  // was made inside; insertion 5 kept a bookmark between the controls.
  const tracked = (): Document =>
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", id: 2, tag: "clause" },
            content: [
              {
                type: "insertion",
                info: stamp(4, 2),
                content: [link([text("x")])],
                resolutionJoins: { before: 0, after: 4, remove: 0 },
              },
              {
                type: "insertion",
                info: stamp(2, 0),
                content: [link([text("x")])],
                resolutionJoins: { before: 0, after: 3, remove: 0 },
              },
              {
                type: "insertion",
                info: stamp(3, 1),
                content: [link([text("x")])],
                resolutionJoins: { before: 4, after: 0, remove: 0 },
              },
            ],
          },
          {
            type: "insertion",
            info: stamp(5, 3),
            content: [link([{ type: "bookmarkStart", id: 1, name: "_Link" }])],
            resolutionJoins: { before: 0, after: 0, remove: 1 },
          },
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", id: 293648, tag: "clause" },
            content: [link([text("a")])],
          },
        ],
      },
    ]);
  const paragraph = (document: Document) =>
    document.package.document.content.at(0);

  test("rejecting the faced insertion retires the facing join", () => {
    const rejected = paragraph(resolve(tracked(), [3], REJECT));
    if (rejected?.type !== "paragraph") throw new Error("expected a paragraph");
    const control = rejected.content.at(0);
    if (control?.type !== "inlineSdt") throw new Error("expected a control");
    const facing = control.content.at(-1);
    expect(facing?.type === "insertion" && facing.resolutionJoins).toEqual({
      before: 0,
      after: 0,
      remove: 0,
    });
  });

  test("the bookmark stands between the links whichever batch goes first", () => {
    const acceptedFirst = resolvedInTurn(tracked(), [
      { ids: [2, 4], decision: ACCEPT },
      { ids: [3, 5], decision: REJECT },
    ]);
    const rejectedFirst = resolvedInTurn(tracked(), [
      { ids: [3, 5], decision: REJECT },
      { ids: [2, 4], decision: ACCEPT },
    ]);
    expect(paragraph(rejectedFirst)).toEqual(paragraph(acceptedFirst));
    expect(paragraph(acceptedFirst)).toEqual({
      type: "paragraph",
      paraId: "00000001",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 2, tag: "clause" },
          content: [
            link([text("xx")]),
            { type: "bookmarkStart", id: 1, name: "_Link" },
            link([text("a")]),
          ],
        },
      ],
    });
  });
});

describe("a kept range marker stands beside a pending tracked change, never inside it", () => {
  // Another reviewer's deletions inside A's insertions, with C's bookmark
  // start and end inserted between them, and C's text in a paragraph before.
  const pendingPair = (id: number, deletion: number, value: string) =>
    ({
      type: "insertion",
      info: { id, author: "A" },
      content: [
        {
          type: "deletion",
          info: { id: deletion, author: "B", date: "2026-01-02T03:04:05Z" },
          content: [text(value)],
        },
      ],
    }) as const;
  const tracked = (): Document =>
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "insertion",
            info: stamp(303651, 2),
            content: [text("x")],
            resolutionJoins: { before: 0, after: 0, remove: 0 },
          },
        ],
      },
      {
        type: "paragraph",
        paraId: "00000002",
        content: [
          pendingPair(1, 5000, "a"),
          {
            type: "insertion",
            info: stamp(5001, 0),
            content: [{ type: "bookmarkStart", id: 1, name: "_Ref1" }],
            resolutionJoins: { before: 0, after: 0, remove: 4 },
          },
          {
            type: "insertion",
            info: stamp(303650, 1),
            content: [{ type: "bookmarkEnd", id: 1 }],
            resolutionJoins: { before: 0, after: 0, remove: 0 },
          },
          pendingPair(303648, 303649, "aa"),
        ],
      },
    ]);
  const second = (document: Document) => {
    const block = document.package.document.content.at(1);
    if (block?.type !== "paragraph") throw new Error("expected a paragraph");
    return block.content;
  };
  const ids = [5001, 303650, 303651];

  test("rejecting at once and in any order leaves the markers between the changes", () => {
    const atOnce = second(resolve(tracked(), ids, REJECT));
    expect(markersInsidePending(atOnce)).toEqual([]);
    expect(atOnce.map((node) => node.type)).toEqual([
      "insertion",
      "bookmarkStart",
      "bookmarkEnd",
      "insertion",
    ]);
    for (const order of [ids, ids.toReversed(), [303650, 5001, 303651]]) {
      const inTurn = second(
        resolvedInTurn(
          tracked(),
          order.map((id) => ({ ids: [id], decision: REJECT })),
        ),
      );
      expect(markersInsidePending(inTurn)).toEqual([]);
      expect(inTurn.map((node) => node.type)).toEqual(
        atOnce.map((node) => node.type),
      );
    }
  });
});

describe("pieces of one rejected insertion meet around the markers that waited beside them", () => {
  // Insertion 1's link, split into two pieces by insertion 2, continues the
  // link after it; insertion 303649 kept a bookmark between the controls.
  const tracked = (): Document =>
    documentOf([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", id: 1, tag: "clause" },
            content: [
              {
                type: "insertion",
                info: stamp(1, 0),
                content: [link([text("x")])],
                resolutionJoins: { before: 0, after: 3, remove: 0 },
              },
            ],
          },
          {
            type: "insertion",
            info: stamp(303649, 2),
            content: [link([{ type: "bookmarkStart", id: 1, name: "_Link" }])],
            resolutionJoins: { before: 0, after: 0, remove: 1 },
          },
          {
            type: "inlineSdt",
            properties: { sdtType: "richText", id: 293648, tag: "clause" },
            content: [
              {
                type: "insertion",
                info: stamp(2, 1),
                content: [link([text("xxx")])],
                resolutionJoins: { before: 4, after: 4, remove: 4 },
              },
              {
                type: "insertion",
                info: stamp(303648, 0),
                content: [link([text("x")])],
                resolutionJoins: { before: 0, after: 3, remove: 0 },
              },
              link([text("a")]),
            ],
          },
        ],
      },
    ]);
  const first = (document: Document) => document.package.document.content.at(0);

  test("rejecting the pieces last places the bookmark as rejecting at once does", () => {
    const atOnce = resolve(tracked(), [1, 2, 303648, 303649], REJECT);
    const inTurn = resolvedInTurn(
      tracked(),
      [[303649], [2], [1, 303648]].map((ids) => ({ ids, decision: REJECT })),
    );
    expect(first(atOnce)).toEqual({
      type: "paragraph",
      paraId: "00000001",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText", id: 1, tag: "clause" },
          content: [
            link([{ type: "bookmarkStart", id: 1, name: "_Link" }, text("a")]),
          ],
        },
      ],
    });
    expect(first(inTurn)).toEqual(first(atOnce));
  });
});
