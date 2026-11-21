/**
 * A boundary a rejection keeps goes as deep as the cut it stands at joins,
 * whether the changes around it are rejected at once or one after another.
 *
 * - A boundary inserted inside a pending insertion cuts it in two. Rejected,
 *   it never goes inside that pending insertion: it stands between its
 *   pieces, whether the new text around it goes first or with it.
 * - A boundary kept beside a removal's seam, with a container beyond it
 *   whose inner cut continues the seam, goes into the record that inner cut
 *   rejoins.
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

/** The ids of the tracked records dated `second`, pieces included. */
const idsDated = (document: Document, second: number): number[] => {
  const ids: number[] = [];
  const visit = (value: unknown): void => {
    if (Array.isArray(value)) {
      for (const item of value) visit(item);
      return;
    }
    if (typeof value !== "object" || value === null) return;
    const node = value as {
      type?: string;
      info?: { id: number; date?: string };
    };
    if (
      node.type === "insertion" &&
      node.info?.date === revision(0, second).date
    )
      ids.push(node.info.id);
    for (const child of Object.values(value)) visit(child);
  };
  visit(document.package.document.content);
  return ids;
};

const insertAt = (
  document: Document,
  offset: number,
  content: ParagraphContent[],
  id: number,
  second: number,
): Document =>
  apply(document, {
    type: "insertContent",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset },
    slice: { content, openStart: 0, openEnd: 0 },
    newIds: NEW_IDS,
    revision: revision(id, second),
  });

test("a boundary kept inside a rejected insertion stands between the pieces of the pending insertion its cut split, at once or in turn", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              {
                type: "insertion",
                info: { id: 2, author: "A" },
                content: [link([text("ad")])],
              },
            ],
          },
        ],
      },
    },
  };
  // New text inside the link's insertion, then a bookmark end inside that.
  const inserted = insertAt(
    source,
    1,
    [text("b"), { type: "run", content: [{ type: "tab" }] }],
    5001,
    0,
  );
  const tracked = insertAt(
    inserted,
    2,
    [{ type: "bookmarkEnd", id: 1 }],
    5002,
    1,
  );
  const insertedText = idsDated(tracked, 0);
  const boundary = idsDated(tracked, 1);
  expect({ text: insertedText.length, boundary: boundary.length }).toEqual({
    text: 2,
    boundary: 1,
  });
  const inTurn = resolve(
    resolve(tracked, boundary, REVISION_DECISIONS.REJECT),
    insertedText,
    REVISION_DECISIONS.REJECT,
  );
  const atOnce = resolve(
    tracked,
    [...boundary, ...insertedText],
    REVISION_DECISIONS.REJECT,
  );
  // It never goes inside the pending insertion: it stands between its
  // pieces, recording the cut it stands in.
  for (const rejected of [inTurn, atOnce])
    expect(rejected.package.document.content).toEqual([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "insertion",
            info: { id: 2, author: "A" },
            content: [link([text("a")])],
          },
          {
            type: "bookmarkEnd",
            id: 1,
            keptCut: { depth: 4, keptFrom: "2026-03-04T05:00:01Z" },
          },
          {
            type: "insertion",
            info: { id: 700, author: "A" },
            content: [link([text("d")])],
          },
        ],
      },
    ]);
});

test("a boundary kept beside a seam goes into the link a container's inner cut rejoins, at once or in turn", () => {
  const control = (content: ParagraphContent[]): ParagraphContent => ({
    type: "inlineSdt",
    properties: { sdtType: "richText", id: 1, tag: "clause" },
    content,
  });
  const emptyBidi: ParagraphContent = {
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [],
  };
  const tab: ParagraphContent = { type: "run", content: [{ type: "tab" }] };
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [control([link([emptyBidi, tab]), link([])])],
          },
        ],
      },
    },
  };
  // Text inside the link cuts it; a tab before that cuts the control; a
  // bookmark end after the tab stands between the control's pieces.
  const typed = apply(source, {
    type: "insertText",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    text: "x",
    newIds: NEW_IDS,
    revision: revision(5001, 0),
  });
  const tabbed = insertAt(typed, 0, [tab], 5002, 1);
  const tracked = insertAt(
    tabbed,
    1,
    [{ type: "bookmarkEnd", id: 1 }],
    5003,
    2,
  );
  const perOp = [0, 1, 2].map((second) => idsDated(tracked, second));
  expect(perOp.map((ids) => ids.length)).toEqual([1, 1, 1]);
  let inTurn = tracked;
  for (const ids of perOp.toReversed())
    inTurn = resolve(inTurn, ids, REVISION_DECISIONS.REJECT);
  const atOnce = resolve(tracked, perOp.flat(), REVISION_DECISIONS.REJECT);
  for (const rejected of [inTurn, atOnce])
    expect(rejected.package.document.content).toEqual([
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          control([
            link([emptyBidi, { type: "bookmarkEnd", id: 1 }, tab]),
            link([]),
          ]),
        ],
      },
    ]);
});
