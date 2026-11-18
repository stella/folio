/**
 * Accepting an insertion made inside another hands its join to the older
 * insertion's piece it meets, which is a piece of the record it cut, not a
 * change made inside it: rejecting that piece later removes it and joins
 * nothing beyond it. A bookmark the rejection keeps stays where it was, out
 * of the link the accepted text went into, in any order.
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

const link = (children: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});

const insertX = (document: Document, id: number, second: number): Document =>
  apply(document, {
    type: "insertText",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    text: "x",
    newIds: NEW_IDS,
    revision: { id, author: "C", date: `2026-03-04T05:00:0${second}Z` },
  });

test("a bookmark kept by rejecting an insertion stays out of a link accepted text went into, in any order", () => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            // A content control holding nothing to type in stands first.
            content: [
              {
                type: "inlineSdt",
                properties: { sdtType: "richText", id: 1, tag: "clause" },
                content: [
                  link([
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      content: [],
                    },
                  ]),
                ],
              },
            ],
          },
        ],
      },
    },
  };
  // A bookmark and a link go in; then text goes in at the link's start,
  // twice.
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    slice: {
      content: [
        { type: "bookmarkStart", id: 1, name: "_Ref1" },
        link([{ type: "run", content: [{ type: "text", text: "a" }] }]),
      ],
      openStart: 0,
      openEnd: 0,
    },
    newIds: NEW_IDS,
    revision: { id: 1, author: "C", date: "2026-03-04T05:00:00Z" },
  });
  const tracked = insertX(insertX(inserted, 2, 1), 3, 2);
  // The first insertion's pieces: the bookmark's and the link's.
  const first = [1, 700];
  const oneByOne = resolve(
    resolve(
      resolve(tracked, [3], REVISION_DECISIONS.REJECT),
      [2],
      REVISION_DECISIONS.ACCEPT,
    ),
    first,
    REVISION_DECISIONS.REJECT,
  );
  const acceptedFirst = resolve(
    resolve(tracked, [2], REVISION_DECISIONS.ACCEPT),
    [...first, 3],
    REVISION_DECISIONS.REJECT,
  );
  for (const resolved of [oneByOne, acceptedFirst])
    expect(resolved.package.document.content).toMatchObject([
      {
        type: "paragraph",
        content: [
          { type: "inlineSdt" },
          { type: "bookmarkStart", id: 1 },
          {
            type: "hyperlink",
            children: [{ type: "run", content: [{ type: "text", text: "x" }] }],
          },
        ],
      },
    ]);
});
