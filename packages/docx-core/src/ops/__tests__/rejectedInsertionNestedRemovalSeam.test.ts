/**
 * An insertion that cut a link whose pieces now sit in pending deletions,
 * with a bookmark it kept between them, defers the link's seam on those
 * deletions when it is rejected: rejecting them later rejoins the link
 * around the bookmark, as rejecting everything at once does.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { type DocumentOp, OP_STORIES, REVISION_DECISIONS } from "../types";

const NEW_IDS = {
  revision: [700, 701, 702, 703, 704, 705, 706, 707],
  control: [800, 801, 802],
};

const apply = (document: Document, op: DocumentOp): Document => {
  const applied = applyDocumentOp(document, op);
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const reject = (document: Document, ids: readonly number[]): Document =>
  apply(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.REJECT,
  });

const link = (children: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children,
});

const tracked = (): Document => {
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
                info: { id: 1, author: "A" },
                content: [
                  link([
                    { type: "run", content: [{ type: "text", text: "aa" }] },
                  ]),
                ],
              },
            ],
          },
        ],
      },
    },
  };
  // A link holding only a bookmark start goes inside the inserted link,
  // cutting it; then the text either side is deleted.
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    slice: {
      content: [link([{ type: "bookmarkStart", id: 1, name: "_Link" }])],
      openStart: 0,
      openEnd: 0,
    },
    newIds: NEW_IDS,
    revision: { id: 2, author: "C", date: "2026-03-04T05:00:00Z" },
  });
  return apply(inserted, {
    type: "deleteRange",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
    newIds: NEW_IDS,
    revision: { id: 3, author: "C", date: "2026-03-04T05:00:01Z" },
  });
};

test("rejecting the insertion that cut a link, then the deletions holding its pieces, rejoins the link around the kept bookmark", () => {
  const document = tracked();
  // The deletion's pieces: one in each piece of the link, one holding the
  // bookmark.
  const deletions = [3, 701, 702];
  const results = [
    reject(reject(document, [2]), deletions),
    reject(reject(document, deletions), [2]),
    reject(document, [2, ...deletions]),
  ];
  for (const rejected of results)
    expect(rejected.package.document.content).toMatchObject([
      {
        type: "paragraph",
        content: [
          {
            type: "insertion",
            info: { id: 1 },
            content: [
              {
                type: "hyperlink",
                children: [
                  { type: "run", content: [{ type: "text", text: "a" }] },
                  { type: "bookmarkStart", id: 1 },
                  { type: "run", content: [{ type: "text", text: "a" }] },
                ],
              },
            ],
          },
        ],
      },
    ]);
  const [inTurn, deletionsFirst, atOnce] = results.map(
    (rejected) => rejected.package.document.content,
  );
  expect(inTurn).toStrictEqual(atOnce);
  expect(deletionsFirst).toStrictEqual(atOnce);
});
