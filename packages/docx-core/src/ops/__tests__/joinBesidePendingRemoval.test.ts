/**
 * A boundary kept from rejected content at a split's seam stands between
 * the pieces of a pending deletion or move source there, whether the split
 * is rejected first, last or with the content: those pieces stand only once
 * that change is rejected.
 */
import { expect, test } from "bun:test";

import type { Document } from "../../model/document";
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

for (const removal of ["deletion", "moveFrom"] as const)
  test(`a kept move range end stands between the pieces of a pending ${removal} in any order`, () => {
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "paragraph",
              paraId: "00000001",
              content: [
                {
                  type: removal,
                  info: { id: 1, author: "A" },
                  content: [
                    { type: "run", content: [{ type: "text", text: "aa" }] },
                  ],
                },
              ],
            },
          ],
        },
      },
    };
    const inserted = apply(source, {
      type: "insertContent",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      slice: {
        content: [{ type: "moveToRangeEnd", id: 1 }],
        openStart: 0,
        openEnd: 0,
      },
      newIds: NEW_IDS,
      revision: { id: 2, author: "C", date: "2026-03-04T05:00:00Z" },
    });
    const split = apply(inserted, {
      type: "splitBlock",
      at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
      newBlockId: "10000000",
      newIds: NEW_IDS,
      revision: { id: 3, author: "C", date: "2026-03-04T05:00:02Z" },
    });
    const results = [
      reject(reject(split, [2]), [3]),
      reject(reject(split, [3]), [2]),
      reject(split, [2, 3]),
    ];
    for (const rejected of results)
      expect(rejected.package.document.content).toMatchObject([
        {
          type: "paragraph",
          content: [
            { type: removal, info: { id: 1 } },
            { type: "moveToRangeEnd", id: 1 },
            { type: removal },
          ],
        },
      ]);
  });
