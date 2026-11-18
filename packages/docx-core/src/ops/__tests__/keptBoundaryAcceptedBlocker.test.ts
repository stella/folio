/**
 * A boundary a rejection keeps beside a later pending insertion stands at
 * the seam while the insertion is pending. Accepting the insertion then
 * places it as rejecting after accepting does: into the piece of the cut
 * record it touches.
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
  id: number,
  decision: RevisionDecision,
): Document =>
  apply(document, {
    type: "resolveRevision",
    story: OP_STORIES.MAIN,
    revisionIds: [id],
    decision,
  });

const tab: ParagraphContent = { type: "run", content: [{ type: "tab" }] };

const insertAt = (
  document: Document,
  content: ParagraphContent[],
  id: number,
  second: number,
): Document =>
  apply(document, {
    type: "insertContent",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    slice: { content, openStart: 0, openEnd: 0 },
    newIds: NEW_IDS,
    revision: { id, author: "C", date: `2026-03-04T05:00:0${second}Z` },
  });

test("a kept boundary beside an accepted later insertion lands in the piece it touches, in either order", () => {
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
                  { type: "run", content: [{ type: "text", text: "aaa" }] },
                ],
              },
            ],
          },
        ],
      },
    },
  };
  // A move range start and a tab cut the insertion; a later tab goes in
  // front of them, inside the cut.
  const first = insertAt(
    source,
    [{ type: "moveToRangeStart", id: 1, name: "move1", author: "A" }, tab],
    5001,
    0,
  );
  const tracked = insertAt(first, [tab], 5002, 1);
  const rejectedFirst = resolve(
    resolve(tracked, 5001, REVISION_DECISIONS.REJECT),
    5002,
    REVISION_DECISIONS.ACCEPT,
  );
  const acceptedFirst = resolve(
    resolve(tracked, 5002, REVISION_DECISIONS.ACCEPT),
    5001,
    REVISION_DECISIONS.REJECT,
  );
  expect(rejectedFirst.package.document.content).toStrictEqual(
    acceptedFirst.package.document.content,
  );
  expect(acceptedFirst.package.document.content).toMatchObject([
    {
      type: "paragraph",
      content: [
        { type: "insertion", content: [{ type: "run" }] },
        tab,
        {
          type: "insertion",
          content: [{ type: "moveToRangeStart", id: 1 }, { type: "run" }],
        },
      ],
    },
  ]);
});
