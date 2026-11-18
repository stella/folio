/**
 * A range boundary the document holds stays where it is when a resolution
 * joins a seam beside it: only boundaries kept from removed content float
 * into the pieces a seam joins.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
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

const revision = (id: number, second: number) => ({
  id,
  author: "C",
  date: `2026-03-04T05:00:0${second}Z`,
});

/** The ids of the tracked records dated `second` (any, when omitted), pieces included. */
const idsDated = (document: Document, second?: number): number[] => {
  const when = second === undefined ? undefined : revision(0, second).date;
  const ids = new Set<number>();
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (
        isTrackedWrapper(node) &&
        (when === undefined || node.info.date === when)
      )
        ids.add(node.info.id);
      visit(childNodes(node) ?? []);
    }
  };
  for (const block of document.package.document.content)
    if (block.type === "paragraph") visit(block.content);
  return [...ids];
};

const paragraph = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000001", content }],
    },
  },
});

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

test("a rejected deletion's own boundary stays beside the record before it", () => {
  // Deleting across a bookmark, then splitting the paragraph just after it,
  // leaves the bookmark alone in one piece of the deletion, which rejecting
  // the split does not merge with the other (they hold different identity
  // handoffs). Rejecting the deletion then restores the paragraph: the
  // bookmark is the deletion's content, never a boundary kept at the seam
  // its join recorded.
  const formatted: ParagraphContent = {
    type: "run",
    content: [{ type: "text", text: "aa" }],
    propertyChanges: [
      {
        type: "runPropertyChange",
        info: { id: 5, author: "A" },
        previousFormatting: { italic: true },
      },
    ],
  };
  const source = paragraph([
    { type: "insertion", info: { id: 1, author: "A" }, content: [text("b")] },
    { type: "bookmarkStart", id: 1, name: "_Ref1" },
    formatted,
  ]);
  const deleted = apply(source, {
    type: "deleteRange",
    from: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 0 },
    to: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 2 },
    newIds: NEW_IDS,
    revision: revision(2, 0),
  });
  const split = apply(deleted, {
    type: "splitBlock",
    at: {
      story: OP_STORIES.MAIN,
      blockId: "00000001",
      offset: 1,
      zeroWidthBefore: 1,
    },
    newBlockId: "10000000",
    newIds: NEW_IDS,
    revision: revision(3, 1),
  });
  const deletion = idsDated(split, 0);
  // The split's paragraph mark carries its revision.
  const splitIds = [3];
  const inTurn = resolve(
    resolve(split, splitIds, REVISION_DECISIONS.REJECT),
    deletion,
    REVISION_DECISIONS.REJECT,
  );
  const atOnce = resolve(
    split,
    [...deletion, ...splitIds],
    REVISION_DECISIONS.REJECT,
  );
  for (const rejected of [inTurn, atOnce])
    expect(rejected.package.document.content).toMatchObject([
      {
        type: "paragraph",
        content: [
          { type: "insertion", content: [text("b")] },
          { type: "bookmarkStart", id: 1 },
          { type: "run", content: [{ type: "text", text: "aa" }] },
        ],
      },
    ]);
});

test("a removed cut whose pieces all go floats no boundary beyond them", () => {
  // Content inserted inside a pending insertion cuts it in two. Rejecting
  // both leaves nothing of the record the cut split, so the move range end
  // after it stays outside the link that follows.
  const source = paragraph([
    {
      type: "insertion",
      info: { id: 1, author: "A" },
      content: [{ type: "bookmarkEnd", id: 3 }, text("aa")],
    },
    { type: "moveToRangeEnd", id: 1 },
    {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: [text("b")],
    },
  ]);
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: OP_STORIES.MAIN, blockId: "00000001", offset: 1 },
    slice: {
      content: [
        { type: "mathEquation", display: "inline", ommlXml: "<m:oMath/>" },
      ],
      openStart: 0,
      openEnd: 0,
    },
    newIds: NEW_IDS,
    revision: revision(2, 0),
  });
  const rejected = resolve(
    inserted,
    idsDated(inserted),
    REVISION_DECISIONS.REJECT,
  );
  expect(rejected.package.document.content).toMatchObject([
    {
      type: "paragraph",
      content: [
        { type: "bookmarkEnd", id: 3 },
        { type: "moveToRangeEnd", id: 1 },
        { type: "hyperlink", children: [text("b")] },
      ],
    },
  ]);
});
