/**
 * Content inserted inside a pending insertion splits it into two pieces. A
 * rejection removing that content while another pending insertion stands
 * between the pieces defers the seam on it, at the depth the content cut:
 * rejecting the insertion standing there later joins the pieces, or keeps
 * its markers in that cut, as rejecting both at once does.
 */
import { expect, test } from "bun:test";

import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const piece = (id: number, value: string): ParagraphContent => ({
  type: "insertion",
  info: { id, author: "A" },
  content: [text(value)],
});

/** New content made inside the pending insertion: removing it rejoins it. */
const cutting: ParagraphContent = {
  type: "insertion",
  info: { id: 2, author: "C", date: "2026-03-04T05:00:00Z" },
  content: [text("b")],
  resolutionJoins: { before: 3, after: 3, remove: 3 },
};

/** Later content at the cut's first edge, joined to nothing there. */
const standing = (content: ParagraphContent): ParagraphContent => ({
  type: "insertion",
  info: { id: 3, author: "C", date: "2026-03-04T05:00:01Z" },
  content: [content],
  resolutionJoins: { before: 0, after: 0, remove: 0 },
});

const documentOf = (between: ParagraphContent): Document => ({
  package: {
    document: {
      content: [
        {
          type: "paragraph",
          paraId: "00000001",
          content: [piece(1, "a"), standing(between), cutting, piece(4, "aa")],
        },
      ],
    },
  },
});

const rejected = (document: Document, ids: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.REJECT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const contentOf = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : undefined;
};

test("a kept marker between pending pieces records their cut in either order", () => {
  const document = documentOf({ type: "moveToRangeEnd", id: 1 });
  const atOnce = contentOf(rejected(document, [2, 3]));
  const inTurn = contentOf(rejected(rejected(document, [2]), [3]));
  expect(inTurn).toEqual(atOnce);
  expect(atOnce?.[1]).toMatchObject({
    type: "moveToRangeEnd",
    keptCut: { depth: 3 },
  });
});

test("rejecting content between pending pieces joins them in either order", () => {
  const document = documentOf(text("x"));
  const atOnce = contentOf(rejected(document, [2, 3]));
  const inTurn = contentOf(rejected(rejected(document, [2]), [3]));
  expect(atOnce).toEqual([piece(1, "aaa")]);
  expect(inTurn).toEqual(atOnce);
});
