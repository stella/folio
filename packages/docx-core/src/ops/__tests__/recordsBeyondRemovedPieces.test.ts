/**
 * The records beyond a removal's neighbours are pieces of the record it cut
 * only where the removal was made inside them. A pending deletion's pieces
 * hold no content made inside them, and a pending insertion beyond pieces
 * that go with the removal is no piece of the record they were: resolving
 * the removal before or with its neighbours leaves the same document.
 */
import { expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const text = (value: string): ParagraphContent => ({
  type: "run",
  content: [{ type: "text", text: value }],
});

const link = (content: ParagraphContent[]): ParagraphContent => ({
  type: "hyperlink",
  rId: "rId9",
  href: "https://example.org",
  children: content,
});

const change = (
  type: "insertion" | "deletion",
  id: number,
  date: string,
  content: ParagraphContent[],
  joins: { before: number; after: number; remove: number },
): ParagraphContent => ({
  type,
  info: { id, author: "A", date: `2026-03-04T05:00:0${date}Z` },
  content,
  resolutionJoins: joins,
});

const documentOf = (paragraph: Omit<Paragraph, "type">): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", ...paragraph },
        { type: "paragraph", paraId: "00000002", content: [] },
      ],
    },
  },
});

const resolved = (
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
const accepted = (document: Document, ids: number[]) =>
  resolved(document, ids, REVISION_DECISIONS.ACCEPT);
const rejected = (document: Document, ids: number[]) =>
  resolved(document, ids, REVISION_DECISIONS.REJECT);

const contentOf = (document: Document) => {
  const [paragraph] = document.package.document.content;
  return paragraph?.type === "paragraph" ? paragraph.content : undefined;
};

test("a marker kept between a deletion's pieces stays beside a later insertion", () => {
  // Content made between the pieces of a pending deletion, unlike across it,
  // and a later insertion made at its start, standing between it and the
  // first piece.
  const document = documentOf({
    content: [
      change("deletion", 1, "1", [link([text("t")])], {
        before: 0,
        after: 0,
        remove: 0,
      }),
      change("insertion", 5, "5", [link([text("x")])], {
        before: 0,
        after: 4,
        remove: 0,
      }),
      change(
        "insertion",
        3,
        "3",
        [link([text("y")]), { type: "bookmarkEnd", id: 1 }],
        { before: 0, after: 0, remove: 1 },
      ),
      change("deletion", 4, "1", [text("a")], {
        before: 0,
        after: 0,
        remove: 0,
      }),
    ],
  });
  const inTurn = contentOf(
    accepted(rejected(accepted(document, [1, 4]), [3]), [5]),
  );
  const rejectedFirst = contentOf(accepted(rejected(document, [3]), [1, 4, 5]));
  expect(inTurn).toEqual([link([text("x")]), { type: "bookmarkEnd", id: 1 }]);
  expect(rejectedFirst).toEqual(inTurn);
});

test("a marker kept with the pieces it was made inside reads no cut beyond them", () => {
  // Content made inside a pending insertion splits it; a later insertion
  // made at the first piece's start stands beyond it.
  const document = documentOf({
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
    content: [
      change("insertion", 4, "2", [text("x")], {
        before: 0,
        after: 3,
        remove: 0,
      }),
      change("insertion", 1, "0", [text("a")], {
        before: 0,
        after: 0,
        remove: 0,
      }),
      change(
        "insertion",
        3,
        "1",
        [link([{ type: "bookmarkStart", id: 2, name: "_Link" }])],
        { before: 0, after: 0, remove: 2 },
      ),
      change("insertion", 2, "0", [text("a")], {
        before: 0,
        after: 0,
        remove: 0,
      }),
      { type: "bookmarkStart", id: 1, name: "_Ref1" },
    ],
  });
  const inTurn = contentOf(
    accepted(rejected(rejected(document, [1, 2]), [3]), [4]),
  );
  const atOnce = contentOf(accepted(rejected(document, [1, 2, 3]), [4]));
  expect(atOnce).toEqual(inTurn);
  expect(atOnce?.[1]).toMatchObject({
    type: "bookmarkStart",
    id: 2,
    keptCut: { depth: 1 },
  });
});
