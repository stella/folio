/**
 * A cut belongs to the record it was made in, whatever later insertions stand
 * beside it: the other piece of a removed insertion beyond one, the pieces of
 * a pending insertion with only kept markers between them, and a marker kept
 * from outside the record a later insertion was made in. Resolving the
 * changes in either order leaves the same document.
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

const BOOKMARK_END: ParagraphContent = { type: "bookmarkEnd", id: 1 };
const BOOKMARK_START: ParagraphContent = {
  type: "bookmarkStart",
  id: 1,
  name: "_Ref1",
};

const insertion = (
  id: number,
  date: string,
  content: ParagraphContent[],
  joins: { before: number; after: number; remove: number },
): ParagraphContent => ({
  type: "insertion",
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

test("a cut goes with both pieces of the insertion it was made in, past a later insertion", () => {
  const document = documentOf({
    pPrMark: { kind: "ins", info: { id: 9, author: "A" } },
    content: [
      insertion(1, "1", [text("a")], { before: 0, after: 0, remove: 0 }),
      insertion(3, "2", [BOOKMARK_END], { before: 0, after: 0, remove: 3 }),
      insertion(4, "3", [text("x")], { before: 0, after: 3, remove: 0 }),
      insertion(2, "1", [text("a")], { before: 0, after: 0, remove: 0 }),
    ],
  });
  const atOnce = contentOf(rejected(document, [1, 2, 3, 4]));
  expect(atOnce).toEqual([
    { ...BOOKMARK_END, keptCut: { keptFrom: "2026-03-04T05:00:02Z" } },
  ]);
  expect(
    contentOf(rejected(rejected(rejected(document, [1, 2]), [4]), [3])),
  ).toEqual(atOnce);
  expect(
    contentOf(rejected(rejected(rejected(document, [4]), [1, 2]), [3])),
  ).toEqual(atOnce);
});

test("a marker kept between two pieces of a pending insertion joins them once accepted", () => {
  const document = documentOf({
    content: [
      text("a"),
      insertion(1, "1", [link([text("b")])], {
        before: 0,
        after: 0,
        remove: 2,
      }),
      insertion(3, "2", [BOOKMARK_END], { before: 0, after: 0, remove: 4 }),
      insertion(2, "1", [link([text("c")])], {
        before: 0,
        after: 0,
        remove: 2,
      }),
      text("a"),
    ],
  });
  const acceptedFirst = contentOf(rejected(accepted(document, [1, 2]), [3]));
  expect(acceptedFirst).toEqual([
    text("a"),
    link([text("b"), BOOKMARK_END, text("c")]),
    text("a"),
  ]);
  expect(contentOf(accepted(rejected(document, [3]), [1, 2]))).toEqual(
    acceptedFirst,
  );
  expect(contentOf(rejected(rejected(document, [3]), [1, 2]))).toEqual(
    contentOf(rejected(document, [1, 2, 3])),
  );
});

test("an insertion at the edge of a cut record's piece joins the pieces once accepted", () => {
  const document = documentOf({
    content: [
      link([text("a")]),
      insertion(1, "1", [BOOKMARK_END], { before: 0, after: 0, remove: 3 }),
      insertion(2, "2", [link([text("x")])], {
        before: 0,
        after: 3,
        remove: 0,
      }),
      link([text("b")]),
    ],
  });
  const acceptedFirst = contentOf(rejected(accepted(document, [2]), [1]));
  expect(acceptedFirst).toEqual([link([text("a"), BOOKMARK_END, text("xb")])]);
  expect(contentOf(accepted(rejected(document, [1]), [2]))).toEqual(
    acceptedFirst,
  );
  expect(contentOf(rejected(rejected(document, [1]), [2]))).toEqual(
    contentOf(rejected(document, [1, 2])),
  );
});

test("pieces of a rejected insertion meet where what was made inside them goes with them", () => {
  // Content made inside an inserted link splits it; text added at the link's
  // end was accepted, so the last piece is joined to it.
  const document = documentOf({
    content: [
      insertion(1, "1", [link([BOOKMARK_START])], {
        before: 0,
        after: 0,
        remove: 0,
      }),
      insertion(3, "3", [link([text("y")])], {
        before: 2,
        after: 4,
        remove: 2,
      }),
      insertion(2, "1", [link([text("a")])], {
        before: 0,
        after: 3,
        remove: 0,
      }),
      link([text("x")]),
    ],
  });
  const atOnce = contentOf(rejected(document, [1, 2, 3]));
  expect(atOnce).toEqual([link([BOOKMARK_START, text("x")])]);
  expect(contentOf(rejected(rejected(document, [3]), [1, 2]))).toEqual(atOnce);
});

test("a marker kept inside a rejected insertion's pieces goes with the insertion's own cut", () => {
  const document = documentOf({
    content: [
      link([text("a")]),
      insertion(1, "1", [link([text("b")])], {
        before: 0,
        after: 0,
        remove: 2,
      }),
      insertion(3, "2", [BOOKMARK_END], { before: 0, after: 0, remove: 2 }),
      insertion(2, "1", [link([text("c")])], {
        before: 0,
        after: 0,
        remove: 2,
      }),
      link([text("d")]),
      link([]),
    ],
  });
  const atOnce = contentOf(rejected(document, [1, 2, 3]));
  // The empty link beyond the cut is no piece of it.
  expect(atOnce).toEqual([
    link([text("a"), BOOKMARK_END, text("d")]),
    link([]),
  ]);
  expect(contentOf(rejected(rejected(document, [3]), [1, 2]))).toEqual(atOnce);
});

test("a marker kept inside a one-sided insertion's pieces goes into the piece it was joined to", () => {
  const document = documentOf({
    content: [
      link([text("a")]),
      insertion(1, "1", [link([text("x")])], {
        before: 3,
        after: 0,
        remove: 0,
      }),
      insertion(3, "2", [BOOKMARK_END], { before: 0, after: 0, remove: 4 }),
      insertion(2, "1", [link([text("x")])], {
        before: 3,
        after: 0,
        remove: 0,
      }),
    ],
  });
  const atOnce = contentOf(rejected(document, [1, 2, 3]));
  expect(atOnce).toEqual([link([text("a"), BOOKMARK_END])]);
  expect(contentOf(rejected(rejected(document, [3]), [1, 2]))).toEqual(atOnce);
  expect(contentOf(rejected(rejected(document, [1, 2]), [3]))).toEqual(atOnce);
});

test("an accepted deletion of one piece leaves the cut with the other piece", () => {
  const document = documentOf({
    content: [
      link([text("a")]),
      insertion(1, "1", [BOOKMARK_END], { before: 0, after: 0, remove: 3 }),
      {
        type: "deletion",
        info: { id: 2, author: "A", date: "2026-03-04T05:00:02Z" },
        content: [link([text("b")])],
        resolutionJoins: { before: 0, after: 0, remove: 0 },
      },
      text("c"),
    ],
  });
  const acceptedFirst = contentOf(rejected(accepted(document, [2]), [1]));
  expect(acceptedFirst).toEqual([link([text("a"), BOOKMARK_END]), text("c")]);
  expect(contentOf(accepted(rejected(document, [1]), [2]))).toEqual(
    acceptedFirst,
  );
});

test("a marker kept from outside the record a later insertion was made in stays beside it", () => {
  const document = documentOf({
    content: [
      text("a"),
      insertion(1, "1", [BOOKMARK_START, link([text("b")])], {
        before: 0,
        after: 0,
        remove: 1,
      }),
      insertion(2, "2", [link([text("x")])], {
        before: 4,
        after: 0,
        remove: 0,
      }),
      text("c"),
    ],
  });
  const acceptedFirst = contentOf(rejected(accepted(document, [2]), [1]));
  expect(acceptedFirst).toEqual([
    text("a"),
    BOOKMARK_START,
    link([text("x")]),
    text("c"),
  ]);
  expect(contentOf(accepted(rejected(document, [1]), [2]))).toEqual(
    acceptedFirst,
  );
});
