/**
 * A tracked deletion of a range takes in the records it covers whole with the
 * comment anchors they hold, and resolves as the direct deletion does.
 *
 * Over paragraphs of runs and hyperlinks, some inside content controls, with
 * comment anchors inside the hyperlinks (both ends of a comment, or one end
 * whose other end lies outside the paragraph), every range whose direct
 * deletion applies:
 *
 * - is tracked, never refused;
 * - accepts to the direct deletion's document;
 * - rejects to the document it started from.
 *
 * Documents are compared under the review projection the tracked-operation
 * laws use.
 */
import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type {
  CommentRangeEnd,
  CommentRangeStart,
  Document,
  Hyperlink,
  ParagraphContent,
  Run,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphLength } from "../offsets";
import {
  type DeleteRangeOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(60_000));

type Anchor = CommentRangeStart | CommentRangeEnd;
type Part = "text" | "anchor";
type Piece =
  | { kind: "run" }
  | { kind: "link"; inControl: boolean; parts: readonly Part[] };

const start = (id: number): Anchor => ({ type: "commentRangeStart", id });
const end = (id: number): Anchor => ({ type: "commentRangeEnd", id });

/**
 * The comment anchors a paragraph holds, in document order: each comment's
 * start before its end, or one end whose other end lies outside it.
 */
const ANCHOR_ORDERS: readonly (readonly Anchor[])[] = [
  [start(1), end(1)],
  [start(1), start(2), end(1), end(2)],
  [start(1), start(2), end(2), end(1)],
  [end(1), start(2), end(2)],
  [start(1), end(1), start(2)],
  [start(1)],
  [end(1)],
];

const pieceArbitrary: fc.Arbitrary<Piece> = fc.oneof(
  fc.constant({ kind: "run" } as const),
  fc.record({
    kind: fc.constant("link" as const),
    inControl: fc.boolean(),
    parts: fc.array(fc.constantFrom<Part>("text", "anchor"), {
      minLength: 1,
      maxLength: 4,
    }),
  }),
);

const run = (): Run => ({
  type: "run",
  content: [{ type: "text", text: "ab" }],
});

const documentOf = (
  pieces: readonly Piece[],
  order: readonly Anchor[],
): Document => {
  const anchors = [...order];
  let control = 0;
  const content = pieces.map((piece): ParagraphContent => {
    if (piece.kind === "run") return run();
    const children: Hyperlink["children"] = piece.parts.flatMap(
      (part): Hyperlink["children"] => {
        if (part === "text") return [run()];
        const anchor = anchors.shift();
        return anchor === undefined ? [] : [anchor];
      },
    );
    const link: Hyperlink = {
      type: "hyperlink",
      rId: "rId9",
      href: "https://example.org",
      children: children.length === 0 ? [run()] : children,
    };
    if (!piece.inControl) return link;
    control += 1;
    return {
      type: "inlineSdt",
      properties: { sdtType: "richText", id: control, tag: "clause" },
      content: [link],
    };
  });
  return {
    package: {
      document: {
        content: [{ type: "paragraph", paraId: "00000001", content }],
      },
    },
  };
};

const revision = { id: 900, author: "C", date: "2026-03-04T05:00:00Z" };
const newIds = {
  revision: Array.from({ length: 32 }, (_, index) => 303_648 + index),
  control: Array.from({ length: 16 }, (_, index) => 293_648 + index),
};

type Range = {
  from: number;
  to: number;
  fromZeroWidth: number | undefined;
  toZeroWidth: number | undefined;
};

const at = (offset: number, zeroWidthBefore: number | undefined) => ({
  story: OP_STORIES.MAIN,
  blockId: "00000001",
  offset,
  ...(zeroWidthBefore === undefined ? {} : { zeroWidthBefore }),
});

const directDeletion = (range: Range): DeleteRangeOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_RANGE,
  from: at(range.from, range.fromZeroWidth),
  to: at(range.to, range.toZeroWidth),
});

const resolved = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
): Document => {
  if (revisionIds.length === 0) return document;
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const caseArbitrary = fc
  .record({
    pieces: fc.array(pieceArbitrary, { minLength: 1, maxLength: 6 }),
    order: fc.constantFrom(...ANCHOR_ORDERS),
    ends: fc.tuple(fc.nat(), fc.nat()),
    fromZeroWidth: fc.option(fc.nat({ max: 2 }), { nil: undefined }),
    toZeroWidth: fc.option(fc.nat({ max: 2 }), { nil: undefined }),
  })
  .map(({ pieces, order, ends, fromZeroWidth, toZeroWidth }) => {
    const document = documentOf(pieces, order);
    const paragraph = document.package.document.content[0];
    const length =
      paragraph?.type === "paragraph" ? paragraphLength(paragraph) : 0;
    const [first, second] = ends.map((value) => value % (length + 1));
    const range: Range = {
      from: Math.min(first ?? 0, second ?? 0),
      to: Math.max(first ?? 0, second ?? 0),
      fromZeroWidth,
      toZeroWidth,
    };
    return { document, range };
  });

describe("a tracked deletion over records holding comment anchors", () => {
  test("is tracked, accepts to the direct deletion and rejects to the document", () => {
    assertProperty(
      fc.property(caseArbitrary, ({ document, range }) => {
        const direct = applyDocumentOp(document, directDeletion(range));
        if (direct.isErr()) return;
        const tracked = applyDocumentOp(document, {
          ...directDeletion(range),
          revision,
          newIds,
        });
        if (tracked.isErr()) throw tracked.error;
        const accepted = resolved(
          tracked.value.document,
          tracked.value.revisions,
          REVISION_DECISIONS.ACCEPT,
        );
        expect(projectReview({ document: accepted })).toStrictEqual(
          projectReview({ document: direct.value.document }),
        );
        const rejected = resolved(
          tracked.value.document,
          tracked.value.revisions,
          REVISION_DECISIONS.REJECT,
        );
        expect(projectReview({ document: rejected })).toStrictEqual(
          projectReview({ document }),
        );
      }),
      {
        numRuns: 400,
        id: "a tracked deletion over records holding comment anchors resolves as the direct deletion",
      },
    );
  });
});
