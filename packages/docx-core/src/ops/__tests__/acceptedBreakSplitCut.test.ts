import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const SPLIT = "2026-03-04T05:00:00Z";

const wrapped = (text: string): ParagraphContent => ({
  type: "inlineWrapper",
  kind: "bidi",
  control: "embedding",
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

/**
 * A wrapper a tracked split cut in two (insertion 1, cut two records deep),
 * the break's later deletion (2), and `before` at the end of the first piece.
 */
const splitDocument = (before: ParagraphContent[]): Document => {
  const first: Paragraph = {
    type: "paragraph",
    paraId: "10000000",
    content: [wrapped("a"), ...before],
    pPrMark: {
      kind: "ins",
      info: { id: 1, author: "A", date: SPLIT },
      resolutionJoin: 2,
      deletion: {
        info: { id: 2, author: "C", date: "2026-03-04T05:00:01Z" },
      },
    },
  };
  return {
    package: {
      document: {
        content: [
          first,
          { type: "paragraph", paraId: "00000001", content: [wrapped("a")] },
        ],
      },
    },
  };
};

const marker: ParagraphContent = {
  type: "moveToRangeStart",
  id: 1,
  name: "move1",
  author: "A",
};

/** A marker inserted at the first piece's end after the break was deleted. */
const insertedMarker: ParagraphContent = {
  type: "insertion",
  info: { id: 3, author: "C", date: "2026-03-04T05:00:02Z" },
  content: [marker],
  resolutionJoins: { before: 0, after: 0, remove: 0 },
};

const accepted = (document: Document, ids: number[]): Document => {
  const applied = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: ids,
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (applied.isErr()) throw applied.error;
  return applied.value.document;
};

const inTurn = (document: Document, groups: number[][]): Document => {
  let current = document;
  for (const ids of groups) current = accepted(current, ids);
  return current;
};

const body = (document: Document) => document.package.document.content;

describe("an accepted inserted break whose deletion is pending", () => {
  test("leaves the deletion holding the split's cut and date", () => {
    const [first] = body(accepted(splitDocument([]), [1]));
    expect(first?.type === "paragraph" ? first.pPrMark : undefined).toEqual({
      kind: "del",
      info: { id: 2, author: "C", date: "2026-03-04T05:00:01Z" },
      resolutionJoin: 2,
      splitDate: SPLIT,
    });
  });

  test("accepting the deletion later rejoins the pieces as accepting both at once", () => {
    const document = splitDocument([]);
    const rejoined = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [
          {
            type: "inlineWrapper",
            kind: "bidi",
            control: "embedding",
            content: [
              {
                type: "run",
                content: [
                  { type: "text", text: "a" },
                  { type: "text", text: "a" },
                ],
              },
            ],
          },
        ],
      },
    ];
    expect(body(accepted(document, [1, 2]))).toEqual(rejoined);
    expect(body(inTurn(document, [[1], [2]]))).toEqual(rejoined);
    expect(body(inTurn(document, [[2], [1]]))).toEqual(rejoined);
  });

  test("an accepted change at the cut's edge leaves the pieces apart in every order", () => {
    const document = splitDocument([insertedMarker]);
    const apart = [
      {
        type: "paragraph",
        paraId: "00000001",
        content: [wrapped("a"), marker, wrapped("a")],
      },
    ];
    expect(body(accepted(document, [1, 2, 3]))).toEqual(apart);
    expect(body(inTurn(document, [[1], [2], [3]]))).toEqual(apart);
    expect(body(inTurn(document, [[3], [1], [2]]))).toEqual(apart);
    expect(body(inTurn(document, [[1], [2, 3]]))).toEqual(apart);
  });
});
