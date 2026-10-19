import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphMarkCarrier, paragraphMarkRemovesParagraph } from "../blockDeletion";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const resolveAccepted = (document: Document, revisionIds = [7]): Document => {
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision: REVISION_DECISIONS.ACCEPT,
  });
  if (result.isErr()) throw result.error;
  return result.value.document;
};

const table = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [
        { type: "tableCell", content: [{ type: "paragraph", paraId: "00000002", content: [] }] },
      ],
    },
  ],
} as const;
const finalParagraph = { type: "paragraph", paraId: "00000003", content: [] } as const;

test("accepting an emptied paragraph-mark deletion removes its paragraph before a table", () => {
  const info = { id: 7, author: "Reviewer", date: "2026-10-19" };
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      {
        type: "deletion",
        info: { ...info, id: 8 },
        content: [{ type: "run", content: [{ type: "deletedText", text: "removed" }] }],
      },
    ],
    pPrMark: { kind: "del", info },
  };
  const document: Document = {
    package: { document: { content: [paragraph, table, finalParagraph] } },
  };

  expect(resolveAccepted(document, [7, 8]).package.document.content).toEqual([
    table,
    finalParagraph,
  ]);
});

test("accepting a non-empty paragraph-mark deletion runs it into the first cell paragraph", () => {
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
    pPrMark: { kind: "del", info: { id: 7, author: "Reviewer", date: "2026-10-19" } },
  };
  const document: Document = {
    package: { document: { content: [paragraph, table, finalParagraph] } },
  };

  expect(resolveAccepted(document).package.document.content).toEqual([
    {
      ...table,
      rows: [
        {
          ...table.rows[0],
          cells: [
            {
              ...table.rows[0].cells[0],
              content: [
                {
                  type: "paragraph",
                  paraId: "00000002",
                  content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
                },
              ],
            },
          ],
        },
      ],
    },
    finalParagraph,
  ]);
});

describe("paragraphMarkRemovesParagraph", () => {
  test.each([
    { followingKind: "table", holdsNoContent: true, markWasAdded: false, expected: true },
    { followingKind: "table", holdsNoContent: false, markWasAdded: false, expected: true },
    { followingKind: "table", holdsNoContent: false, markWasAdded: true, expected: true },
    { followingKind: "paragraph", holdsNoContent: false, markWasAdded: true, expected: true },
    { followingKind: undefined, holdsNoContent: false, markWasAdded: true, expected: true },
  ] as const)(
    "returns $expected before $followingKind when empty=$holdsNoContent and added=$markWasAdded",
    ({ expected, ...options }) => {
      expect(paragraphMarkRemovesParagraph(options)).toBe(expected);
    },
  );
});

describe("paragraphMarkCarrier", () => {
  test("the following paragraph owns a deleted break", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "paragraph",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("following");
  });

  test("a container-final paragraph keeps its required terminator", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: undefined,
        endsContainer: true,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a cleared split paragraph's retained mark stays before a table", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: false,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a paragraph mark remains the carrier before a table while its review resolves", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a resolved inserted paragraph stays as the table's preceding carrier", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: "table",
        endsContainer: false,
        removesCurrentMark: true,
        removesParagraph: false,
      }),
    ).toBe("current");
  });

  test("a removed inserted paragraph needs no container-final carrier", () => {
    expect(
      paragraphMarkCarrier({
        followingKind: undefined,
        endsContainer: true,
        removesCurrentMark: true,
        removesParagraph: true,
      }),
    ).toBe("none");
  });
});
