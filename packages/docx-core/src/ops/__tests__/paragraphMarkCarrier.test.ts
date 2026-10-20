import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphMarkOutcome } from "../blockDeletion";
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

test("refuses one owned reason when an empty mark precedes a section-ending table cell", () => {
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [],
    pPrMark: { kind: "ins", info: { id: 7, author: "Reviewer" } },
  };
  const sectionTable = structuredClone(table);
  const target = sectionTable.rows[0]?.cells[0]?.content[0];
  if (target?.type !== "paragraph") throw new Error("The fixture has a first-cell paragraph.");
  target.sectionProperties = { pageWidth: 11_906 };
  const document: Document = {
    package: { document: { content: [paragraph, sectionTable, finalParagraph] } },
  };

  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [7],
    decision: REVISION_DECISIONS.REJECT,
  });

  expect(result.isErr() && result.error.reason).toBe("untrackable");
});

describe("paragraphMarkOutcome", () => {
  test.each([
    {
      follows: "paragraph",
      holdsContent: false,
      markWasAdded: false,
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn: false,
      expected: "join",
    },
    {
      follows: "paragraph",
      holdsContent: true,
      markWasAdded: true,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: false,
      expected: "join",
    },
    {
      follows: "table",
      holdsContent: false,
      markWasAdded: false,
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn: false,
      tableRemovalIsTrackable: false,
      expected: "refuse",
    },
    {
      follows: "table",
      holdsContent: false,
      markWasAdded: true,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: true,
      expected: "remove",
    },
    {
      follows: "table",
      holdsContent: true,
      markWasAdded: false,
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn: true,
      expected: "run-in",
    },
    {
      follows: "table",
      holdsContent: true,
      markWasAdded: true,
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn: false,
      expected: "keep",
    },
    {
      follows: "nothing",
      holdsContent: false,
      markWasAdded: true,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: false,
      expected: "remove",
    },
    {
      follows: "nothing",
      holdsContent: false,
      markWasAdded: false,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: false,
      expected: "keep",
    },
    {
      follows: "nothing",
      holdsContent: false,
      markWasAdded: true,
      containerStillEndsWithParagraph: false,
      followingTableCanReceiveRunIn: false,
      expected: "keep",
    },
    {
      follows: "nothing",
      holdsContent: true,
      markWasAdded: true,
      containerStillEndsWithParagraph: true,
      followingTableCanReceiveRunIn: false,
      expected: "keep",
    },
  ] as const)("returns $expected for $follows", ({ expected, ...input }) => {
    expect(paragraphMarkOutcome(input).kind).toBe(expected);
  });
});
