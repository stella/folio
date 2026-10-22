import { describe, expect, test } from "bun:test";

import type { Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphMarkOutcome } from "../blockDeletion";
import { paragraphHoldsContent } from "../offsets";
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
        {
          type: "tableCell",
          content: [{ type: "paragraph", paraId: "00000002", content: [] }],
        },
      ],
    },
  ],
} as const;
const finalParagraph = {
  type: "paragraph",
  paraId: "00000003",
  content: [],
} as const;

test("accepting an emptied paragraph-mark deletion removes its paragraph before a table", () => {
  const info = { id: 7, author: "Reviewer", date: "2026-10-19" };
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      {
        type: "deletion",
        info: { ...info, id: 8 },
        content: [
          { type: "run", content: [{ type: "deletedText", text: "removed" }] },
        ],
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

test("rejecting an empty inserted final paragraph after a table keeps it", () => {
  const info = { id: 7, author: "Reviewer", date: "2026-10-19" };
  const first = { type: "paragraph", paraId: "00000001", content: [] } as const;
  const inserted: Paragraph = {
    type: "paragraph",
    paraId: "00000003",
    content: [],
    pPrMark: { kind: "ins", info },
  };
  const document: Document = {
    package: { document: { content: [first, table, inserted] } },
  };
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [7],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (result.isErr()) throw result.error;

  expect(result.value.document.package.document.content).toEqual([
    first,
    table,
    finalParagraph,
  ]);
});

test("rejecting an empty inserted final paragraph in a cell keeps the cell terminator", () => {
  const info = { id: 7, author: "Reviewer", date: "2026-10-19" };
  const kept = {
    type: "paragraph",
    paraId: "00000002",
    content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
  } as const;
  const cell = (content: readonly Paragraph[]) => ({
    type: "table",
    rows: [{ type: "tableRow", cells: [{ type: "tableCell", content }] }],
  });
  const document = {
    package: {
      document: {
        content: [
          cell([
            kept,
            {
              type: "paragraph",
              paraId: "00000004",
              content: [],
              pPrMark: { kind: "ins", info },
            },
          ]),
          finalParagraph,
        ],
      },
    },
  } as Document;
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [7],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (result.isErr()) throw result.error;

  expect(result.value.document.package.document.content).toEqual([
    cell([kept, { type: "paragraph", paraId: "00000004", content: [] }]),
    finalParagraph,
  ]);
});

test("accepting a mark deletion before a table keeps a bookmark the paragraph holds", () => {
  const info = { id: 7, author: "Reviewer", date: "2026-10-19" };
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      { type: "bookmarkStart", id: 1, name: "kept" },
      { type: "bookmarkEnd", id: 1 },
    ],
    pPrMark: { kind: "del", info },
  };
  const document: Document = {
    package: { document: { content: [paragraph, table, finalParagraph] } },
  };

  const content = JSON.stringify(
    resolveAccepted(document).package.document.content,
  );
  expect(content).toContain('"bookmarkStart"');
  expect(content).toContain('"bookmarkEnd"');
});

test("accepting a non-empty paragraph-mark deletion runs it into the first cell paragraph", () => {
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
    pPrMark: {
      kind: "del",
      info: { id: 7, author: "Reviewer", date: "2026-10-19" },
    },
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
                  content: [
                    { type: "run", content: [{ type: "text", text: "kept" }] },
                  ],
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

test("rejecting a marker-only inline content control does not touch the following table", () => {
  const paragraph: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [
      {
        type: "inlineSdt",
        properties: { sdtType: "richText", id: 1, tag: "clause" },
        content: [
          {
            type: "hyperlink",
            rId: "rId9",
            href: "https://example.org",
            children: [{ type: "bookmarkStart", id: 1, name: "_Link" }],
          },
        ],
      },
    ],
    propertyChanges: [
      {
        type: "paragraphPropertyChange",
        info: { id: 1, author: "A" },
        previousFormatting: { alignment: "end" },
      },
    ],
    pPrMark: { kind: "ins", info: { id: 2, author: "A" } },
  };
  const followingTable = {
    ...table,
    columnWidths: [2400, 2400],
    rows: [
      {
        ...table.rows[0],
        preservedAttributes: [{ name: "rsidTr", value: "00ABCDEF" }],
      },
    ],
  } as const;
  const document: Document = {
    package: {
      document: { content: [paragraph, followingTable, finalParagraph] },
    },
  };

  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [2],
    decision: REVISION_DECISIONS.REJECT,
  });
  if (result.isErr()) throw result.error;

  expect(result.value.document.package.document.content).toEqual([
    followingTable,
    finalParagraph,
  ]);
  expect(result.value.touched.modified).not.toContain("00000002");
});

test("only atomic fields make marker-only inline containers paragraph content", () => {
  const marker = { type: "bookmarkStart", id: 1, name: "_Link" } as const;
  const hyperlink = {
    type: "hyperlink",
    rId: "rId9",
    children: [marker],
  } as const;

  expect(
    paragraphHoldsContent({ type: "paragraph", content: [hyperlink] }),
  ).toBe(false);
  expect(
    paragraphHoldsContent({
      type: "paragraph",
      content: [
        {
          type: "inlineSdt",
          properties: { sdtType: "richText" },
          content: [hyperlink],
        },
      ],
    }),
  ).toBe(false);
  expect(
    paragraphHoldsContent({
      type: "paragraph",
      content: [
        {
          type: "simpleField",
          instruction: "PAGE",
          fieldType: "PAGE",
          content: [hyperlink],
        },
      ],
    }),
  ).toBe(true);
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
  if (target?.type !== "paragraph")
    throw new Error("The fixture has a first-cell paragraph.");
  target.sectionProperties = { pageWidth: 11_906 };
  const document: Document = {
    package: {
      document: { content: [paragraph, sectionTable, finalParagraph] },
    },
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
