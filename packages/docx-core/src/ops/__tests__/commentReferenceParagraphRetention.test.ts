import { expect, test } from "bun:test";

import type { Paragraph, ParagraphContent } from "../../model/document";
import { paragraphHoldsContent, paragraphLength } from "../offsets";
import { planParagraphBoundaryRemoval } from "../rangeBoundaries";

const commentReference = { type: "commentReference", id: 7 } as const;
const bookmarkStart = {
  type: "bookmarkStart",
  id: 9,
  name: "_Ref9",
} as const;

test("comment references keep their offset but do not retain a paragraph mark", () => {
  const referenceOnly: Paragraph = {
    type: "paragraph",
    content: [commentReference],
  };
  const textAndReference: Paragraph = {
    type: "paragraph",
    content: [
      { type: "run", content: [{ type: "text", text: "kept" }] },
      commentReference,
    ],
  };

  expect(paragraphLength(referenceOnly)).toBe(1);
  expect(paragraphHoldsContent(referenceOnly)).toBe(false);
  expect(paragraphLength(textAndReference)).toBe(5);
  expect(paragraphHoldsContent(textAndReference)).toBe(true);
});

test("removing a reference-only paragraph transfers its carrier to the previous paragraph", () => {
  const previous: Paragraph = {
    type: "paragraph",
    paraId: "00000001",
    content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
  };
  const removed: Paragraph = {
    type: "paragraph",
    paraId: "00000002",
    content: [commentReference],
  };
  const followingTable = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [{ type: "paragraph", paraId: "00000003", content: [] }],
          },
        ],
      },
    ],
  } as const;

  expect(
    planParagraphBoundaryRemoval({
      blocks: [previous, removed, followingTable],
      index: 1,
      place: "blocks",
    }),
  ).toEqual({
    type: "removed",
    blocks: [
      { ...previous, content: [...previous.content, commentReference] },
      followingTable,
    ],
    effects: [{ type: "retained-range-boundaries", receiver: previous }],
  });
});

test.each([
  {
    name: "bookmark before comment reference",
    carriers: [bookmarkStart, commentReference],
  },
  {
    name: "comment reference before bookmark",
    carriers: [commentReference, bookmarkStart],
  },
] satisfies readonly { name: string; carriers: ParagraphContent[] }[])(
  "removing a mixed-carrier paragraph preserves $name order in the previous paragraph",
  ({ carriers }) => {
    const previous: Paragraph = {
      type: "paragraph",
      paraId: "00000001",
      content: [{ type: "run", content: [{ type: "text", text: "kept" }] }],
    };
    const removed: Paragraph = {
      type: "paragraph",
      paraId: "00000002",
      content: carriers,
    };
    const followingTable = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              type: "tableCell",
              content: [{ type: "paragraph", paraId: "00000003", content: [] }],
            },
          ],
        },
      ],
    } as const;

    const plan = planParagraphBoundaryRemoval({
      blocks: [previous, removed, followingTable],
      index: 1,
      place: "blocks",
    });

    expect(plan).toEqual({
      type: "removed",
      blocks: [
        { ...previous, content: [...previous.content, ...carriers] },
        followingTable,
      ],
      effects: [{ type: "retained-range-boundaries", receiver: previous }],
    });
  },
);

test("an unhosted comment reference prevents removing its only paragraph", () => {
  const onlyParagraph: Paragraph = {
    type: "paragraph",
    paraId: "00000002",
    content: [commentReference],
  };
  const followingTable = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [{ type: "paragraph", paraId: "00000003", content: [] }],
          },
        ],
      },
    ],
  } as const;

  expect(
    planParagraphBoundaryRemoval({
      blocks: [onlyParagraph, followingTable],
      index: 0,
      place: "blocks",
    }),
  ).toEqual({ type: "blocked" });
});
