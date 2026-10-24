import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
  Table,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { paragraphHoldsContent } from "../offsets";
import { planParagraphBoundaryRemoval } from "../rangeBoundaries";

setDefaultTimeout(propertyTestTimeout(30_000));

const host = {
  type: "paragraph",
  paraId: "00000001",
  content: [
    { type: "commentRangeStart", id: 7 },
    { type: "run", content: [{ type: "text", text: "surviving text" }] },
  ],
} satisfies Paragraph;

const cellParagraph = {
  type: "paragraph",
  paraId: "00000004",
  content: [],
} satisfies Paragraph;
const table = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [cellParagraph] }],
    },
  ],
} satisfies Table;
const finalParagraph = {
  type: "paragraph",
  paraId: "00000005",
  content: [],
} satisfies Paragraph;
const end = { type: "commentRangeEnd", id: 7 } satisfies ParagraphContent;

const REMOVAL_CASES = {
  "before-table": {
    prefix: "host",
    context: "before-table",
    payload: "boundaries",
    expected: "receiver",
  },
  terminal: {
    prefix: "host",
    context: "terminal",
    payload: "boundaries",
    expected: "receiver",
  },
  "before-table-no-boundaries": {
    prefix: "none",
    context: "before-table",
    payload: "empty",
    expected: "none",
  },
  "terminal-no-boundaries": {
    prefix: "host",
    context: "terminal",
    payload: "empty",
    expected: "none",
  },
  "blocked-no-receiver": {
    prefix: "none",
    context: "before-table",
    payload: "boundaries",
    expected: "blocked",
  },
  "blocked-table-predecessor": {
    prefix: "table",
    context: "terminal",
    payload: "boundaries",
    expected: "blocked",
  },
} as const;

const BOUNDARIES = [
  end,
  { type: "bookmarkEnd", id: 9 },
  { type: "moveToRangeEnd", id: 11 },
] satisfies ParagraphContent[];

const retainedCarriersIn = (
  value: unknown,
  retained: ReadonlySet<object>,
): object[] => {
  if (value === null || typeof value !== "object") return [];
  if (retained.has(value)) return [value];
  const children = Array.isArray(value) ? value : Object.values(value);
  return children.flatMap((child) => retainedCarriersIn(child, retained));
};

const PREFIXES = {
  host: [host],
  table: [table],
  none: [],
} satisfies Record<
  (typeof REMOVAL_CASES)[keyof typeof REMOVAL_CASES]["prefix"],
  readonly BlockContent[]
>;

test("paragraph removal plans declare exactly the retained-boundary receiver across structural cases", () => {
  const exercised = new Set<string>();
  assertProperty(
    fc.property(
      fc.shuffledSubarray(BOUNDARIES, {
        minLength: BOUNDARIES.length,
        maxLength: BOUNDARIES.length,
      }),
      fc.boolean(),
      (boundaries, nested) => {
        for (const [name, fixture] of Object.entries(REMOVAL_CASES)) {
          exercised.add(name);
          const payload = fixture.payload === "empty" ? [] : boundaries;
          const retired = {
            type: "paragraph",
            paraId: "00000002",
            content:
              nested && payload.length > 0
                ? [
                    {
                      type: "inlineWrapper",
                      kind: "bidi",
                      control: "embedding",
                      content: payload,
                    },
                  ]
                : payload,
          } satisfies Paragraph;
          const prefix = PREFIXES[fixture.prefix];
          const hasCommentCarrier = payload.some(
            (node) => node.type === "commentRangeEnd",
          );
          const placedInPrevious =
            payload.length > 0 &&
            (fixture.context === "terminal" || hasCommentCarrier);
          const blocks: BlockContent[] = [
            ...prefix,
            retired,
            ...(fixture.context === "before-table"
              ? [table, finalParagraph]
              : []),
          ];
          const plan = planParagraphBoundaryRemoval({
            blocks,
            index: prefix.length,
            place: fixture.context === "before-table" ? "blocks" : "previous",
          });
          expect(blocks.at(prefix.length)).toBe(retired);
          if (plan.type === "blocked") {
            expect(fixture.expected).toBe("blocked");
            continue;
          }
          expect(fixture.expected).not.toBe("blocked");
          expect(plan.blocks).not.toContain(retired);
          const between =
            fixture.context === "before-table" && !hasCommentCarrier
              ? payload
              : [];
          expect(
            retainedCarriersIn(plan.blocks, new Set<object>(payload)),
          ).toEqual(payload);
          expect(
            plan.blocks.filter(
              (block) =>
                block.type === "bookmarkEnd" || block.type === "moveToRangeEnd",
            ),
          ).toEqual(between);
          if (!placedInPrevious) {
            expect(fixture.expected).toBe("none");
            expect(plan.effects).toEqual([]);
            expect(plan.blocks.includes(host)).toBe(fixture.prefix === "host");
            for (const block of plan.blocks)
              expect(blocks.includes(block)).toBe(true);
            continue;
          }
          expect(fixture.expected).toBe("receiver");
          expect(plan.effects.at(0)?.receiver).toBe(host);
          expect(plan.effects).toHaveLength(1);
          expect(plan.blocks.at(0)).toEqual({
            ...host,
            content: [...host.content, ...payload],
          });
          for (const block of plan.blocks.slice(1))
            expect(
              blocks.includes(block) ||
                between.some((boundary) => boundary === block),
            ).toBe(true);
        }
      },
    ),
    {
      numRuns: 30,
      id: "paragraph removal plans declare exactly the retained-boundary receiver across structural cases",
    },
  );
  expect([...exercised].sort()).toEqual(Object.keys(REMOVAL_CASES).sort());
});

test("bookmark-only paragraph carriers remain between blocks", () => {
  const bookmarkStart = {
    type: "bookmarkStart",
    id: 15,
    name: "_Ref15",
  } as const;
  const bookmarkEnd = { type: "bookmarkEnd", id: 15 } as const;
  const retired = {
    type: "paragraph",
    paraId: "00000002",
    content: [bookmarkStart, bookmarkEnd],
  } satisfies Paragraph;
  const blocks: BlockContent[] = [host, retired, table, finalParagraph];

  const plan = planParagraphBoundaryRemoval({
    blocks,
    index: 1,
    place: "blocks",
  });

  expect(plan).toEqual({
    type: "removed",
    blocks: [host, bookmarkStart, bookmarkEnd, table, finalParagraph],
    effects: [],
  });
  expect(
    retainedCarriersIn(
      plan.type === "removed" ? plan.blocks : [],
      new Set<object>([bookmarkStart, bookmarkEnd]),
    ),
  ).toEqual([bookmarkStart, bookmarkEnd]);
});

test.each(["anchor-only", "tracked-material"] as const)(
  "accepting %s paragraph removal modifies only its actual comment-anchor receiver",
  (payload) => {
    const retired = {
      type: "paragraph",
      paraId: "00000002",
      content: [
        ...(payload === "tracked-material"
          ? [
              {
                type: "deletion",
                info: { id: 8, author: "Reviewer" },
                content: [
                  {
                    type: "run",
                    content: [
                      { type: "deletedText", text: "removed material" },
                    ],
                  },
                ],
              } satisfies ParagraphContent,
            ]
          : []),
        end,
      ],
      pPrMark: { kind: "del", info: { id: 4, author: "Reviewer" } },
    } satisfies Paragraph;
    const source: Document = {
      package: {
        document: {
          content: [host, retired, table, finalParagraph],
          comments: [
            {
              id: 7,
              author: "Reviewer",
              content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
            },
          ],
        },
      },
    };
    expect(paragraphHoldsContent(retired)).toBe(payload === "tracked-material");
    const plan = planParagraphBoundaryRemoval({
      blocks: source.package.document.content,
      index: 1,
      place: "blocks",
    });
    if (plan.type === "blocked")
      throw new TypeError("The fixture has its preceding anchor receiver.");
    const named = new Set([
      retired.paraId,
      ...plan.effects.map(({ receiver }) => receiver.paraId),
    ]);
    const result = applyDocumentOp(source, {
      type: "resolveRevision",
      story: "main",
      revisionIds: payload === "tracked-material" ? [4, 8] : [4],
      decision: "accept",
    });
    if (result.isErr()) throw result.error;
    expect(new Set(result.value.touched.modified)).toEqual(
      new Set([host.paraId]),
    );
    expect(result.value.touched.removed).toEqual([retired.paraId]);
    expect(result.value.touched.inserted).toEqual([]);
    for (const id of [
      ...result.value.touched.modified,
      ...result.value.touched.removed,
    ])
      expect(named.has(id)).toBe(true);
    const blocks = result.value.document.package.document.content;
    expect(blocks).toEqual([
      { ...host, content: [...host.content, end] },
      table,
      finalParagraph,
    ]);
    const followingTable = blocks.at(1);
    expect(followingTable).toBe(table);
    expect(blocks.at(2)).toBe(finalParagraph);
    if (followingTable?.type !== "table")
      throw new TypeError("The following table stays.");
    expect(followingTable.rows.at(0)?.cells.at(0)?.content.at(0)).toBe(
      cellParagraph,
    );
  },
);
