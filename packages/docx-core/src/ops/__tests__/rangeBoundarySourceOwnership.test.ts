import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { IDENTITY_SPACES, identityKeysIn } from "../ids";
import { childNodes, type InlineNode, rebuildNode } from "../leaves";
import { rangeBoundariesIn } from "../rangeBoundaries";
import { isTrackedWrapper } from "../review";
import type { DocumentOp } from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const FIRST_DATE = "2026-03-04T05:00:00Z";
const SECOND_DATE = "2026-03-04T05:00:01Z";

const apply = (document: Document, operation: DocumentOp) => {
  const result = applyDocumentOp(document, operation);
  if (result.isErr()) throw result.error;
  return result.value.document;
};

const paragraphOf = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    throw new TypeError("Expected the fixture paragraph.");
  return paragraph;
};

const idsForDate = (document: Document, date: string) => {
  const ids: number[] = [];
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node) && node.info.date === date)
        ids.push(node.info.id);
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraphOf(document).content);
  return ids;
};

const reject = (document: Document, revisionIds: number[]) => {
  const present = new Set(identityKeysIn(document.package));
  const wanted = revisionIds.filter((id) =>
    present.has(`${IDENTITY_SPACES.REVISION}:${id}`),
  );
  if (wanted.length === 0) return document;
  return apply(document, {
    type: "resolveRevision",
    story: "main",
    revisionIds: wanted,
    decision: "reject",
  });
};

type FixtureOptions = {
  foreign: ParagraphContent;
  text: string;
  offset: number;
  marker: Extract<ParagraphContent, { type: "bookmarkEnd" | "moveToRangeEnd" }>;
};

const assertSourceOwnership = ({
  foreign,
  text,
  offset,
  marker,
}: FixtureOptions): void => {
  const source: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [foreign] },
        ],
      },
    },
  };
  const first = apply(source, {
    type: "insertText",
    at: { story: "main", blockId: "00000001", offset: 0 },
    text,
    runProps: {},
    revision: { id: 2, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  });
  const second = apply(first, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset },
    slice: { content: [marker], openStart: 0, openEnd: 0 },
    revision: { id: 3, author: "C", date: SECOND_DATE },
    newIds: { revision: [200, 201, 202, 203, 204, 205] },
  });
  const batch = reject(second, [
    ...idsForDate(second, FIRST_DATE),
    ...idsForDate(second, SECOND_DATE),
  ]);
  const partial = reject(second, idsForDate(second, SECOND_DATE));
  const sourceFragments: InlineNode[] = [];
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (isTrackedWrapper(node) && node.info.date === FIRST_DATE)
        sourceFragments.push(node);
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraphOf(partial).content);
  expect(
    sourceFragments.some((node) =>
      rangeBoundariesIn(childNodes(node) ?? []).some(
        (boundary) =>
          boundary.type === marker.type && boundary.id === marker.id,
      ),
    ),
  ).toBe(true);
  const reverse = reject(partial, idsForDate(second, FIRST_DATE));
  expect(projectReview({ document: batch, projection: "π′" })).toEqual(
    projectReview({ document: reverse, projection: "π′" }),
  );
  expect(paragraphOf(batch).content).toEqual([foreign, marker]);
};

const standing = { type: "bookmarkEnd", id: 3 } satisfies ParagraphContent;

const foreignShapes = [
  {
    type: "insertion",
    info: { id: 1, author: "A" },
    content: [standing],
  },
  { type: "hyperlink", anchor: "target", children: [standing] },
  {
    type: "inlineSdt",
    properties: { id: 10, sdtType: "richText" },
    content: [standing],
  },
  {
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [standing],
  },
] satisfies ParagraphContent[];

test("batch rejection leaves a marker outside a foreign insertion when both source-cut pieces go", () => {
  assertSourceOwnership({
    foreign: {
      type: "insertion",
      info: { id: 1, author: "A" },
      content: [standing],
    },
    text: "xxx",
    offset: 1,
    marker: { type: "moveToRangeEnd", id: 1 },
  });
});

test("batch rejection keeps an authored start before a foreign bidi comment carrier", () => {
  const foreign = {
    type: "inlineWrapper",
    kind: "bidi",
    control: "embedding",
    content: [{ type: "commentReference", id: 7 }],
  } satisfies ParagraphContent;
  const marker = {
    type: "bookmarkStart",
    id: 1,
    name: "_Ref1",
  } satisfies ParagraphContent;
  const source: Document = {
    package: {
      document: {
        content: [
          { type: "paragraph", paraId: "00000001", content: [foreign] },
        ],
        comments: [
          {
            id: 7,
            author: "A",
            content: [{ type: "paragraph", paraId: "7FFFFFF0", content: [] }],
          },
        ],
      },
    },
  };
  const operations = [
    {
      type: "insertContent",
      at: { story: "main", blockId: "00000001", offset: 0 },
      slice: {
        content: [{ type: "run", content: [{ type: "text", text: "aa" }] }],
        openStart: 0,
        openEnd: 0,
      },
      revision: { id: 2, author: "C", date: FIRST_DATE },
      newIds: { revision: [100, 101, 102, 103] },
    },
    {
      type: "insertContent",
      at: { story: "main", blockId: "00000001", offset: 1 },
      slice: { content: [marker], openStart: 0, openEnd: 0 },
      revision: { id: 3, author: "C", date: SECOND_DATE },
      newIds: { revision: [200, 201, 202, 203, 204, 205] },
    },
    ...[2, 3].map(
      (index) =>
        ({
          type: "insertText",
          at: { story: "main", blockId: "00000001", offset: 0 },
          text: "x",
          runProps: {},
          revision: {
            id: index + 2,
            author: "C",
            date: `2026-03-04T05:00:0${index}Z`,
          },
          newIds: { revision: [300 + index * 10, 301 + index * 10] },
        }) as const,
    ),
  ] satisfies DocumentOp[];
  let edited = source;
  for (const operation of operations) edited = apply(edited, operation);
  const perOperation = operations.map(({ revision }) =>
    idsForDate(edited, revision.date),
  );
  const batch = reject(edited, perOperation.flat());
  let reverse = edited;
  for (const ids of perOperation.toReversed()) reverse = reject(reverse, ids);
  expect(projectReview({ document: batch, projection: "π′" })).toEqual(
    projectReview({ document: reverse, projection: "π′" }),
  );
  expect(paragraphOf(batch).content).toEqual([marker, foreign]);
});

test("retained boundary attachment belongs to surviving source occurrences, never merely compatible neighbors", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(...foreignShapes),
      fc.integer({ min: 3, max: 20 }),
      fc.nat({ max: 19 }),
      fc.constantFrom("bookmarkEnd", "moveToRangeEnd"),
      (foreign, length, cut, type) => {
        assertSourceOwnership({
          foreign,
          text: "x".repeat(length),
          offset: 1 + (cut % (length - 1)),
          marker: { type, id: 9 },
        });
      },
    ),
    {
      numRuns: 80,
      id: "retained boundary attachment belongs to surviving source occurrences, never merely compatible neighbors",
    },
  );
});

test("source occurrence ownership stays distinct when authored containers reuse the same immutable object", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(
        ...foreignShapes.filter(
          (node) => node.type === "hyperlink" || node.type === "inlineWrapper",
        ),
      ),
      fc.integer({ min: 2, max: 5 }),
      (shared, count) => {
        const marker = {
          type: "bookmarkEnd",
          id: 9,
        } satisfies ParagraphContent;
        const source: Document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [
                    {
                      type: "insertion",
                      info: { id: 1, author: "C" },
                      resolutionJoins: { before: 0, after: 0, remove: 1 },
                      content: [marker],
                    },
                    ...Array.from({ length: count }, () => shared),
                  ],
                },
              ],
            },
          },
        };
        const content = paragraphOf(reject(source, [1])).content;
        expect(content.at(0)).toEqual(
          rebuildNode(shared, [marker, ...(childNodes(shared) ?? [])]),
        );
        expect(content.slice(1)).toHaveLength(count - 1);
        for (const untouched of content.slice(1))
          expect(untouched).toEqual(shared);
        for (const original of paragraphOf(source).content.slice(1))
          expect(original).toBe(shared);
        expect(childNodes(shared)).toEqual([standing]);
      },
    ),
    {
      numRuns: 30,
      id: "source occurrence ownership stays distinct when authored containers reuse the same immutable object",
    },
  );
});
