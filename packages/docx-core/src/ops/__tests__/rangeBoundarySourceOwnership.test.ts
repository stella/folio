import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { IDENTITY_SPACES, identityKeysIn } from "../ids";
import { childNodes, type InlineNode, rebuildNode } from "../leaves";
import { rangeBoundariesIn } from "../rangeBoundaries";
import { isTrackedWrapper } from "../review";
import type { DocumentOp } from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const FIRST_DATE = "2026-03-04T05:00:00Z";
const SECOND_DATE = "2026-03-04T05:00:01Z";
const THIRD_DATE = "2026-03-04T05:00:02Z";

type RangeBoundary = Extract<
  ParagraphContent,
  {
    type:
      | "bookmarkStart"
      | "bookmarkEnd"
      | "moveFromRangeStart"
      | "moveFromRangeEnd"
      | "moveToRangeStart"
      | "moveToRangeEnd";
  }
>;

type BoundaryPair = {
  insertedSide: "start" | "end";
  inserted: RangeBoundary;
  standing: RangeBoundary;
};

const boundaryPairs = [
  {
    insertedSide: "end",
    inserted: { type: "bookmarkEnd", id: 9 },
    standing: { type: "bookmarkStart", id: 9, name: "_Ref9" },
  },
  {
    insertedSide: "start",
    inserted: { type: "bookmarkStart", id: 9, name: "_Ref9" },
    standing: { type: "bookmarkEnd", id: 9 },
  },
  {
    insertedSide: "end",
    inserted: {
      type: "moveFromRangeEnd",
      id: 9,
    },
    standing: {
      type: "moveFromRangeStart",
      id: 9,
      name: "_MoveFrom9",
      author: "A",
    },
  },
  {
    insertedSide: "start",
    inserted: {
      type: "moveFromRangeStart",
      id: 9,
      name: "_MoveFrom9",
      author: "A",
    },
    standing: { type: "moveFromRangeEnd", id: 9 },
  },
  {
    insertedSide: "end",
    inserted: { type: "moveToRangeEnd", id: 9 },
    standing: {
      type: "moveToRangeStart",
      id: 9,
      name: "_MoveTo9",
      author: "A",
    },
  },
  {
    insertedSide: "start",
    inserted: {
      type: "moveToRangeStart",
      id: 9,
      name: "_MoveTo9",
      author: "A",
    },
    standing: { type: "moveToRangeEnd", id: 9 },
  },
] satisfies readonly BoundaryPair[];

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

const boundaryOffsets = (
  document: Document,
  target: RangeBoundary,
): number[] => {
  const offsets: number[] = [];
  let textOffset = 0;
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (node.type === "text") {
        textOffset += node.text.length;
        continue;
      }
      if (
        node.type === "bookmarkStart" ||
        node.type === "bookmarkEnd" ||
        node.type === "moveFromRangeStart" ||
        node.type === "moveFromRangeEnd" ||
        node.type === "moveToRangeStart" ||
        node.type === "moveToRangeEnd"
      ) {
        if (node.type === target.type && node.id === target.id)
          offsets.push(textOffset);
        continue;
      }
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraphOf(document).content);
  return offsets;
};

const paragraphText = (document: Document): string => {
  let text = "";
  const visit = (nodes: readonly InlineNode[]): void => {
    for (const node of nodes) {
      if (node.type === "text") {
        text += node.text;
        continue;
      }
      visit(childNodes(node) ?? []);
    }
  };
  visit(paragraphOf(document).content);
  return text;
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

test("directional source edges preserve separately floated bookmarks through interior rejections", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 1, max: 3 }),
      fc.constantFrom("a", "漢"),
      (interiorInsertions, authoredText) => {
        const firstLink = {
          type: "hyperlink",
          rId: "rId1",
          href: "https://example.org/first",
          children: [
            {
              type: "run",
              content: [
                { type: "tab" },
                { type: "tab" },
                { type: "text", text: authoredText },
              ],
            },
            { type: "bookmarkEnd", id: 8 },
          ],
        } satisfies ParagraphContent;
        const secondLink = {
          type: "hyperlink",
          rId: "rId2",
          href: "https://example.org/second",
          children: [{ type: "run", content: [{ type: "text", text: "z" }] }],
        } satisfies ParagraphContent;
        const source: Document = {
          package: {
            document: {
              content: [
                {
                  type: "paragraph",
                  paraId: "00000001",
                  content: [firstLink, secondLink],
                },
              ],
            },
          },
        };
        const operations = [
          {
            type: "insertText",
            at: { story: "main", blockId: "00000001", offset: 2 },
            text: "xxxx",
            runProps: {},
            revision: { id: 10, author: "C", date: FIRST_DATE },
            newIds: { revision: [100, 101, 102, 103] },
          },
          {
            type: "insertContent",
            at: { story: "main", blockId: "00000001", offset: 3 },
            slice: {
              content: [{ type: "bookmarkStart", id: 1, name: "_Ref1" }],
              openStart: 0,
              openEnd: 0,
            },
            revision: { id: 11, author: "C", date: SECOND_DATE },
            newIds: { revision: [200, 201, 202, 203, 204, 205] },
          },
          {
            type: "insertContent",
            at: { story: "main", blockId: "00000001", offset: 4 },
            slice: {
              content: [{ type: "bookmarkEnd", id: 1 }],
              openStart: 0,
              openEnd: 0,
            },
            revision: { id: 12, author: "C", date: THIRD_DATE },
            newIds: { revision: [250, 251, 252, 253, 254, 255] },
          },
          ...Array.from({ length: interiorInsertions }, (_, index) => ({
            type: "insertText" as const,
            at: { story: "main" as const, blockId: "00000001", offset: 4 },
            text: "x",
            runProps: {},
            revision: {
              id: index + 13,
              author: "C",
              date: `2026-03-04T05:00:0${index + 3}Z`,
            },
            newIds: { revision: [300 + index * 10, 301 + index * 10] },
          })),
        ] satisfies DocumentOp[];
        const applied = applyDocumentOps(source, operations);
        if (applied.isErr()) throw applied.error;
        const inverse: typeof applied.value.inverse = JSON.parse(
          JSON.stringify(applied.value.inverse),
        );
        const restored = applyDocumentOps(applied.value.document, inverse);
        if (restored.isErr()) throw restored.error;
        expect(restored.value.document).toStrictEqual(source);

        const groups = [FIRST_DATE, SECOND_DATE, THIRD_DATE].map((date) =>
          idsForDate(applied.value.document, date),
        );
        for (let index = 0; index < interiorInsertions; index++)
          groups.push(
            idsForDate(
              applied.value.document,
              `2026-03-04T05:00:0${index + 3}Z`,
            ),
          );
        for (const group of groups) expect(group.length).toBeGreaterThan(0);
        const batch = reject(applied.value.document, groups.flat());
        let reverse = applied.value.document;
        for (const group of groups.toReversed())
          reverse = reject(reverse, group);
        const partial = reject(
          applied.value.document,
          groups.at(1) ?? panic("The bookmark-start group exists."),
        );
        const partialThenRest = reject(
          partial,
          groups.filter((_, index) => index !== 1).flat(),
        );
        const resolved = applyDocumentOp(applied.value.document, {
          type: "resolveRevision",
          story: "main",
          revisionIds: groups.flat(),
          decision: "reject",
        });
        if (resolved.isErr()) throw resolved.error;
        const resolutionInverse: typeof resolved.value.inverse = JSON.parse(
          JSON.stringify(resolved.value.inverse),
        );
        const resolutionRestored = applyDocumentOps(
          resolved.value.document,
          resolutionInverse,
        );
        if (resolutionRestored.isErr()) throw resolutionRestored.error;
        expect(resolutionRestored.value.document).toStrictEqual(
          applied.value.document,
        );
        expect(projectReview({ document: batch, projection: "π′" })).toEqual(
          projectReview({ document: reverse, projection: "π′" }),
        );
        expect(projectReview({ document: batch, projection: "π′" })).toEqual(
          projectReview({ document: partialThenRest, projection: "π′" }),
        );
        const endFirstThenRest = reject(
          reject(
            applied.value.document,
            groups.at(2) ?? panic("The bookmark-end group exists."),
          ),
          groups.filter((_, index) => index !== 2).flat(),
        );
        expect(projectReview({ document: batch, projection: "π′" })).toEqual(
          projectReview({ document: endFirstThenRest, projection: "π′" }),
        );
        const content = paragraphOf(batch).content;
        const first = content.at(0);
        const second = content.at(-1);
        expect(first?.type).toBe("hyperlink");
        expect(second).toStrictEqual(secondLink);
        if (first?.type !== "hyperlink") return;
        expect(first.children).toContainEqual({
          type: "bookmarkStart",
          id: 1,
          name: "_Ref1",
        });
        expect(first.children).toContainEqual({
          type: "bookmarkEnd",
          id: 1,
        });
        expect(first.children).toContainEqual({
          type: "bookmarkEnd",
          id: 8,
        });
      },
    ),
    {
      numRuns: 12,
      id: "directional source edges preserve separately floated bookmarks through interior rejections",
    },
  );
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

type AdjacentRejectionOptions = {
  pair: BoundaryPair;
  text: string;
  offset: number;
  nested: boolean;
};

const assertAdjacentRejectionPreservesBoundary = ({
  pair,
  text,
  offset,
  nested,
}: AdjacentRejectionOptions): void => {
  const run = {
    type: "run",
    content: [{ type: "text", text }],
  } satisfies ParagraphContent;
  const sourceContent: ParagraphContent[] =
    pair.insertedSide === "end" ? [pair.standing, run] : [run, pair.standing];
  const sourceContentNodes = nested
    ? [
        {
          type: "inlineWrapper",
          kind: "bidi",
          control: "embedding",
          content: sourceContent,
        } satisfies ParagraphContent,
      ]
    : sourceContent;
  const source: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: sourceContentNodes,
          },
        ],
      },
    },
  };
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset },
    slice: { content: [pair.inserted], openStart: 0, openEnd: 0 },
    revision: { id: 2, author: "C", date: SECOND_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  });
  const deleted = apply(inserted, {
    type: "deleteRange",
    from: { story: "main", blockId: "00000001", offset },
    to: { story: "main", blockId: "00000001", offset: offset + 1 },
    revision: { id: 3, author: "C", date: THIRD_DATE },
    newIds: { revision: [200, 201, 202, 203] },
  });
  const addedIds = idsForDate(deleted, SECOND_DATE);
  const deletedIds = idsForDate(deleted, THIRD_DATE);
  expect(addedIds.length).toBeGreaterThan(0);
  expect(deletedIds.length).toBeGreaterThan(0);
  const batch = reject(deleted, [...addedIds, ...deletedIds]);
  const afterAdded = reject(deleted, addedIds);
  const afterDeleted = reject(deleted, deletedIds);
  const rejectAddedFirst = reject(reject(deleted, addedIds), deletedIds);
  const rejectDeletedFirst = reject(reject(deleted, deletedIds), addedIds);
  for (const result of [
    batch,
    afterAdded,
    afterDeleted,
    rejectAddedFirst,
    rejectDeletedFirst,
  ]) {
    expect(paragraphText(result)).toBe(text);
    expect(boundaryOffsets(result, pair.standing)).toEqual([
      pair.insertedSide === "end" ? 0 : text.length,
    ]);
    expect(boundaryOffsets(result, pair.inserted)).toEqual([offset]);
  }
  expect(projectReview({ document: batch, projection: "π′" })).toEqual(
    projectReview({ document: rejectAddedFirst, projection: "π′" }),
  );
  expect(projectReview({ document: batch, projection: "π′" })).toEqual(
    projectReview({ document: rejectDeletedFirst, projection: "π′" }),
  );
};

test("adjacent tracked insertions and deletions preserve range boundaries", () => {
  assertProperty(
    fc.property(
      fc.constantFrom(...boundaryPairs),
      fc.constantFrom("é漢abc", "Ж字def"),
      fc.integer({ min: 1, max: 4 }),
      fc.boolean(),
      (pair, text, offset, nested) => {
        assertAdjacentRejectionPreservesBoundary({
          pair,
          text,
          offset,
          nested,
        });
      },
    ),
    {
      numRuns: 60,
      id: "adjacent tracked insertions and deletions preserve range boundaries",
    },
  );
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
