import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import { projectReview } from "../../../../../test/reviewProjection";
import type { Document, ParagraphContent } from "../../model/document";
import { applyDocumentOp, stampOf } from "../apply";
import { storyParagraphs } from "../blocks";
import { contractViolation } from "../contract";
import {
  acceptCopyOriginAncestor,
  acceptCopyOriginAncestors,
  addCopyOriginAncestor,
  copyOriginAncestorsOf,
  copyOriginBoundaryKeysInBlocks,
  hasValidCopyOriginAncestors,
  isCopyOriginAncestors,
  rejectCopyOriginAncestors,
  rejectCopyOriginAncestorsInNodes,
  shiftCopyOriginAncestors,
} from "../copyOriginAncestors";
import { IDENTITY_SPACES, identityKeysIn } from "../ids";
import { deleteBetween, insertSliceAt } from "../inline";
import { childNodes, type InlineNode, leafSpans } from "../leaves";
import {
  isTrackedWrapper,
  stampedRevisionIds,
  WRAP_KINDS,
  wrapTracked,
} from "../review";
import type { DocumentOp, RevisionStamp } from "../types";
import { type OpSeed, trackedOpFor } from "./documentArbitraries";

const FIRST_DATE = "2026-03-04T05:00:00Z";
const SECOND_DATE = "2026-03-04T05:00:01Z";

const documentOf = (content: ParagraphContent[]): Document => ({
  package: {
    document: {
      content: [{ type: "paragraph", paraId: "00000001", content }],
    },
  },
});

const apply = (document: Document, operation: DocumentOp) => {
  const result = applyDocumentOp(document, operation);
  if (result.isErr()) throw result.error;
  return result.value;
};

const paragraphOf = (document: Document) => {
  const paragraph = document.package.document.content.at(0);
  if (paragraph?.type !== "paragraph")
    throw new TypeError("Expected a paragraph.");
  return paragraph;
};

const allIds = (document: Document) => {
  const prefix = `${IDENTITY_SPACES.REVISION}:`;
  return identityKeysIn(document.package).flatMap((key) =>
    key.startsWith(prefix) ? [Number(key.slice(prefix.length))] : [],
  );
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
  const present = new Set(allIds(document));
  const wanted = revisionIds.filter((revisionId) => present.has(revisionId));
  if (wanted.length === 0) return document;
  return apply(document, {
    type: "resolveRevision",
    story: "main",
    revisionIds: wanted,
    decision: "reject",
  }).document;
};

const accept = (document: Document, revisionIds: number[]) => {
  const present = new Set(allIds(document));
  const wanted = revisionIds.filter((revisionId) => present.has(revisionId));
  if (wanted.length === 0) return document;
  return apply(document, {
    type: "resolveRevision",
    story: "main",
    revisionIds: wanted,
    decision: "accept",
  }).document;
};

const bookmarkEnds = (
  nodes: readonly InlineNode[],
): Extract<InlineNode, { type: "bookmarkEnd" }>[] =>
  nodes.flatMap((node) =>
    node.type === "bookmarkEnd" ? [node] : bookmarkEnds(childNodes(node) ?? []),
  );

const project = (document: Document) =>
  projectReview({ document, projection: "π′" });

const sourceInsertion = {
  type: "insertion",
  info: { id: 1, author: "A" },
  content: [{ type: "bookmarkEnd", id: 3 }],
} satisfies ParagraphContent;

const copiedDocument = () => {
  const source = documentOf([
    { type: "run", content: [{ type: "text", text: "😀" }] },
    sourceInsertion,
    { type: "run", content: [{ type: "tab" }] },
  ]);
  const first = apply(source, {
    type: "insertText",
    at: { story: "main", blockId: "00000001", offset: 2 },
    text: "x",
    runProps: {},
    revision: { id: 2, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  });
  const second = apply(first.document, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 0 },
    slice: {
      content: paragraphOf(first.document).content,
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 3, author: "C", date: SECOND_DATE },
    newIds: { revision: [200, 201, 202, 203, 204, 205] },
  });
  return { source, document: second.document };
};

test("rejecting a copied closed slice preserves the original marker position", () => {
  const { source, document } = copiedDocument();
  const paragraph = paragraphOf(document);
  const copied = paragraph.content.flatMap((node) => {
    const visit = (current: InlineNode): InlineNode[] => {
      if (current.type === "bookmarkEnd") return [current];
      return (childNodes(current) ?? []).flatMap(visit);
    };
    return visit(node);
  });
  expect(copied.map(copyOriginAncestorsOf)).toContainEqual([1]);

  const ownIds = [
    ...idsForDate(document, FIRST_DATE),
    ...idsForDate(document, SECOND_DATE),
  ];
  const inReverse = reject(
    reject(document, idsForDate(document, SECOND_DATE)),
    idsForDate(document, FIRST_DATE),
  );
  expect(project(reject(document, ownIds))).toEqual(project(inReverse));
  expect(project(reject(document, allIds(document)))).toEqual(
    project(reject(source, allIds(source))),
  );
});

test("rejecting copied move markers preserves unrelated same-side source markers", () => {
  const source = documentOf([
    {
      type: "insertion",
      info: { id: 1, author: "A" },
      content: [{ type: "moveToRangeEnd", id: 1 }],
    },
  ]);
  const first = apply(source, {
    type: "insertText",
    at: { story: "main", blockId: "00000001", offset: 0 },
    text: "xxx",
    runProps: {},
    revision: { id: 2, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  });
  const second = apply(first.document, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 1 },
    slice: {
      content: [{ type: "moveToRangeEnd", id: 1 }],
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 3, author: "C", date: SECOND_DATE },
    newIds: { revision: [200, 201, 202, 203, 204, 205] },
  });
  const ownIds = [
    ...idsForDate(second.document, FIRST_DATE),
    ...idsForDate(second.document, SECOND_DATE),
  ];
  const inReverse = reject(
    reject(second.document, idsForDate(second.document, SECOND_DATE)),
    idsForDate(second.document, FIRST_DATE),
  );
  expect(project(reject(second.document, ownIds))).toEqual(project(inReverse));
});

test("an aliased source marker is not mistaken for its copied occurrence", () => {
  const marker = { type: "bookmarkEnd", id: 3 } satisfies ParagraphContent;
  const source = documentOf([
    marker,
    { type: "run", content: [{ type: "text", text: "aa" }] },
  ]);
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 0 },
    slice: { content: [marker], openStart: 0, openEnd: 0 },
    revision: { id: 2, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  }).document;
  const markers = bookmarkEnds(paragraphOf(inserted).content);
  expect(markers).toHaveLength(2);
  expect(markers.filter((node) => node === marker)).toHaveLength(1);
  expect(copyOriginAncestorsOf(marker)).toBeUndefined();
  expect(markers.map(copyOriginAncestorsOf).filter(Boolean)).toEqual([[0]]);
});

test("a block-level marker participates in the story copy census", () => {
  const blockMarker = { type: "bookmarkEnd", id: 3 } satisfies ParagraphContent;
  const source: Document = {
    package: {
      document: {
        content: [
          blockMarker,
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "aa" }] }],
          },
        ],
      },
    },
  };
  expect(
    copyOriginBoundaryKeysInBlocks(source.package.document.content),
  ).toContain("bookmarkEnd:3");
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 0 },
    slice: { content: [blockMarker], openStart: 0, openEnd: 0 },
    revision: { id: 2, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  }).document;
  const paragraph = inserted.package.document.content.find(
    (block) => block.type === "paragraph",
  );
  if (paragraph?.type !== "paragraph")
    throw new TypeError("Expected the inserted paragraph.");
  const paragraphMarkers = bookmarkEnds(paragraph.content);
  expect(copyOriginAncestorsOf(blockMarker)).toBeUndefined();
  expect(paragraphMarkers.map(copyOriginAncestorsOf)).toEqual([[0]]);
});

test("a copied moved boundary resolves at its original position in either batch order", () => {
  const original = documentOf([
    { type: "run", content: [{ type: "text", text: "aa" }] },
  ]);
  const defaults = {
    kind: 0,
    block: 0,
    first: 0,
    second: 0,
    third: 0,
    zeroWidth: undefined,
    depth: 1,
    text: "x",
    formatting: {},
    inherit: false,
    runPatch: {},
    paragraphPatch: {},
    newParagraph: undefined,
    content: [],
    fresh: 268435456,
  } satisfies OpSeed;
  const seeds = [
    {
      ...defaults,
      kind: 606232908,
      third: 1,
      content: [
        {
          type: "hyperlink",
          href: "https://example.org",
          rId: "rId9",
          children: [
            {
              type: "inlineWrapper",
              kind: "bidi",
              control: "embedding",
              content: [{ type: "moveToRangeEnd", id: 1 }],
            },
          ],
        },
      ],
    },
    { ...defaults, kind: 1102199946, first: 1242241877, inherit: true },
    defaults,
    defaults,
  ] satisfies OpSeed[];
  let current = original;
  const stamps: RevisionStamp[] = [];
  for (const [index, seed] of seeds.entries()) {
    const operation = trackedOpFor(current, seed, index);
    const applied = applyDocumentOp(current, operation);
    if (applied.isErr()) throw applied.error;
    current = applied.value.document;
    const stamp = stampOf(operation);
    if (stamp !== undefined) stamps.push(stamp);
  }
  const paragraphs = storyParagraphs(current.package.document).map(
    ({ paragraph }) => paragraph,
  );
  const perOperation = stamps.map((stamp) =>
    stampedRevisionIds(paragraphs, stamp, new Set()),
  );
  const rejectStamped = (document: Document, ids: readonly number[]) => {
    const present = new Set(allIds(document));
    const revisionIds = ids.filter((id) => present.has(id));
    if (revisionIds.length === 0) return document;
    return apply(document, {
      type: "resolveRevision",
      story: "main",
      revisionIds,
      decision: "reject",
    }).document;
  };
  const batch = rejectStamped(current, perOperation.flat());
  let reverse = current;
  for (const ids of perOperation.toReversed())
    reverse = rejectStamped(reverse, ids);
  expect(project(batch)).toEqual(project(reverse));
});

test("accepting a copied insertion removes only its owner ordinal", () => {
  const { document } = copiedDocument();
  const accepted = accept(document, idsForDate(document, SECOND_DATE));
  expect(hasValidCopyOriginAncestors(paragraphOf(accepted).content)).toBe(true);
  expect(
    bookmarkEnds(paragraphOf(accepted).content).map(copyOriginAncestorsOf),
  ).toEqual([undefined, undefined]);
});

test("removing a paragraph drops copied owners before hoisting its boundaries", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            pPrMark: {
              kind: "del",
              info: { id: 1, author: "A" },
            },
            content: [
              {
                type: "insertion",
                info: { id: 2, author: "B" },
                content: [
                  { type: "bookmarkEnd", id: 3, copyOriginAncestors: [0] },
                  { type: "bookmarkStart", id: 4 },
                  { type: "bookmarkEnd", id: 4 },
                ],
              },
            ],
          },
          {
            type: "table",
            rows: [
              {
                type: "tableRow",
                cells: [
                  {
                    type: "tableCell",
                    content: [
                      {
                        type: "paragraph",
                        paraId: "00000002",
                        content: [],
                      },
                    ],
                  },
                ],
              },
            ],
          },
          { type: "paragraph", paraId: "00000003", content: [] },
        ],
      },
    },
  };

  const result = applyDocumentOp(document, {
    type: "resolveRevision",
    story: "main",
    revisionIds: [1],
    decision: "accept",
  });

  if (result.isErr()) throw result.error;
  const blocks = result.value.document.package.document.content;
  expect(blocks.map((block) => block.type)).toEqual([
    "bookmarkStart",
    "bookmarkEnd",
    "table",
    "paragraph",
  ]);
});

test("nested tracked copies retain each distinct added-ancestor ordinal", () => {
  const { document } = copiedDocument();
  const secondCopy = apply(document, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 0 },
    slice: {
      content: paragraphOf(document).content,
      openStart: 0,
      openEnd: 0,
    },
    revision: { id: 4, author: "C", date: "2026-03-04T05:00:02Z" },
    newIds: { revision: [300, 301, 302, 303, 304, 305, 306, 307] },
  }).document;
  const originLists = bookmarkEnds(paragraphOf(secondCopy).content)
    .map(copyOriginAncestorsOf)
    .filter((origins) => origins !== undefined);
  expect(hasValidCopyOriginAncestors(paragraphOf(secondCopy).content)).toBe(
    true,
  );
  expect(originLists.some((origins) => origins.length >= 2)).toBe(true);
});

test("a same-stamp copy inserted into an existing insertion records that owner", () => {
  const source = documentOf([
    { type: "bookmarkEnd", id: 3 },
    {
      type: "insertion",
      info: { id: 2, author: "C", date: FIRST_DATE },
      content: [{ type: "run", content: [{ type: "text", text: "xy" }] }],
    },
  ]);
  const inserted = apply(source, {
    type: "insertContent",
    at: { story: "main", blockId: "00000001", offset: 1 },
    slice: {
      content: [
        {
          type: "insertion",
          info: { id: 2, author: "C", date: FIRST_DATE },
          content: [{ type: "bookmarkEnd", id: 3 }],
        },
      ],
      openStart: 1,
      openEnd: 1,
    },
    revision: { id: 4, author: "C", date: FIRST_DATE },
    newIds: { revision: [100, 101, 102, 103] },
  }).document;
  const origins = bookmarkEnds(paragraphOf(inserted).content)
    .map(copyOriginAncestorsOf)
    .filter((value) => value !== undefined);
  expect(origins).toContainEqual([0]);
  expect(hasValidCopyOriginAncestors(paragraphOf(inserted).content)).toBe(true);
});

test("splitting an added wrapper transfers copied ownership to the replacement", () => {
  const marker = {
    type: "bookmarkEnd",
    id: 3,
    copyOriginAncestors: [0],
  } satisfies ParagraphContent;
  const existingInsertion = {
    type: "insertion",
    info: { id: 10, author: "A", date: FIRST_DATE },
    content: [
      { type: "run", content: [{ type: "text", text: "ab" }] },
      marker,
      { type: "run", content: [{ type: "text", text: "cd" }] },
    ],
  } satisfies ParagraphContent;
  const wrapped = wrapTracked({
    items: [
      { type: "run", content: [{ type: "text", text: "x" }] },
      existingInsertion,
      { type: "run", content: [{ type: "text", text: "z" }] },
    ],
    from: { offset: 0, zeroWidthBefore: 0 },
    to: { offset: 3, zeroWidthBefore: Number.MAX_SAFE_INTEGER },
    kind: WRAP_KINDS.INSERTION,
    stamp: { id: 20, author: "C", date: SECOND_DATE },
    copiedBoundaries: new Set([marker]),
  });
  if (wrapped.kind !== "wrapped")
    throw new TypeError("Expected the crossing insertion to be wrapped.");
  const findReplacement = (
    nodes: readonly InlineNode[],
  ): InlineNode | undefined => {
    for (const node of nodes) {
      if (
        node.type === "insertion" &&
        node.info.id === 20 &&
        bookmarkEnds(node.content).some((boundary) => boundary.id === 3)
      )
        return node;
      const nested = findReplacement(childNodes(node) ?? []);
      if (nested !== undefined) return nested;
    }
    return undefined;
  };
  const replacement = findReplacement(wrapped.content);
  if (replacement?.type !== "insertion")
    throw new TypeError("Expected a replacement insertion around the marker.");
  expect(bookmarkEnds(replacement.content).map(copyOriginAncestorsOf)).toEqual([
    [0],
  ]);
  expect(
    bookmarkEnds(acceptCopyOriginAncestor(replacement.content)).map(
      copyOriginAncestorsOf,
    ),
  ).toEqual([undefined]);
  expect(
    bookmarkEnds(rejectCopyOriginAncestorsInNodes(replacement.content, 1)),
  ).toHaveLength(0);
});

test("an open deletion slice and its inverse preserve copied-boundary ownership", () => {
  const { document } = copiedDocument();
  const paragraph = paragraphOf(document);
  const marker = leafSpans(paragraph.content).find(
    ({ node }) =>
      node.type === "bookmarkEnd" && copyOriginAncestorsOf(node) !== undefined,
  );
  if (marker === undefined) throw new TypeError("Expected a copied boundary.");
  const cut = deleteBetween(paragraph.content, marker.before, marker.after);
  expect(cut.removed.openStart + cut.removed.openEnd).toBeGreaterThan(0);
  expect(hasValidCopyOriginAncestors(cut.removed.content)).toBe(true);
  const restored = insertSliceAt(cut.content, marker.before, cut.removed);
  expect(restored).toStrictEqual(paragraph.content);
});

test("copy-origin arrays require sorted unique safe ordinals and a bounded nonempty tuple", () => {
  expect(isCopyOriginAncestors([0])).toBe(true);
  expect(isCopyOriginAncestors([])).toBe(false);
  expect(isCopyOriginAncestors([1, 1])).toBe(false);
  expect(isCopyOriginAncestors([1, 0])).toBe(false);
  expect(isCopyOriginAncestors(Array(1))).toBe(false);
  expect(isCopyOriginAncestors([-1])).toBe(false);
  expect(isCopyOriginAncestors([Number.MAX_SAFE_INTEGER + 1])).toBe(false);
  expect(
    isCopyOriginAncestors(Array.from({ length: 257 }, (_, index) => index)),
  ).toBe(false);
});

test("shared owner transforms remove accepted owners and drop rejected copies", () => {
  expect(addCopyOriginAncestor([0, 2], 1)).toEqual([0, 1, 3]);
  expect(addCopyOriginAncestor(undefined, 2)).toEqual([2]);
  expect(shiftCopyOriginAncestors([0, 2], 1)).toEqual([0, 3]);
  expect(acceptCopyOriginAncestors([0, 2], 0)).toEqual([1]);
  expect(acceptCopyOriginAncestors([1, 3], 1)).toEqual([2]);
  expect(rejectCopyOriginAncestors([0, 2], 1)).toEqual({ type: "drop" });
  expect(rejectCopyOriginAncestors([2, 4], 2)).toEqual({
    type: "retain",
    copyOriginAncestors: [0, 2],
  });
});

test(
  "copy-origin ordinal transforms preserve ownership across add, shift, accept, and reject",
  () => {
    const ancestry = fc.integer({ min: 1, max: 12 }).chain((depth) =>
      fc
        .option(
          fc.subarray(
            Array.from({ length: depth }, (_, index) => index),
            {
              minLength: 1,
            },
          ),
          { nil: undefined },
        )
        .chain((origins) =>
          fc.integer({ min: 0, max: depth }).map((ordinal) => ({
            depth,
            origins,
            ordinal,
          })),
        ),
    );

    assertProperty(
      fc.property(ancestry, ({ origins, ordinal }) => {
        if (origins !== undefined && !isCopyOriginAncestors(origins))
          throw new TypeError("Expected sorted unique owner ordinals.");

        const added = addCopyOriginAncestor(origins, ordinal);
        if (added === undefined)
          throw new TypeError(
            "A bounded owner addition must be representable.",
          );
        expect(acceptCopyOriginAncestors(added, ordinal)).toEqual(origins);

        const shifted = shiftCopyOriginAncestors(origins, ordinal);
        expect(acceptCopyOriginAncestors(shifted, ordinal)).toEqual(origins);

        const rejected = rejectCopyOriginAncestors(origins, ordinal);
        if (origins === undefined) {
          expect(rejected).toEqual({ type: "retain" });
        } else if (origins.some((owner) => owner < ordinal)) {
          expect(rejected).toEqual({ type: "drop" });
        } else {
          expect(rejected).toEqual({
            type: "retain",
            copyOriginAncestors: origins.map((owner) => owner - ordinal),
          });
        }
      }),
      {
        numRuns: 100,
        id: "copy-origin ordinal transforms preserve add-shift-accept-reject laws",
      },
    );
  },
  propertyTestTimeout(10_000),
);

test("copy-origin transforms refuse ownership beyond the configured cap", () => {
  const maximumOrigins = Array.from({ length: 256 }, (_, index) => index);
  if (!isCopyOriginAncestors(maximumOrigins))
    throw new TypeError("The configured cap must be valid.");
  expect(addCopyOriginAncestor(maximumOrigins, 256)).toBeUndefined();
});

test("the operations contract rejects copied ownership on block-level boundaries", () => {
  const document: Document = {
    package: {
      document: {
        content: [
          { type: "bookmarkEnd", id: 3, copyOriginAncestors: [0] },
          {
            type: "paragraph",
            paraId: "00000001",
            content: [{ type: "run", content: [{ type: "text", text: "a" }] }],
          },
        ],
      },
    },
  };
  expect(contractViolation(document)?.message).toContain(
    "copied range boundary has invalid added-revision ancestry",
  );
});
