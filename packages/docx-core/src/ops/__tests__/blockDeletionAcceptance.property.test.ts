/**
 * Accepting a tracked deletion of whole paragraphs equals deleting them
 * directly, and rejecting it restores the original, for the bookmarks,
 * comment ranges and move ranges the paragraphs hold: ranges wholly inside
 * them and ranges continuing before or after them. What a deleted paragraph
 * held goes to the nearest paragraph that stays, never between blocks, and
 * the edit reports that paragraph. Before a table, where no direct deletion
 * applies, acceptance keeps every boundary in order and leaves none between
 * blocks.
 */

import { expect, mock, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { checkCoverageSample } from "../../../../../test/property-coverage";
import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  BlockContent,
  Document,
  Paragraph,
  ParagraphContent,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import * as rangeBoundariesModule from "../rangeBoundaries";
import { isBlockBoundary, rangeBoundariesIn } from "../rangeBoundaries";
import * as reviewModule from "../review";
import {
  type DeleteBlocksOp,
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

/** The coverage sample's run count. */
const COVERAGE_RUNS = 200;

const FAMILIES = ["bookmark", "comment", "move"] as const;
type Family = (typeof FAMILIES)[number];

/** Where a range lies against the deleted paragraphs. */
const RELATIONS = ["inside", "continuing-before", "continuing-after"] as const;
type Relation = (typeof RELATIONS)[number];

const FOLLOWERS = ["paragraph", "table", "nothing"] as const;
type Follower = (typeof FOLLOWERS)[number];

type RangeSpec = { family: Family; from: number; to: number };

type Case = {
  prefix: number;
  selected: number;
  follower: Follower;
  ranges: readonly RangeSpec[];
};

const paraIdOf = (index: number): string =>
  (index + 1).toString(16).toUpperCase().padStart(8, "0");

const startOf = (family: Family, id: number): ParagraphContent => {
  switch (family) {
    case "bookmark":
      return { type: "bookmarkStart", id, name: `_Ref${id}` };
    case "comment":
      return { type: "commentRangeStart", id };
    case "move":
      return {
        type: "moveFromRangeStart",
        id,
        name: `move${id}`,
        author: "Author",
      };
  }
};

const endOf = (family: Family, id: number): ParagraphContent[] => {
  switch (family) {
    case "bookmark":
      return [{ type: "bookmarkEnd", id }];
    case "comment":
      // A tracked paragraph deletion takes in no comment reference.
      return [{ type: "commentRangeEnd", id }];
    case "move":
      return [{ type: "moveFromRangeEnd", id }];
  }
};

/** The story paragraphs, the paragraph after a table when one follows. */
const paragraphCount = ({ prefix, selected, follower }: Case): number =>
  prefix + selected + (follower === "nothing" ? 0 : 1);

const documentOf = (sample: Case): Document => {
  const count = paragraphCount(sample);
  const paragraphs: Paragraph[] = Array.from({ length: count }, (_, index) => {
    const id = (range: RangeSpec) => 10 + sample.ranges.indexOf(range);
    const starts = sample.ranges
      .filter(({ from }) => from === index)
      .map((range) => startOf(range.family, id(range)));
    const ends = sample.ranges
      .filter(({ to }) => to === index)
      .flatMap((range) => endOf(range.family, id(range)));
    return {
      type: "paragraph",
      paraId: paraIdOf(index),
      content: [
        ...starts,
        { type: "run", content: [{ type: "text", text: `p${index}` }] },
        ...ends,
      ],
    };
  });
  const before = paragraphs.slice(0, sample.prefix + sample.selected);
  const after = paragraphs.slice(sample.prefix + sample.selected);
  const table: BlockContent = {
    type: "table",
    rows: [
      {
        type: "tableRow",
        cells: [
          {
            type: "tableCell",
            content: [{ type: "paragraph", paraId: "7FFFFF01", content: [] }],
          },
        ],
      },
    ],
  };
  const comments = sample.ranges.flatMap((range, index) =>
    range.family === "comment"
      ? [
          {
            id: 10 + index,
            author: "Reviewer",
            anchorKind: "range" as const,
            content: [],
          },
        ]
      : [],
  );
  return {
    package: {
      document: {
        content:
          sample.follower === "table"
            ? [...before, table, ...after]
            : [...before, ...after],
        ...(comments.length === 0 ? {} : { comments }),
      },
    },
  };
};

const caseArbitrary: fc.Arbitrary<Case> = fc
  .record({
    prefix: fc.integer({ min: 0, max: 2 }),
    selected: fc.integer({ min: 1, max: 3 }),
    follower: fc.constantFrom(...FOLLOWERS),
  })
  .chain((shape) => {
    const last = paragraphCount({ ...shape, ranges: [] }) - 1;
    const range = fc
      .record({
        family: fc.constantFrom(...FAMILIES),
        from: fc.integer({ min: 0, max: last }),
        to: fc.integer({ min: 0, max: last }),
      })
      .map(({ family, from, to }) => ({
        family,
        from: Math.min(from, to),
        to: Math.max(from, to),
      }));
    return fc
      .array(range, { minLength: 1, maxLength: 4 })
      .map((ranges): Case => ({
        prefix: shape.prefix,
        selected: shape.selected,
        follower: shape.follower,
        ranges,
      }));
  });

const relationOf = (
  { from, to }: RangeSpec,
  { prefix, selected }: Case,
): Relation | undefined => {
  const first = prefix;
  const last = prefix + selected - 1;
  if (from >= first && to <= last) return "inside";
  if (from < first && to >= first && to <= last) return "continuing-before";
  if (from >= first && from <= last && to > last) return "continuing-after";
  return undefined;
};

const deletionOf = (sample: Case): DeleteBlocksOp => ({
  type: DOCUMENT_OP_TYPES.DELETE_BLOCKS,
  story: OP_STORIES.MAIN,
  blockIds: Array.from({ length: sample.selected }, (_, index) =>
    paraIdOf(sample.prefix + index),
  ),
  revision: { id: 100, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
  newIds: { revision: Array.from({ length: 64 }, (_, index) => 101 + index) },
});

/** Every range boundary in `blocks`, at any depth, in document order. */
const boundariesIn = (blocks: readonly BlockContent[]): string[] =>
  blocks.flatMap((block): string[] => {
    if (isBlockBoundary(block)) return [`${block.type}:${block.id}`];
    if (block.type === "paragraph")
      return rangeBoundariesIn(block.content).map(
        (boundary) => `${boundary.type}:${boundary.id}`,
      );
    if (block.type === "table")
      return block.rows.flatMap((row) =>
        row.cells.flatMap((cell) => boundariesIn(cell.content)),
      );
    return [];
  });

/** How many boundaries each paragraph directly in `blocks` holds. */
const boundaryCounts = (
  blocks: readonly BlockContent[],
): ReadonlyMap<string, number> =>
  new Map(
    blocks.flatMap((block) =>
      block.type === "paragraph"
        ? [[block.paraId ?? "", rangeBoundariesIn(block.content).length]]
        : [],
    ),
  );

const resolveAll = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) => {
  const resolved = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  });
  if (resolved.isErr()) throw resolved.error;
  return resolved.value;
};

/** The law for one case, recording what it exercised in `exercised`. */
const checkCase = (sample: Case, exercised: Set<string>): void => {
  const document = documentOf(sample);
  const op = deletionOf(sample);
  const tracked = applyDocumentOp(document, op);
  if (tracked.isErr()) return;
  const rejected = resolveAll(
    tracked.value.document,
    tracked.value.revisions,
    REVISION_DECISIONS.REJECT,
  );
  expect(rejected.document).toStrictEqual(document);
  const accepted = resolveAll(
    tracked.value.document,
    tracked.value.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  const undo = applyDocumentOps(accepted.document, accepted.inverse);
  if (undo.isErr()) throw undo.error;
  expect(undo.value.document).toStrictEqual(tracked.value.document);
  const blocks = accepted.document.package.document.content;
  // Nothing a deleted paragraph held stands between blocks.
  expect(blocks.filter(isBlockBoundary)).toEqual([]);
  // A paragraph that took boundaries is one the edit reports.
  const before = boundaryCounts(
    tracked.value.document.package.document.content,
  );
  const reported = new Set([
    ...accepted.touched.modified,
    ...accepted.touched.inserted,
  ]);
  for (const [paraId, count] of boundaryCounts(blocks))
    if (count > (before.get(paraId) ?? 0))
      expect(reported.has(paraId)).toBe(true);
  const direct = applyDocumentOp(document, { ...op, revision: undefined });
  if (direct.isOk()) {
    expect(accepted.document).toStrictEqual(direct.value.document);
  } else {
    // No direct deletion applies before a table: every boundary of a
    // range that survives stays, in order.
    const surviving = new Set(
      (accepted.document.package.document.comments ?? []).map(({ id }) =>
        String(id),
      ),
    );
    const kept = (key: string) =>
      !key.startsWith("comment") || surviving.has(key.split(":")[1] ?? "");
    expect(boundariesIn(blocks)).toEqual(
      boundariesIn(document.package.document.content).filter(kept),
    );
  }
  exercised.add(`${sample.follower}:${direct.isOk() ? "direct" : "kept"}`);
  for (const range of sample.ranges) {
    const relation = relationOf(range, sample);
    if (relation !== undefined) exercised.add(`${range.family}:${relation}`);
  }
};

const law = (exercised: Set<string>) =>
  fc.property(caseArbitrary, (sample) => checkCase(sample, exercised));

/** A comment range starting in a deleted paragraph and ending in the next. */
const COMMENT_CONTINUING_AFTER: Case = {
  prefix: 1,
  selected: 1,
  follower: "paragraph",
  ranges: [{ family: "comment", from: 1, to: 2 }],
};

/** A bookmark pair in a paragraph deleted between a kept one and a table. */
const BEFORE_TABLE: Case = {
  prefix: 1,
  selected: 1,
  follower: "table",
  ranges: [{ family: "bookmark", from: 1, to: 1 }],
};

test("accepting a tracked deletion of whole paragraphs equals deleting them directly, for the ranges they hold", () => {
  assertProperty(law(new Set()), {
    numRuns: 100,
    id: "accepting a tracked deletion of whole paragraphs equals deleting them directly",
  });
  const exercised = new Set<string>();
  checkCoverageSample(law(exercised), COVERAGE_RUNS);
  const required = [
    "paragraph:direct",
    "nothing:direct",
    "table:kept",
    ...FAMILIES.flatMap((family) =>
      RELATIONS.map((relation) => `${family}:${relation}`),
    ),
  ];
  expect(required.filter((kind) => !exercised.has(kind))).toEqual([]);
});

test("accepting the deletion of a paragraph before a table hands its bookmark pair to the paragraph before it", () => {
  const exercised = new Set<string>();
  checkCase(BEFORE_TABLE, exercised);
  expect(exercised.has("table:kept")).toBe(true);
  const tracked = applyDocumentOp(
    documentOf(BEFORE_TABLE),
    deletionOf(BEFORE_TABLE),
  ).unwrap();
  const accepted = resolveAll(
    tracked.document,
    tracked.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  const [receiver] = accepted.document.package.document.content;
  expect(receiver?.type === "paragraph" && receiver.content).toEqual([
    { type: "run", content: [{ type: "text", text: "p0" }] },
    { type: "bookmarkStart", id: 10, name: "_Ref10" },
    { type: "bookmarkEnd", id: 10 },
  ]);
  expect(accepted.touched.modified).toContain(paraIdOf(0));
});

test("leaving a deleted paragraph's bookmarks between blocks fails the case before a table", () => {
  // The rule this file guards, broken on purpose: bookmarks and move ranges
  // of a paragraph that goes stand between blocks where it was.
  const original = { ...rangeBoundariesModule };
  mock.module("../rangeBoundaries", () => ({
    ...original,
    planParagraphBoundaryRemoval: (
      options: Parameters<typeof original.planParagraphBoundaryRemoval>[0],
    ): ReturnType<typeof original.planParagraphBoundaryRemoval> => {
      const paragraph = options.blocks[options.index];
      const carriers =
        paragraph?.type === "paragraph"
          ? original.paragraphRemains(paragraph.content)
          : [];
      if (
        carriers.length === 0 ||
        carriers.some((carrier) => !original.isBoundary(carrier)) ||
        carriers.some(
          (carrier) =>
            carrier.type === "commentRangeStart" ||
            carrier.type === "commentRangeEnd",
        )
      )
        return original.planParagraphBoundaryRemoval(options);
      return {
        type: "removed",
        blocks: [
          ...options.blocks.slice(0, options.index),
          ...carriers.filter(
            (carrier): carrier is rangeBoundariesModule.BlockBoundary =>
              original.isBoundary(carrier),
          ),
          ...options.blocks.slice(options.index + 1),
        ],
        effects: [],
      };
    },
  }));
  try {
    expect(() => checkCase(BEFORE_TABLE, new Set())).toThrow();
  } finally {
    mock.module("../rangeBoundaries", () => original);
  }
});

test("accepting the deletion of a paragraph holding a comment's start hands the start to the paragraph before it", () => {
  const exercised = new Set<string>();
  checkCase(COMMENT_CONTINUING_AFTER, exercised);
  expect(exercised.has("comment:continuing-after")).toBe(true);
  const tracked = applyDocumentOp(
    documentOf(COMMENT_CONTINUING_AFTER),
    deletionOf(COMMENT_CONTINUING_AFTER),
  ).unwrap();
  const accepted = resolveAll(
    tracked.document,
    tracked.revisions,
    REVISION_DECISIONS.ACCEPT,
  );
  expect(boundariesIn(accepted.document.package.document.content)).toEqual([
    "commentRangeStart:10",
    "commentRangeEnd:10",
  ]);
  expect(
    accepted.document.package.document.comments?.map(({ id }) => id),
  ).toEqual([10]);
});

test("refusing a paragraph holding a bare comment anchor leaves comment ranges out of the law", () => {
  // The rule this file guards, broken on purpose: a tracked deletion of
  // whole paragraphs refuses the bare comment anchors it meets instead of
  // leaving them standing, so a comment range never reaches the law.
  const original = { ...reviewModule };
  mock.module("../review", () => ({
    ...original,
    wrapTracked: (
      options: Parameters<typeof original.wrapTracked>[0],
    ): ReturnType<typeof original.wrapTracked> => {
      const { commentAnchors: _standing, ...refusing } = options;
      return original.wrapTracked(refusing);
    },
  }));
  try {
    const exercised = new Set<string>();
    checkCase(COMMENT_CONTINUING_AFTER, exercised);
    expect(exercised.has("comment:continuing-after")).toBe(false);
  } finally {
    mock.module("../review", () => original);
  }
});
