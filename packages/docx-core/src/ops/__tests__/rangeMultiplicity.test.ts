import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import {
  type BoundaryOccurrence,
  classifyRangeMultiplicity,
  planWellFormedRanges,
  rangeBoundaryDefinitions,
} from "../rangeBoundaries";

setDefaultTimeout(propertyTestTimeout(30_000));

const BOUNDARY_SIDES = rangeBoundaryDefinitions();

const marker = (
  type: keyof typeof BOUNDARY_SIDES,
  kept: boolean,
): BoundaryOccurrence => ({
  range: `${BOUNDARY_SIDES[type].family}:3`,
  side: BOUNDARY_SIDES[type].side,
  kept,
  copied: false,
});

test("standing duplicate ends remain ambiguous when a rejection keeps their start", () => {
  const occurrences = [
    marker("bookmarkEnd", false),
    marker("bookmarkEnd", false),
    marker("bookmarkStart", true),
  ];
  expect(classifyRangeMultiplicity(occurrences).get("bookmark:3")).toEqual({
    type: "ambiguous",
  });
  expect(planWellFormedRanges(occurrences)).toEqual({ moved: [] });
});

test("every kept boundary stays, whatever other occurrences of it the story holds", () => {
  const copy = (
    type: keyof typeof BOUNDARY_SIDES,
    kept: boolean,
  ): BoundaryOccurrence => ({ ...marker(type, kept), copied: true });
  // Whether a kept boundary stays decides nothing about the others: their
  // own revisions, resolved before or after, decide theirs.
  for (const occurrences of [
    [copy("bookmarkStart", false), marker("bookmarkStart", true)],
    [copy("bookmarkStart", true), marker("bookmarkStart", true)],
    [marker("bookmarkStart", false), marker("bookmarkStart", true)],
    [marker("bookmarkStart", true), marker("bookmarkStart", true)],
  ])
    expect(planWellFormedRanges(occurrences)).toEqual({ moved: [] });
  // A copy an added revision owns orders nothing: the kept start closes on
  // the end, not on the copy before it.
  expect(
    planWellFormedRanges([
      copy("bookmarkStart", false),
      marker("bookmarkEnd", false),
      marker("bookmarkStart", true),
    ]).moved,
  ).toEqual([{ moving: 2, anchor: 1, side: "before" }]);
});

test("a unique reversed pair moves only the kept half", () => {
  expect(
    planWellFormedRanges([
      marker("bookmarkEnd", false),
      marker("bookmarkStart", true),
    ]),
  ).toEqual({ moved: [{ moving: 1, anchor: 0, side: "before" }] });
  expect(
    planWellFormedRanges([
      marker("bookmarkEnd", true),
      marker("bookmarkStart", false),
    ]),
  ).toEqual({ moved: [{ moving: 0, anchor: 1, side: "after" }] });
});

const normalize = (occurrences: readonly BoundaryOccurrence[]) => {
  const plan = planWellFormedRanges(occurrences);
  const normalized = [...occurrences];
  for (const { moving, anchor, side } of plan.moved) {
    const movingNode = occurrences.at(moving);
    const anchorNode = occurrences.at(anchor);
    expect(movingNode).toBeDefined();
    expect(anchorNode).toBeDefined();
    if (movingNode === undefined || anchorNode === undefined) continue;
    normalized.splice(normalized.indexOf(movingNode), 1);
    const insertion =
      normalized.indexOf(anchorNode) + (side === "after" ? 1 : 0);
    normalized.splice(insertion, 0, movingNode);
  }
  return normalized;
};

const byIdentity = (left: object, right: object): number =>
  JSON.stringify(left).localeCompare(JSON.stringify(right));

const occurrenceArbitrary = fc
  .record({
    definition: fc.constantFrom(...Object.entries(BOUNDARY_SIDES)),
    id: fc.integer({ min: 0, max: 5 }),
    kept: fc.boolean(),
    copied: fc.boolean(),
  })
  .map(({ definition: [, { family, side }], id, kept, copied }) => ({
    range: `${family}:${id}`,
    side,
    kept,
    copied,
  }));

test("normalization keeps every occurrence, moves only kept ones, and reaches a fixed point", () => {
  assertProperty(
    fc.property(
      fc.array(occurrenceArbitrary, { maxLength: 60 }),
      (occurrences) => {
        const normalized = normalize(occurrences);
        // Nothing goes, so no occurrence's survival depends on another's.
        expect(normalized.toSorted(byIdentity)).toEqual(
          occurrences.toSorted(byIdentity),
        );
        expect(normalized.filter(({ kept }) => !kept)).toEqual(
          occurrences.filter(({ kept }) => !kept),
        );
        const keptRanges = new Set(
          normalized.filter(({ kept }) => kept).map(({ range }) => range),
        );
        // A copy an added revision owns orders no range.
        const ownedCopies = new Set(
          normalized.flatMap(({ kept, copied }, index) =>
            !kept && copied ? [index] : [],
          ),
        );
        for (const [range, multiplicity] of classifyRangeMultiplicity(
          normalized,
          ownedCopies,
        )) {
          if (multiplicity.type !== "pair" || !keptRanges.has(range)) continue;
          expect(multiplicity.start).toBeLessThan(multiplicity.end);
        }
        expect(planWellFormedRanges(normalized)).toEqual({ moved: [] });
      },
    ),
    {
      numRuns: 300,
      id: "normalization keeps every occurrence, moves only kept ones, and reaches a fixed point",
    },
  );
});

test("of a kept start and end out of order, the one kept from the later revision moves", () => {
  const dated = (
    type: keyof typeof BOUNDARY_SIDES,
    keptFrom: string,
  ): BoundaryOccurrence => ({ ...marker(type, true), keptFrom });
  // The start came from the later revision: it moves before the end.
  expect(
    planWellFormedRanges([
      dated("bookmarkEnd", ""),
      dated("bookmarkStart", "2026-03-04T05:00:00Z"),
    ]).moved,
  ).toEqual([{ moving: 1, anchor: 0, side: "before" }]);
  // The end came from the later revision, or both are as old: it moves after the start.
  for (const endFrom of ["2026-03-04T05:00:00Z", ""])
    expect(
      planWellFormedRanges([
        dated("bookmarkEnd", endFrom),
        dated("bookmarkStart", ""),
      ]).moved,
    ).toEqual([{ moving: 0, anchor: 1, side: "after" }]);
});
