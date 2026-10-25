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
  boundary: `${type}:3`,
  range: `${BOUNDARY_SIDES[type].family}:3`,
  side: BOUNDARY_SIDES[type].side,
  kept,
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
  expect(planWellFormedRanges(occurrences)).toEqual({ dropped: [], moved: [] });
});

test("a unique reversed pair moves only the kept half", () => {
  expect(
    planWellFormedRanges([
      marker("bookmarkEnd", false),
      marker("bookmarkStart", true),
    ]),
  ).toEqual({
    dropped: [],
    moved: [{ moving: 1, anchor: 0, side: "before" }],
  });
  expect(
    planWellFormedRanges([
      marker("bookmarkEnd", true),
      marker("bookmarkStart", false),
    ]),
  ).toEqual({
    dropped: [],
    moved: [{ moving: 0, anchor: 1, side: "after" }],
  });
});

const normalize = (occurrences: readonly BoundaryOccurrence[]) => {
  const plan = planWellFormedRanges(occurrences);
  const dropped = new Set(plan.dropped);
  const normalized = occurrences.filter((_, index) => !dropped.has(index));
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

const occurrenceArbitrary = fc
  .record({
    definition: fc.constantFrom(...Object.entries(BOUNDARY_SIDES)),
    id: fc.integer({ min: 0, max: 5 }),
    kept: fc.boolean(),
  })
  .map(({ definition: [type, { family, side }], id, kept }) => ({
    boundary: `${type}:${id}`,
    range: `${family}:${id}`,
    side,
    kept,
  }));

test("normalization preserves standing occurrences, retained quotas, and reaches a fixed point", () => {
  assertProperty(
    fc.property(
      fc.array(occurrenceArbitrary, { maxLength: 60 }),
      (occurrences) => {
        const normalized = normalize(occurrences);
        expect(normalized.filter(({ kept }) => !kept)).toEqual(
          occurrences.filter(({ kept }) => !kept),
        );
        for (const boundary of new Set(
          occurrences.map((item) => item.boundary),
        )) {
          const source = occurrences.filter(
            (item) => item.boundary === boundary,
          );
          const standing = source.filter(({ kept }) => !kept).length;
          const expectedCount = standing > 0 ? standing : 1;
          expect(
            normalized.filter((item) => item.boundary === boundary).length,
          ).toBe(expectedCount);
        }
        const keptRanges = new Set(
          normalized.filter(({ kept }) => kept).map(({ range }) => range),
        );
        for (const [range, multiplicity] of classifyRangeMultiplicity(
          normalized,
        )) {
          if (multiplicity.type !== "pair" || !keptRanges.has(range)) continue;
          expect(multiplicity.start).toBeLessThan(multiplicity.end);
        }
        expect(planWellFormedRanges(normalized)).toEqual({
          dropped: [],
          moved: [],
        });
      },
    ),
    {
      numRuns: 300,
      id: "normalization preserves standing occurrences, retained quotas, and reaches a fixed point",
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
