import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  type CanonicalInlineGap,
  CanonicalInlineGaps,
  type CanonicalSeamExtent,
} from "./canonicalInlineGaps";

setDefaultTimeout(propertyTestTimeout(30_000));

type Add = { offset: number; position: number; zeroWidthBefore: number };

/** Per-offset lists of gaps and seams, as a paragraph's projection kept them before the compact table. */
const reference = (
  offsets: number,
  adds: readonly Add[],
  preferred: (offset: number) => number,
) => {
  const boundaries: CanonicalInlineGap[][] = Array.from(
    { length: offsets },
    () => [],
  );
  const seams: CanonicalSeamExtent[][] = Array.from(
    { length: offsets },
    () => [],
  );
  for (const { offset, position, zeroWidthBefore } of adds) {
    const gaps = boundaries[offset];
    const extents = seams[offset];
    if (!gaps || !extents) continue;
    const existing = gaps.findIndex((entry) => entry.position === position);
    const entry = { position, zeroWidthBefore };
    if (existing < 0) {
      gaps.push(entry);
      extents.push({ first: zeroWidthBefore, last: zeroWidthBefore });
      continue;
    }
    const extent = extents[existing];
    const kept = gaps[existing];
    if (!extent || !kept) continue;
    extent.first = Math.min(extent.first, zeroWidthBefore);
    extent.last = Math.max(extent.last, zeroWidthBefore);
    const target = preferred(offset);
    if (
      Math.abs(zeroWidthBefore - target) <
      Math.abs(kept.zeroWidthBefore - target)
    )
      gaps[existing] = entry;
  }
  return { boundaries, seams };
};

test("the compact gap table answers what per-offset gap and seam lists answered", () => {
  assertProperty(
    fc.property(
      fc.integer({ min: 1, max: 12 }).chain((offsets) =>
        fc.record({
          offsets: fc.constant(offsets),
          adds: fc.array(
            fc.record({
              offset: fc.integer({ min: 0, max: offsets - 1 }),
              position: fc.integer({ min: 0, max: 6 }),
              zeroWidthBefore: fc.integer({ min: 0, max: 4 }),
            }),
            { maxLength: 40 },
          ),
          preferred: fc.array(fc.integer({ min: 0, max: 4 }), {
            minLength: offsets,
            maxLength: offsets,
          }),
        }),
      ),
      ({ offsets, adds, preferred }) => {
        const prefer = (offset: number) => preferred[offset] ?? 0;
        const table = new CanonicalInlineGaps(offsets);
        for (const { offset, position, zeroWidthBefore } of adds)
          expect(
            table.add(offset, position, zeroWidthBefore, () => prefer(offset)),
          ).toBe(true);
        const expected = reference(offsets, adds, prefer);
        expect(table.offsets).toBe(offsets);
        expect(table.toBoundaries()).toEqual(expected.boundaries);
        expect(table.toSeams()).toEqual(expected.seams);
        expect(table.hasEmptyOffset()).toBe(
          expected.boundaries.some((gaps) => gaps.length === 0),
        );
        for (let position = 0; position <= 7; position += 1)
          expect(table.offsetOf(position)).toBe(
            expected.boundaries.findIndex((gaps) =>
              gaps.some((gap) => gap.position === position),
            ),
          );
      },
    ),
    {
      numRuns: 200,
      id: "the compact gap table answers what per-offset gap and seam lists answered",
    },
  );
});

test("an offset outside the text is refused and reads as no gaps", () => {
  const table = new CanonicalInlineGaps(2);
  expect(table.add(2, 0, 0, () => 0)).toBe(false);
  expect(table.add(-1, 0, 0, () => 0)).toBe(false);
  expect(table.at(2)).toBeUndefined();
  expect(table.at(-1)).toBeUndefined();
  expect(table.at(0)).toEqual([]);
});

test("a gap read from the table is a copy", () => {
  const table = new CanonicalInlineGaps(1);
  table.add(0, 0, 0, () => 0);
  table.add(0, 1, 1, () => 0);
  const [first, second] = table.at(0) ?? [];
  if (first) first.position = 9;
  if (second) second.position = 9;
  expect(table.toBoundaries()).toEqual([
    [
      { position: 0, zeroWidthBefore: 0 },
      { position: 1, zeroWidthBefore: 1 },
    ],
  ]);
});
