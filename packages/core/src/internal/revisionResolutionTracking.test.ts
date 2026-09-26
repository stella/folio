import { expect, test } from "bun:test";
import fc from "fast-check";
import { StepMap } from "prosemirror-transform";

import { indexedPositionMap } from "./revisionResolutionTracking";

test("indexed position mapping matches StepMap at every changed boundary", () => {
  fc.assert(
    fc.property(
      fc.array(
        fc.record({
          gap: fc.integer({ min: 0, max: 4 }),
          oldSize: fc.integer({ min: 0, max: 5 }),
          newSize: fc.integer({ min: 0, max: 5 }),
        }),
        { maxLength: 24 },
      ),
      (changes) => {
        let position = 0;
        const ranges: number[] = [];
        for (const { gap, oldSize, newSize } of changes) {
          position += gap;
          ranges.push(position, oldSize, newSize);
          position += oldSize;
        }
        const forward = new StepMap(ranges);
        for (const map of [forward, forward.invert()]) {
          const indexed = indexedPositionMap(map);
          const limit = Math.max(position, map.map(position, 1)) + 8;
          for (let point = 0; point <= limit; point++) {
            expect(indexed(point, -1)).toBe(map.map(point, -1));
            expect(indexed(point, 1)).toBe(map.map(point, 1));
            for (const assoc of [-1, 1] as const) {
              const actual = indexed.mapResult(point, assoc);
              const expected = map.mapResult(point, assoc);
              expect(actual.pos).toBe(expected.pos);
              expect(actual.deletedAcross).toBe(expected.deletedAcross);
            }
          }
        }
      },
    ),
    { seed: 2_609_261, numRuns: 100 },
  );
});
