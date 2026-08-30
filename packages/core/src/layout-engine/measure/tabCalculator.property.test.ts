import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../../test/property-testing";

import { computeTabStops } from "./tabCalculator";

/** Ten inches of stops at the smallest accepted interval of 20 twips. */
const MAX_DEFAULT_TAB_STOP_COUNT = 720;

describe("computeTabStops (properties)", () => {
  test("returns a bounded set of stops for every numeric default interval", () => {
    fc.assert(
      fc.property(fc.double(), (defaultTabInterval) => {
        const stops = computeTabStops({ defaultTabInterval });

        expect(stops.length).toBeLessThanOrEqual(MAX_DEFAULT_TAB_STOP_COUNT);
      }),
      propertyConfig({ numRuns: 1000 }),
    );
  });
});
