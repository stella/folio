import { expect, test } from "bun:test";
import fc from "fast-check";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "./canonicalBrowserTrace";

test("canonical browser regression paths still exercise each reported action sequence", () => {
  let cases = 0;
  for (const { seed, path, kinds } of CANONICAL_BROWSER_HISTORY_REPLAYS) {
    const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
    expect(traces).toHaveLength(1);
    for (const actions of traces) {
      cases += 1;
      expect(actions.map(({ kind }) => kind)).toEqual(kinds);
    }
  }
  expect(cases).toBe(CANONICAL_BROWSER_HISTORY_REPLAYS.length);
  expect(cases).toBeGreaterThan(0);
});
