import { expect, setDefaultTimeout, test } from "bun:test";
import { propertyTestTimeout } from "../test/property-testing";
import fc from "fast-check";
import {
  CANONICAL_BROWSER_HISTORY_REPLAYS,
  CANONICAL_BROWSER_SAVE_REPLAYS,
  canonicalBrowserTraceArbitrary,
} from "../tests/visual/canonicalBrowserTrace";

setDefaultTimeout(propertyTestTimeout(5_000));

test("canonical browser regression paths still exercise each reported action sequence", () => {
  const replays = [...CANONICAL_BROWSER_HISTORY_REPLAYS, ...CANONICAL_BROWSER_SAVE_REPLAYS];
  let cases = 0;
  for (const { seed, path, kinds } of replays) {
    const traces = fc.sample(canonicalBrowserTraceArbitrary, { seed, path, numRuns: 1 });
    expect(traces).toHaveLength(1);
    for (const actions of traces) {
      cases += 1;
      expect(actions.map(({ kind }) => kind)).toEqual(kinds);
    }
  }
  expect(cases).toBe(replays.length);
  expect(cases).toBeGreaterThan(0);
});
