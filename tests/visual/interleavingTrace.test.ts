import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../test/property-testing";
import { interleavingTraceArbitrary } from "./interleavingTrace";

setDefaultTimeout(propertyTestTimeout(5_000));

test("interleaving traces replay and preserve pending AI/human overlap when shrinking", () => {
  const config = { seed: 197, numRuns: 300 };
  const traces = fc.sample(interleavingTraceArbitrary, config);
  expect(fc.sample(interleavingTraceArbitrary, config)).toEqual(traces);
  for (const trace of traces) {
    expect(trace.actions.at(0)?.kind).toBe("suggest");
    expect(trace.actions.at(1)?.kind).toBe("typing");
  }
  expect(new Set(traces.flatMap(({ actions }) => actions.map(({ kind }) => kind)))).toEqual(
    new Set(["suggest", "typing", "accept", "reject", "undo", "redo"]),
  );
  const failure = fc.check(
    fc.property(interleavingTraceArbitrary, () => false),
    propertyConfig({ seed: 197 }),
  );
  expect(failure.failed).toBe(true);
  const minimal = failure.counterexample?.at(0);
  expect(minimal?.actions).toHaveLength(3);
  expect(minimal?.actions.at(0)?.kind).toBe("suggest");
  expect(minimal?.actions.at(1)?.kind).toBe("typing");
});
