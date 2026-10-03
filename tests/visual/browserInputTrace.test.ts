import type { BrowserDragTarget } from "./browserDragTarget";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../test/property-testing";

import { documentShape } from "../../packages/core/src/__tests__/documentShapes";
import { SUGGESTION_INPUT_KINDS } from "../../packages/core/src/__tests__/suggestionInputKinds";
import {
  BROWSER_SHAPE_TARGETS,
  BROWSER_TRACE_FIXED_SEEDS,
  browserImeActionArbitrary,
  browserInputTraceArbitrary,
  browserSuggestionActionKinds,
  parseBrowserInputTraceConfig,
} from "./browserInputTrace";

setDefaultTimeout(propertyTestTimeout(5_000));

test("browser traces replay deterministically and exercise the declared shapes and input kinds", () => {
  const arbitrary = browserInputTraceArbitrary(12);
  const config = { seed: 197, numRuns: 800 };
  const traces = fc.sample(arbitrary, config);
  expect(fc.sample(arbitrary, config)).toEqual(traces);
  expect(new Set(traces.map(({ shape }) => shape))).toEqual(
    new Set(Object.keys(BROWSER_SHAPE_TARGETS)),
  );
  expect(new Set(browserSuggestionActionKinds)).toEqual(new Set(SUGGESTION_INPUT_KINDS));
  const kinds = new Set(traces.flatMap(({ actions }) => actions.map(({ kind }) => kind)));
  for (const kind of [...SUGGESTION_INPUT_KINDS, "undo", "redo", "historyBurst", "selectionDrag"]) {
    expect(kinds.has(kind)).toBe(true);
  }
  const targetFeatures = {
    table: "table",
    list: "list-decimal",
    note: "footnote",
    field: "field",
    inlineObject: "image",
  } as const satisfies Record<BrowserDragTarget, string>;
  for (const { shape, actions } of traces) {
    const fixture = documentShape(shape);
    for (const action of actions) {
      if (action.kind === "selectionDrag") {
        expect(action.target).toBe(BROWSER_SHAPE_TARGETS[shape]);
        if (action.target === "list") {
          expect(
            fixture.features.some(
              (feature) => feature === "list-decimal" || feature === "list-bullet",
            ),
          ).toBe(true);
        } else {
          expect(fixture.features).toContain(targetFeatures[action.target]);
        }
      }
      if (action.kind === "dragCellDelete") expect(fixture.features).toContain("table");
    }
  }
});

test("IME generator covers cancellation, repeated updates and Unicode with valid shrinking", () => {
  const actions = fc.sample(browserImeActionArbitrary, { seed: 431, numRuns: 300 });
  expect(new Set(actions.map(({ completion }) => completion))).toEqual(
    new Set(["commit", "cancel"]),
  );
  expect(actions.some(({ updates }) => updates.length > 1)).toBe(true);
  for (const text of ["東京", "e\u0301", "👩🏽‍⚖️", "مرحبا"]) {
    expect(actions.some(({ updates }) => updates.some((update) => update.includes(text)))).toBe(
      true,
    );
  }
  // An intentionally failing oracle proves lifecycle payloads shrink, rather than
  // treating each sampled multi-update action as an opaque constant.
  const failure = fc.check(
    fc.property(browserImeActionArbitrary, () => false),
    propertyConfig({ seed: 431 }),
  );
  expect(failure.failed).toBe(true);
  const minimal = failure.counterexample?.at(0);
  expect(minimal?.updates).toHaveLength(1);
  expect(minimal?.updates.at(0)).toBe("alpha");
  expect(minimal?.completion).toBe("commit");
});

test("browser trace sequences shrink to one action", () => {
  const failure = fc.check(
    fc.property(browserInputTraceArbitrary(12), () => false),
    propertyConfig({ seed: 197 }),
  );
  expect(failure.failed).toBe(true);
  expect(failure.counterexample?.at(0)?.actions).toHaveLength(1);
});

test("browser lanes retain pinned seeds and reject malformed overrides", () => {
  expect(parseBrowserInputTraceConfig({}, "nightly")).toEqual({
    seeds: BROWSER_TRACE_FIXED_SEEDS.nightly,
    runs: 20,
  });
  expect(parseBrowserInputTraceConfig({}, "pullRequest")).toEqual({
    seeds: BROWSER_TRACE_FIXED_SEEDS.pullRequest,
    runs: 2,
  });
  expect(
    parseBrowserInputTraceConfig({ FOLIO_FUZZ_SEEDS: "11,-29", FOLIO_FUZZ_RUNS: "7" }, "nightly"),
  ).toEqual({ seeds: [11, -29], runs: 7 });
  for (const seeds of ["", " ", "11,", ",29", "11,,29", "NaN", "1.5"]) {
    expect(() => parseBrowserInputTraceConfig({ FOLIO_FUZZ_SEEDS: seeds }, "nightly")).toThrow();
  }
  for (const runs of ["", "0", "-1", "1.5", "NaN"]) {
    expect(() => parseBrowserInputTraceConfig({ FOLIO_FUZZ_RUNS: runs }, "nightly")).toThrow();
  }
});
