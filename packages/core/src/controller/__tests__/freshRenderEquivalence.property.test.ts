/**
 * Fresh-render equivalence: after any sequence of edits, document loads,
 * paused and resumed frames, geometry and zoom changes and re-renders, no pass
 * lays out a state the editor no longer holds, and the settled layout paints
 * the same lines as a from-scratch layout of the same inputs.
 *
 * The mutant checks prove each oracle catches #1142 (a pass laying out the
 * state an edit scheduled after `loadDocument` replaced it) on generated
 * sequences, so neither oracle is vacuous.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import {
  baseFreshRenderEventKinds,
  FRESH_RENDER_MUTANT,
  freshRenderScenarioArbitrary,
  runFreshRenderScenario,
  type FreshRenderVerdict,
} from "./freshRenderHarness";

setDefaultTimeout(propertyTestTimeout(120_000));

const scenarios = freshRenderScenarioArbitrary(baseFreshRenderEventKinds<null>());

const noStaleCommit = (verdict: FreshRenderVerdict): boolean => verdict.staleCommits.length === 0;

const committedIsFresh = (verdict: FreshRenderVerdict): boolean =>
  JSON.stringify(verdict.committed) === JSON.stringify(verdict.fresh);

describe("fresh-render equivalence", () => {
  test("no pass lays out a replaced state, and the settled layout equals a fresh one", () => {
    withFakeTextMeasure(() => {
      fc.assert(
        fc.property(scenarios, (scenario) => {
          const verdict = runFreshRenderScenario({ scenario, ext: null });
          expect(verdict.staleCommits).toEqual([]);
          expect(verdict.committed).toEqual(verdict.fresh);
        }),
        propertyConfig({ numRuns: 300 }),
      );
    });
  });

  describe("oracle self-check against the #1142 mutant", () => {
    const mutantIsCaught = (oracle: (verdict: FreshRenderVerdict) => boolean): boolean => {
      let caught = false;
      withFakeTextMeasure(() => {
        const result = fc.check(
          fc.property(scenarios, (scenario) =>
            oracle(
              runFreshRenderScenario({
                scenario,
                ext: null,
                mutant: FRESH_RENDER_MUTANT.scheduledState,
              }),
            ),
          ),
          propertyConfig({ numRuns: 400 }),
        );
        caught = result.failed;
      });
      return caught;
    };

    test("the stale-commit oracle catches it", () => {
      expect(mutantIsCaught(noStaleCommit)).toBe(true);
    });

    test("the fresh-layout oracle catches it", () => {
      expect(mutantIsCaught(committedIsFresh)).toBe(true);
    });
  });
});
