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
import { schema } from "../../prosemirror/schema";
import { createEmptyDocument } from "../../utils/createDocument";
import {
  baseFreshRenderEventKinds,
  FRESH_RENDER_MUTANT,
  createFreshRenderRig,
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
  test("a load changing font alternates and one paragraph matches a cold measure", () => {
    const fontFamily = schema.marks["fontFamily"]?.create({
      ascii: "Brand Face",
      hAnsi: "Brand Face",
    });
    if (!fontFamily) {
      throw new Error("Expected fontFamily mark in schema");
    }
    const paragraphs = [
      "First paragraph has enough words to wrap across several lines repeatedly",
      "Middle paragraph is edited after the replacement document loads",
      "Last paragraph also has enough words to wrap across several lines repeatedly",
    ];
    const makeDoc = () =>
      schema.node(
        "doc",
        null,
        paragraphs.map((text) => schema.node("paragraph", null, [schema.text(text, [fontFamily])])),
      );
    const makeDocument = (altName: string) => {
      const document = createEmptyDocument();
      document.package.fontTable = { fonts: [{ name: "Brand Face", altName }] };
      return document;
    };

    withFakeTextMeasure(
      () => {
        fc.assert(
          fc.property(fc.constantFrom("Cambria", "Calibri"), (newAlternate) => {
            const rig = createFreshRenderRig({
              initialDoc: makeDoc(),
              leadingFrame: false,
              ext: null,
            });
            const beforeAlternate = rig.committed();
            rig.setInputs({ document: makeDocument("Arial") });
            expect(rig.committed()).not.toEqual(beforeAlternate);
            rig.loadDocument(makeDoc(), {
              layout: "next-render",
              document: makeDocument(newAlternate),
            });
            rig.edit((state) =>
              state.tr.insertText(" revised", state.doc.child(0).nodeSize + 1 + 6),
            );
            rig.tick(100);
            expect(rig.staleCommits).toEqual([]);
            expect(rig.committed()).toEqual(rig.fresh());
          }),
          propertyConfig({ numRuns: 20, seed: 1148 }),
        );
      },
      {
        charWidth: (_char, font) => {
          if (font.includes("Cambria")) {
            return 18;
          }
          if (font.includes("Calibri")) {
            return 12;
          }
          return font.includes("Arial") ? 9 : 6;
        },
      },
    );
  });

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
          propertyConfig({ numRuns: 400, seed: 1142 }),
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
