/**
 * Font loads as fresh-render events: bundled faces arrive as `unicode-range`
 * subsets at any point among edits, document loads, paused frames, geometry
 * and zoom changes, and once the subsets the text needs have loaded and their
 * `loadingdone` is delivered, the committed layout paints the same lines as a
 * from-scratch layout in that font set: every needed subset loaded before
 * anything measured.
 *
 * The font set is `ScriptedFontSet`: text is measured per character in the
 * first face of its stack that has loaded and covers it, else in a wider
 * fallback. Font events follow the browser: every pass paints its text, which
 * requests the subsets it needs; a load completes on its own; `loadingdone`
 * arrives later, when the adapters' `watchLayoutFontLoads` decides whether to
 * lay out.
 *
 * The mutant checks re-create the two halves of #1141 (a load event that lays
 * nothing out, and measurements not bound to the font set they were taken
 * in) to prove the oracle catches each.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import { panic } from "better-result";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";

import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { buildFontString } from "../../layout-engine/measure/measureHelpers";
import { readFontSetSignature, watchLayoutFontLoads } from "../fontReadiness";
import {
  baseFreshRenderEventKinds,
  defineFreshRenderEventKind,
  type FreshRenderEvent,
  type FreshRenderRig,
  type FreshRenderScenario,
  freshRenderScenarioArbitrary,
  type FreshRenderVerdict,
  runFreshRenderScenario,
} from "./freshRenderHarness";
import {
  HOST_UI_FAMILY,
  SCRIPT_TEXT,
  SCRIPTS,
  type Script,
  ScriptedFontSet,
} from "./scriptedFontSet";

setDefaultTimeout(propertyTestTimeout(120_000));

/** Oracle self-check only: the two halves of #1141, re-created. */
const FONT_MUTANT = {
  /** `loadingdone` lays nothing out (the old suppression window swallowed it). */
  unwatched: "unwatched",
  /** Measures and caches carry no font set, so fallback widths are reused. */
  unbound: "unbound",
} as const;

type FontMutant = (typeof FONT_MUTANT)[keyof typeof FONT_MUTANT];

type FontExt = { fontSet: ScriptedFontSet; mutant: FontMutant | null };

// The fake canvas measures through the scenario's font set.
const fontSetHolder: { current: ScriptedFontSet | null } = { current: null };

const advance = (character: string, font: string): number => {
  const fontSet = fontSetHolder.current;
  if (!fontSet) {
    return panic("No scripted font set is active");
  }
  return fontSet.advance(font, character);
};

/** The font the painter sets on the rig's unformatted text. */
const PAINTED_FONT = buildFontString({ fontSize: 11 });

const paragraphEnd = (doc: PMNode, index: number): number => {
  const target = index % doc.childCount;
  let end = 0;
  for (let child = 0; child <= target; child += 1) {
    end += doc.child(child).nodeSize;
  }
  return end - 1;
};

/** Painting the document's text fetches the subsets it needs. */
const paint = (rig: FreshRenderRig<FontExt>): void => {
  rig.ext.fontSet.requestForText(PAINTED_FONT, rig.state.doc.textContent);
};

/** Deliver one queued `loadingdone`, as the adapters hear it. */
const deliverLoadingDone = (rig: FreshRenderRig<FontExt>): void => {
  const { fontSet, mutant } = rig.ext;
  const unwatch =
    mutant === FONT_MUTANT.unwatched
      ? () => undefined
      : watchLayoutFontLoads({
          fontSet,
          measuredFontSet: () => rig.measuredFontSet,
          relayout: rig.fontsChanged,
        });
  fontSet.deliverEvent();
  unwatch();
};

/** Every paragraph gains text in another script, measured before its subset loads. */
const seedScripts: FreshRenderEvent<FontExt> = {
  kind: "seedScripts",
  apply: (rig) =>
    rig.edit((state) => {
      const tr = state.tr;
      for (let paragraph = state.doc.childCount - 1; paragraph >= 0; paragraph -= 1) {
        const script = SCRIPTS[(paragraph + 1) % SCRIPTS.length] ?? "english";
        tr.insertText(` ${SCRIPT_TEXT[script]}`, paragraphEnd(state.doc, paragraph));
      }
      return tr;
    }),
};

const fontEventKinds = (): fc.Arbitrary<FreshRenderEvent<FontExt>>[] => [
  defineFreshRenderEventKind<{ script: Script; paragraph: number }, FontExt>({
    kind: "typeScript",
    arbitrary: fc.record({ script: fc.constantFrom(...SCRIPTS), paragraph: fc.nat(20) }),
    apply: (rig, { script, paragraph }) =>
      rig.edit((state) =>
        state.tr.insertText(` ${SCRIPT_TEXT[script]}`, paragraphEnd(state.doc, paragraph)),
      ),
  }),
  defineFreshRenderEventKind<number, FontExt>({
    kind: "subsetLoaded",
    arbitrary: fc.nat(),
    apply: (rig, pick) => {
      const pending = rig.ext.fontSet.pending();
      const face = pending.at(pick % Math.max(1, pending.length));
      if (face) {
        rig.ext.fontSet.complete(face);
      }
    },
  }),
  // A whole batch lands at once, so an edit can fall between the loads and
  // their `loadingdone`.
  defineFreshRenderEventKind<null, FontExt>({
    kind: "batchLoaded",
    arbitrary: fc.constant(null),
    apply: (rig) => {
      for (const face of rig.ext.fontSet.pending()) {
        rig.ext.fontSet.complete(face);
      }
    },
  }),
  defineFreshRenderEventKind<null, FontExt>({
    kind: "loadingDone",
    arbitrary: fc.constant(null),
    apply: deliverLoadingDone,
  }),
  defineFreshRenderEventKind<number, FontExt>({
    kind: "hostUiFont",
    arbitrary: fc.nat(),
    apply: (rig, pick) => {
      const hostFaces = rig.ext.fontSet.faces.filter((face) => face.family === HOST_UI_FAMILY);
      const face = hostFaces.at(pick % hostFaces.length);
      if (face) {
        void rig.ext.fontSet.request(face);
      }
    },
  }),
];

/**
 * The page shows the final text, every subset it needs loads and every
 * `loadingdone` is heard. Only the needed subsets: loading every face would
 * change the font set once more and force a full layout that hides a stale
 * incremental one.
 */
const settleFonts: FreshRenderEvent<FontExt> = {
  kind: "settleFonts",
  apply: (rig) => {
    const { fontSet } = rig.ext;
    paint(rig);
    for (const face of fontSet.pending()) {
      fontSet.complete(face);
    }
    while (fontSet.queuedEvents.length > 0) {
      deliverLoadingDone(rig);
    }
  },
};

// Font kinds listed twice: a stale measure needs a load to land between a pass
// and the next edit, so loads must be about as common as the base events.
const scenarios = freshRenderScenarioArbitrary([
  ...baseFreshRenderEventKinds<FontExt>(),
  ...fontEventKinds(),
  ...fontEventKinds(),
]).map(({ initialParagraphs, leadingFrame, events }) => ({
  initialParagraphs,
  leadingFrame,
  events: [seedScripts, ...events, settleFonts],
}));

type RunScenarioOptions = {
  scenario: FreshRenderScenario<FontExt>;
  mutant: FontMutant | null;
};

const runScenario = ({ scenario, mutant }: RunScenarioOptions): FreshRenderVerdict => {
  const fontSet = new ScriptedFontSet();
  fontSetHolder.current = fontSet;
  return runFreshRenderScenario({
    scenario,
    ext: { fontSet, mutant },
    extraDeps: (rig) => {
      // Every pass paints the text it lays out, and painting requests the
      // subsets that text needs.
      paint(rig);
      return {
        readFontSetSignature: () =>
          rig.ext.mutant === FONT_MUTANT.unbound
            ? "unbound"
            : readFontSetSignature(rig.ext.fontSet),
      };
    },
  });
};

const committedIsFresh = (verdict: FreshRenderVerdict): boolean =>
  JSON.stringify(verdict.committed) === JSON.stringify(verdict.fresh);

describe("fresh-render equivalence under font loads", () => {
  test("once the needed subsets have loaded, the settled layout equals a preloaded one", () => {
    withFakeTextMeasure(
      () => {
        fc.assert(
          fc.property(scenarios, (scenario) => {
            const verdict = runScenario({ scenario, mutant: null });
            expect(verdict.staleCommits).toEqual([]);
            expect(verdict.committed).toEqual(verdict.fresh);
          }),
          propertyConfig({ numRuns: 200 }),
        );
      },
      { charWidth: advance },
    );
  });

  describe("oracle self-check against the #1141 mutants", () => {
    const mutantIsCaught = (mutant: FontMutant): boolean => {
      let caught = false;
      withFakeTextMeasure(
        () => {
          const result = fc.check(
            fc.property(scenarios, (scenario) =>
              committedIsFresh(runScenario({ scenario, mutant })),
            ),
            propertyConfig({ numRuns: 300 }),
          );
          caught = result.failed;
        },
        { charWidth: advance },
      );
      return caught;
    };

    test("a load event that lays nothing out is caught", () => {
      expect(mutantIsCaught(FONT_MUTANT.unwatched)).toBe(true);
    });

    test("measurements not bound to their font set are caught", () => {
      expect(mutantIsCaught(FONT_MUTANT.unbound)).toBe(true);
    });
  });
});
