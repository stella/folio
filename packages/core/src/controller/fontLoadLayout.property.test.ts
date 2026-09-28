/**
 * Font oracle: the committed layout is the layout of the font set that has
 * settled, whatever order the faces loaded in.
 *
 * Bundled faces ship as `unicode-range` subsets (fontsource: `latin`,
 * `latin-ext`, `greek`, `cyrillic`, ...), fetched only for text that needs
 * them. A document in Czech, Polish, Greek or Cyrillic therefore depends on
 * subsets that can land before the first layout, after it, between an edit and
 * its layout pass, or with their `loadingdone` event delivered after a layout
 * that already used them. For every such interleaving, once every load has
 * settled the committed layout must equal a layout computed from scratch with
 * every subset loaded before anything measured. And when the initial wait
 * completes, the first layout must already be that layout, with no font-ready
 * pass after it: host UI faces and loads the layout already saw relay nothing.
 *
 * Headless: `ScriptedFontSet` stands in for `document.fonts` and its canvas,
 * and drives the same core entry points both adapters wire up
 * (`waitForInitialLayoutFonts`, `runLayoutPipeline`, `watchLayoutFontLoads`).
 */

import { afterEach, beforeEach, describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";
import { clearAllCaches } from "../layout-engine/measure/cache";
import { resetCanvasContext } from "../layout-engine/measure/measureContainer";
import { buildFontString } from "../layout-engine/measure/measureHelpers";
import { LayoutSelectionGate } from "../paged-layout/LayoutSelectionGate";
import { schema } from "../prosemirror/schema";
import {
  HOST_UI_FAMILY,
  SCRIPT_TEXT,
  SCRIPTS,
  type Script,
  ScriptedFontSet,
  scriptedDocument,
} from "./__tests__/scriptedFontSet";
import {
  readFontSetSignature,
  waitForInitialLayoutFonts,
  watchLayoutFontLoads,
} from "./fontReadiness";
import { type LayoutOutcome, type LayoutPipelineDeps, runLayoutPipeline } from "./layoutPipeline";
import { LAYOUT_MEASURE, type LayoutRunOptions } from "./layoutRunOptions";
import { createLayoutSession, type LayoutSession } from "./layoutSession";

setDefaultTimeout(propertyTestTimeout(60_000));

// Documents name proprietary faces; folio's stacks resolve them to the bundled
// substitutes (Calibri → Carlito, Aptos → Lato, the default → Tinos).
const DOCUMENT_FONTS = ["Calibri", "Aptos", null] as const;
type DocumentFont = (typeof DOCUMENT_FONTS)[number];

type ParagraphSpec = { font: DocumentFont; bold: boolean; scripts: Script[] };

const paragraphSpec: fc.Arbitrary<ParagraphSpec> = fc.record({
  font: fc.constantFrom(...DOCUMENT_FONTS),
  bold: fc.boolean(),
  // Two or more scripts make a mixed-script line.
  scripts: fc.array(fc.constantFrom(...SCRIPTS), { minLength: 1, maxLength: 3 }),
});

type Step =
  | { kind: "complete"; pick: number }
  | { kind: "complete-all" }
  | { kind: "deliver" }
  | { kind: "paint" }
  | { kind: "host-ui"; pick: number }
  | { kind: "edit"; pick: number; script: Script };

const step: fc.Arbitrary<Step> = fc.oneof(
  fc.record({ kind: fc.constant("complete" as const), pick: fc.nat() }),
  // Finishing a whole batch at once queues its event, so an edit can land
  // between a batch's loads and its `loadingdone`.
  fc.record({ kind: fc.constant("complete-all" as const) }),
  fc.record({ kind: fc.constant("deliver" as const) }),
  fc.record({ kind: fc.constant("paint" as const) }),
  fc.record({ kind: fc.constant("host-ui" as const), pick: fc.nat() }),
  fc.record({
    kind: fc.constant("edit" as const),
    pick: fc.nat(),
    script: fc.constantFrom(...SCRIPTS),
  }),
);

const FONT_SIZE_PT = 11;

const paragraphNode = ({ font, bold, scripts }: ParagraphSpec): PMNode => {
  const marks = [
    ...(font === null ? [] : [schema.mark("fontFamily", { ascii: font, hAnsi: font })]),
    ...(bold ? [schema.mark("bold")] : []),
  ];
  return schema.node("paragraph", null, [
    schema.text(scripts.map((script) => SCRIPT_TEXT[script]).join(" "), marks),
  ]);
};

/** The fonts the painter sets on each paragraph's text, and that text. */
const paintedRuns = (doc: PMNode): { font: string; text: string }[] => {
  const runs: { font: string; text: string }[] = [];
  doc.descendants((node) => {
    if (!node.isText) {
      return true;
    }
    const family = node.marks.find((mark) => mark.type.name === "fontFamily")?.attrs["ascii"];
    const bold = node.marks.some((mark) => mark.type.name === "bold");
    runs.push({
      font: buildFontString({
        ...(typeof family === "string" ? { fontFamily: family } : {}),
        bold,
        fontSize: FONT_SIZE_PT,
      }),
      text: node.text ?? "",
    });
    return false;
  });
  return runs;
};

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS = { top: 72, right: 72, bottom: 72, left: 72, header: 36, footer: 36 };

const pipelineDeps = (
  session: LayoutSession,
  fontSet: ScriptedFontSet,
  layout: LayoutOutcome["layout"] | null,
): LayoutPipelineDeps<null> => ({
  contentWidth: PAGE_SIZE.w - MARGINS.left - MARGINS.right,
  columns: undefined,
  pageSize: PAGE_SIZE,
  margins: MARGINS,
  pageGap: 24,
  showMarginGuides: false,
  marginGuideColor: undefined,
  syncCoordinator: new LayoutSelectionGate(),
  headerContent: null,
  footerContent: null,
  firstPageHeaderContent: null,
  firstPageFooterContent: null,
  headerContentRId: null,
  footerContentRId: null,
  firstPageHeaderContentRId: null,
  firstPageFooterContentRId: null,
  sectionHeaderFooterRefs: undefined,
  theme: undefined,
  sectionProperties: null,
  document: null,
  defaultTabStop: undefined,
  mirrorMargins: false,
  styles: null,
  layout: layout ?? null,
  hfPMs: null,
  painter: null,
  pagesContainer: null,
  session,
  renderHfFromContentOrPm: () => undefined,
  renderHeaderFooterContentByRId: () => undefined,
  readFontSetSignature: () => readFontSetSignature(fontSet),
  buildFootnoteRenderItems: () => new Map(),
  describeInvalidHighlightMarks: () => "",
  emptyTemplatePreviewEntries: [],
  emptyTemplatePreviewHidden: [],
  hyphenationReadiness: { track: () => undefined, cancel: () => undefined },
});

/** What a reader sees: every line's extent and width, and the page breaks. */
const layoutProjection = (outcome: LayoutOutcome): string =>
  JSON.stringify({ measures: outcome.measures, pages: outcome.layout?.pages });

const useDocument = (fontSet: ScriptedFontSet): void => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    writable: true,
    value: scriptedDocument(fontSet),
  });
  resetCanvasContext();
};

/** The layout of `state` with every subset loaded before anything measured. */
const referenceLayout = (state: EditorState): string => {
  const fontSet = new ScriptedFontSet();
  fontSet.loadAll();
  useDocument(fontSet);
  // From scratch: nothing the editor under test cached may reach the oracle.
  clearAllCaches();
  const outcome = runLayoutPipeline(pipelineDeps(createLayoutSession(), fontSet, null), state, {
    reason: "initial",
  });
  return layoutProjection(outcome);
};

type EditorRun = {
  initial: string;
  initialState: EditorState;
  committed: string;
  state: EditorState;
  fontReadyPasses: number;
};

const GATE_MODES = ["awaited", "timed-out"] as const;
type GateMode = (typeof GATE_MODES)[number];

/**
 * One editor's life against a scripted font set: the initial wait (awaited,
 * or timed out with its loads still in flight), the first layout, then the
 * scripted steps, then every outstanding load and event settled.
 */
const runEditor = async (
  paragraphs: readonly ParagraphSpec[],
  gate: GateMode,
  steps: readonly Step[],
): Promise<EditorRun> => {
  const fontSet = new ScriptedFontSet();
  useDocument(fontSet);
  const session = createLayoutSession();
  let state = EditorState.create({ doc: schema.node("doc", null, paragraphs.map(paragraphNode)) });
  let committed: LayoutOutcome = {};
  let pendingInitial = true;
  let fontReadyPasses = 0;

  const layout = (options: LayoutRunOptions) => {
    const outcome = runLayoutPipeline(
      pipelineDeps(session, fontSet, committed.layout),
      state,
      options,
    );
    if (outcome.layout) {
      committed = outcome;
    }
  };

  const unwatch = watchLayoutFontLoads({
    fontSet,
    measuredFontSet: () => session.lastMeasureInputs?.fontSet ?? null,
    relayout: () => {
      if (pendingInitial) {
        return;
      }
      fontReadyPasses += 1;
      layout({ reason: "font-ready" });
    },
  });

  const completeOne = (pick: number) => {
    const pending = fontSet.pending();
    const face = pending.at(pick % Math.max(1, pending.length));
    if (face) {
      fontSet.complete(face);
    }
  };

  const fontsSettled = waitForInitialLayoutFonts(null, state.doc, fontSet);
  if (gate === "awaited") {
    let pick = steps.length;
    while (fontSet.pending().length > 0) {
      completeOne(pick);
      pick = pick * 7 + 3;
      await Promise.resolve();
    }
    expect(await fontsSettled).toBe(true);
  }
  pendingInitial = false;
  layout({ reason: "initial" });
  const initial = layoutProjection(committed);
  const initialState = state;

  for (const next of steps) {
    switch (next.kind) {
      case "complete":
        completeOne(next.pick);
        break;
      case "complete-all":
        while (fontSet.pending().length > 0) {
          completeOne(0);
        }
        break;
      case "deliver":
        fontSet.deliverEvent();
        break;
      case "paint":
        for (const run of paintedRuns(state.doc)) {
          fontSet.requestForText(run.font, run.text);
        }
        break;
      case "host-ui": {
        const hostFaces = fontSet.faces.filter((face) => face.family === HOST_UI_FAMILY);
        const face = hostFaces.at(next.pick % hostFaces.length);
        if (face) {
          void fontSet.request(face);
        }
        break;
      }
      case "edit": {
        // Append to one paragraph, laid out the way the scheduler runs a
        // transaction: incrementally.
        const target = next.pick % state.doc.childCount;
        let paragraphStart = 0;
        for (let index = 0; index < target; index += 1) {
          paragraphStart += state.doc.child(index).nodeSize;
        }
        const paragraphEnd = paragraphStart + state.doc.child(target).nodeSize - 1;
        state = state.apply(state.tr.insertText(` ${SCRIPT_TEXT[next.script]}`, paragraphEnd));
        layout({ reason: "transaction", measure: LAYOUT_MEASURE.incremental });
        break;
      }
      default:
        next satisfies never;
    }
    await Promise.resolve();
  }

  // Settle: paint the final text (the browser fetches what it needs), finish
  // every load and deliver every event.
  for (const run of paintedRuns(state.doc)) {
    fontSet.requestForText(run.font, run.text);
  }
  while (fontSet.pending().length > 0 || fontSet.queuedEvents.length > 0) {
    completeOne(0);
    fontSet.deliverEvent();
    await Promise.resolve();
  }
  await fontsSettled;
  unwatch();

  return {
    initial,
    initialState,
    committed: layoutProjection(committed),
    state,
    fontReadyPasses,
  };
};

const originalDocument = Object.getOwnPropertyDescriptor(globalThis, "document");

describe("font oracle: layout after the subsets settle equals the preloaded layout", () => {
  beforeEach(() => {
    clearAllCaches();
  });

  afterEach(() => {
    if (originalDocument) {
      Object.defineProperty(globalThis, "document", originalDocument);
    } else {
      Reflect.deleteProperty(globalThis, "document");
    }
    resetCanvasContext();
    clearAllCaches();
  });

  test("any load order, event timing and edit interleaving settles on the preloaded layout", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(paragraphSpec, { minLength: 1, maxLength: 4 }),
        fc.constantFrom(...GATE_MODES),
        fc.array(step, { maxLength: 24 }),
        async (paragraphs, gate, steps) => {
          const run = await runEditor(paragraphs, gate, steps);
          expect(run.committed).toBe(referenceLayout(run.state));
        },
      ),
      propertyConfig({ numRuns: 60 }),
    );
  });

  test("an awaited initial wait lays out once, in the settled fonts", async () => {
    await fc.assert(
      fc.asyncProperty(
        fc.array(paragraphSpec, { minLength: 1, maxLength: 4 }),
        fc.array(
          step.filter((candidate) => candidate.kind !== "edit"),
          { maxLength: 24 },
        ),
        async (paragraphs, steps) => {
          const run = await runEditor(paragraphs, "awaited", steps);
          expect(run.initial).toBe(referenceLayout(run.initialState));
          expect(run.fontReadyPasses).toBe(0);
        },
      ),
      propertyConfig({ numRuns: 40 }),
    );
  });
});
