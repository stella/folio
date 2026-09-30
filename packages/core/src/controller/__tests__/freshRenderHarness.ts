/**
 * Fresh-render equivalence harness.
 *
 * The painted page must be a pure function of the editor's inputs: the document
 * state, the layout inputs (page geometry, markup view, the fonts actually
 * available) and nothing else. A rig drives the real layout scheduler and
 * layout pipeline the way the React and Vue adapters wire them, through a
 * seeded sequence of events (edits, document loads, frames paused and resumed,
 * page geometry, zoom, re-renders), and two oracles judge the result:
 *
 * - every pass lays out the state the editor holds when the pass runs
 *   ({@link FreshRenderRig.staleCommits});
 * - once settled, the committed layout equals a from-scratch layout of the same
 *   inputs, compared as painted line boxes: each line's text and width, each
 *   fragment's position ({@link projectLayout}).
 *
 * Event kinds are pluggable ({@link defineFreshRenderEventKind}), and a rig
 * carries extension state (`ext`) with extra pipeline deps and layout-input
 * signature parts, so a markup view or a font set can become a layout input
 * without changing this file.
 *
 * Headless: text is measured by the fake canvas of `withFakeTextMeasure`, and
 * the pipeline runs without a painter; run a property inside
 * `withFakeTextMeasure`. The adapters' own wiring (which event triggers which
 * pass) is modelled here, not imported, so keep it in step with
 * `PagedEditor.tsx` and `useDocxEditor.ts`.
 */

import { panic } from "better-result";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState, type Plugin, type Transaction } from "prosemirror-state";

import type { LayoutRunReason } from "../../layout-engine/layoutInstrumentation";
import { clearAllCaches } from "../../layout-engine/measure/cache";
import { resetCanvasContext } from "../../layout-engine/measure/measureContainer";
import type { FlowBlock, Layout, Measure, Run } from "../../layout-engine/types";
import { LayoutSelectionGate } from "../../paged-layout/LayoutSelectionGate";
import { schema } from "../../prosemirror/schema";
import type { Document } from "../../types/document";
import { createEmptyDocument } from "../../utils/createDocument";
import { runLayoutPipeline, type LayoutOutcome } from "../layoutPipeline";
import type { LayoutPipelineDeps } from "../layoutPipeline";
import type { LayoutRunOptions } from "../layoutRunOptions";
import { createLayoutScheduler, type SchedulerClock } from "../layoutScheduler";
import { createLayoutSession, type LayoutSession } from "../layoutSession";

// ============================================================================
// FAKE CLOCK: timers always run; frames run only while the page is visible.
// ============================================================================

/** One browser frame. */
export const FRAME_MS = 16;
/** The adapters' debounced document-change notification (a React re-render). */
export const DOCUMENT_CHANGE_NOTIFY_DELAY_MS = 250;
/** Longer than any timer the adapters arm. */
const SETTLE_TIMERS_MS = 1000;
const TRANSACTION_LAYOUT_DEBOUNCE_MS = 32;
const TRANSACTION_LAYOUT_MAX_DELAY_MS = 96;

type FakeTimer = { due: number; callback: () => void };

const createPausableClock = () => {
  let now = 0;
  let nextId = 1;
  let framesPaused = false;
  const timers = new Map<number, FakeTimer>();
  const frames = new Map<number, () => void>();

  const clock: SchedulerClock = {
    now: () => now,
    setTimer: (callback, delayMs) => {
      const id = nextId++;
      timers.set(id, { due: now + delayMs, callback });
      return id;
    },
    clearTimer: (id) => {
      timers.delete(id);
    },
    requestFrame: (callback) => {
      const id = nextId++;
      frames.set(id, callback);
      return id;
    },
    cancelFrame: (id) => {
      frames.delete(id);
    },
  };

  const runDueTimers = (): void => {
    for (;;) {
      let nextDue: [number, FakeTimer] | null = null;
      for (const entry of timers) {
        if (entry[1].due <= now && (nextDue === null || entry[1].due < nextDue[1].due)) {
          nextDue = entry;
        }
      }
      if (nextDue === null) {
        return;
      }
      timers.delete(nextDue[0]);
      nextDue[1].callback();
    }
  };

  const runFrame = (): void => {
    if (framesPaused) {
      return;
    }
    const callbacks = [...frames.values()];
    frames.clear();
    for (const callback of callbacks) {
      callback();
    }
  };

  return {
    clock,
    /** Advance time frame by frame, running due timers and (when visible) frames. */
    tick: (ms: number): void => {
      const end = now + ms;
      while (now < end) {
        now = Math.min(end, now + FRAME_MS);
        runDueTimers();
        runFrame();
      }
    },
    pauseFrames: (): void => {
      framesPaused = true;
    },
    resumeFrames: (): void => {
      framesPaused = false;
    },
    isQuiet: (): boolean => timers.size === 0 && frames.size === 0,
  };
};

// ============================================================================
// LAYOUT PROJECTION: what the page paints, as line boxes.
// ============================================================================

export type LaidOutLine = { text: string; width: number };

export type LaidOutFragment = {
  kind: string;
  blockId: string | number;
  x: number;
  y: number;
  width: number;
  lines?: LaidOutLine[];
};

export type LaidOutPage = { number: number; fragments: LaidOutFragment[] };

export type LayoutArtifactsView = {
  blocks: readonly FlowBlock[];
  measures: readonly Measure[];
  layout: Layout;
};

const runText = (run: Run): string => (run.kind === "text" ? run.text : `[${run.kind}]`);

/** The text a measured line covers, sliced out of its block's runs. */
const lineText = (
  runs: readonly Run[],
  line: { fromRun: number; fromChar: number; toRun: number; toChar: number },
): string => {
  let text = "";
  for (let index = line.fromRun; index <= line.toRun; index += 1) {
    const run = runs[index];
    if (!run) {
      text += "[missing run]";
      continue;
    }
    const full = runText(run);
    const from = index === line.fromRun ? line.fromChar : 0;
    const to = index === line.toRun ? line.toChar : full.length;
    text += full.slice(from, to);
  }
  return text;
};

/**
 * Project a layout onto what it paints: per page, each fragment's box and, for
 * a paragraph, each line's text and width. Two layouts that paint the same
 * lines in the same places project equal; a measure reused for another block's
 * text shows up as a line whose text differs.
 */
export const projectLayout = ({ blocks, measures, layout }: LayoutArtifactsView): LaidOutPage[] => {
  const blockIndex = new Map<string | number, number>();
  for (const [index, block] of blocks.entries()) {
    blockIndex.set(block.id, index);
  }
  return layout.pages.map((page) => ({
    number: page.number,
    fragments: page.fragments.map((fragment): LaidOutFragment => {
      const base = {
        kind: fragment.kind,
        blockId: fragment.blockId,
        x: fragment.x,
        y: fragment.y,
        width: fragment.width,
      };
      if (fragment.kind !== "paragraph") {
        return base;
      }
      const index = blockIndex.get(fragment.blockId);
      const block = index === undefined ? undefined : blocks[index];
      const measure = index === undefined ? undefined : measures[index];
      if (block?.kind !== "paragraph" || measure?.kind !== "paragraph") {
        return { ...base, lines: [{ text: "[unmeasured paragraph]", width: 0 }] };
      }
      return {
        ...base,
        lines: measure.lines.slice(fragment.fromLine, fragment.toLine).map((line) => ({
          text: lineText(block.runs, line),
          width: line.width,
        })),
      };
    }),
  }));
};

// ============================================================================
// RIG
// ============================================================================

/** The layout inputs every rig has. Extension inputs live in `ext`. */
export type FreshRenderInputs = {
  contentWidth: number;
  document: Document;
  /** Paint-only: a change re-renders the adapter but is not a layout input. */
  zoom: number;
};

/** How a document load reaches layout: in the same task (`flushSync`) or on a later render. */
export const LOAD_LAYOUT = {
  immediate: "immediate",
  nextRender: "next-render",
} as const;

export type LoadLayout = (typeof LOAD_LAYOUT)[keyof typeof LOAD_LAYOUT];

export type StaleCommit = {
  reason: LayoutRunReason | undefined;
  laidOut: string;
  current: string;
};

export type FreshRenderRigOptions<TExt> = {
  initialDoc: PMNode;
  /** React schedules with a leading frame; Vue without one. */
  leadingFrame: boolean;
  /** Extension state an event kind owns (a markup view, loaded font subsets). */
  ext: TExt;
  plugins?: readonly Plugin[];
  /** Pipeline deps derived from the rig on every pass (from `ext`, say). */
  extraDeps?: (rig: FreshRenderRig<TExt>) => Partial<LayoutPipelineDeps<null>>;
  /** Extra layout-input signature: a change makes a re-render lay out again. */
  extraSignature?: (rig: FreshRenderRig<TExt>) => string;
  /**
   * Oracle self-check only: re-create a known defect so a test can prove the
   * oracles catch it. `scheduled-state` lays out the state an edit scheduled,
   * as the scheduler did before #1142.
   */
  mutant?: FreshRenderMutant;
};

export const FRESH_RENDER_MUTANT = {
  scheduledState: "scheduled-state",
} as const;

export type FreshRenderMutant = (typeof FRESH_RENDER_MUTANT)[keyof typeof FRESH_RENDER_MUTANT];

export type FreshRenderRig<TExt> = {
  readonly state: EditorState;
  readonly inputs: Readonly<FreshRenderInputs>;
  ext: TExt;
  /** Dispatch a transaction the way the adapters' `handleTransaction` does. */
  edit: (build: (state: EditorState) => Transaction | null) => void;
  /** `ref.loadDocument`: replace the state wholesale; lay it out now or on the next render. */
  loadDocument: (doc: PMNode, options: { layout: LoadLayout; document?: Document }) => void;
  /** Change layout inputs (a prop change); the adapter re-renders at once. */
  setInputs: (patch: Partial<FreshRenderInputs>) => void;
  /**
   * Change extension state that is a layout input (a markup view). The adapter
   * re-renders at once, or on a later render: a pass already scheduled then
   * runs with the new input before the layout-input effect does.
   */
  updateExt: (update: (ext: TExt) => TExt, layout?: LoadLayout) => void;
  /** The adapters' layout-input effect: lay out again when inputs or the document changed. */
  rerender: () => void;
  /**
   * The adapters' font-load follow-up (`watchLayoutFontLoads`'s relayout): lay
   * out the current state. The pipeline itself drops measurements taken in
   * another font set.
   */
  fontsChanged: () => void;
  /** The font set the committed layout was measured in (`watchLayoutFontLoads`). */
  readonly measuredFontSet: string | null;
  pauseFrames: () => void;
  resumeFrames: () => void;
  tick: (ms: number) => void;
  /**
   * Flush a pending render, let timers run out (frames stay as they are), then
   * show the page and run timers and frames until quiet.
   */
  settle: () => void;
  /** Every pass that laid out a state other than the editor's current one. */
  readonly staleCommits: readonly StaleCommit[];
  /** The committed layout, as painted. */
  committed: () => LaidOutPage[] | null;
  /** The committed blocks, measures and layout, for oracles that read positions. */
  committedArtifacts: () => LayoutArtifactsView | null;
  /**
   * A from-scratch layout of the current inputs, with cold caches; or of
   * another state and deps, for an oracle that materialises what the inputs
   * should read as (a review view's resolved document, say).
   */
  fresh: (override?: FreshLayoutOverride) => LaidOutPage[];
};

export type FreshLayoutOverride = {
  state?: EditorState;
  deps?: Partial<LayoutPipelineDeps<null>>;
};

const PAGE_SIZE = { w: 816, h: 1056 };

const docText = (doc: PMNode): string => {
  const paragraphs: string[] = [];
  doc.forEach((node) => {
    paragraphs.push(node.textContent);
  });
  return paragraphs.join(" | ");
};

export const createFreshRenderRig = <TExt>(
  options: FreshRenderRigOptions<TExt>,
): FreshRenderRig<TExt> => {
  const time = createPausableClock();
  const plugins = options.plugins ?? [];
  const createState = (doc: PMNode): EditorState =>
    EditorState.create({ doc, plugins: [...plugins] });

  let state = createState(options.initialDoc);
  let inputs: FreshRenderInputs = {
    contentWidth: 360,
    document: createEmptyDocument(),
    zoom: 1,
  };
  const session = createLayoutSession();
  let committed: LayoutArtifactsView | null = null;
  let pendingRender = false;
  let documentRevision = 0;
  // The layout inputs the committed layout was laid out with. A pass reads the
  // inputs current when it runs, so one the scheduler runs before a pending
  // re-render lays out inputs that render has not seen yet.
  let laidOutSignature: string | null = null;
  const staleCommits: StaleCommit[] = [];
  const syncCoordinator = new LayoutSelectionGate();

  type BuildDepsOptions = {
    layoutSession: LayoutSession;
    previousLayout: Layout | null;
  };

  const buildDeps = ({
    layoutSession,
    previousLayout,
  }: BuildDepsOptions): LayoutPipelineDeps<null> => {
    const margins = {
      top: 72,
      bottom: 72,
      left: (PAGE_SIZE.w - inputs.contentWidth) / 2,
      right: (PAGE_SIZE.w - inputs.contentWidth) / 2,
      header: 36,
      footer: 36,
    };
    return {
      contentWidth: inputs.contentWidth,
      columns: undefined,
      pageSize: PAGE_SIZE,
      margins,
      pageGap: 24,
      showMarginGuides: false,
      marginGuideColor: undefined,
      syncCoordinator,
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
      document: inputs.document,
      defaultTabStop: undefined,
      mirrorMargins: false,
      styles: null,
      layout: previousLayout,
      hfPMs: null,
      painter: null,
      pagesContainer: null,
      session: layoutSession,
      renderHfFromContentOrPm: () => undefined,
      renderHeaderFooterContentByRId: () => undefined,
      // No font set to load: a rig whose `ext` models one supplies it here.
      readFontSetSignature: () => "none",
      buildFootnoteRenderItems: () => new Map(),
      describeInvalidHighlightMarks: () => "",
      emptyTemplatePreviewEntries: [],
      emptyTemplatePreviewHidden: [],
      hyphenationReadiness: { track: () => undefined, cancel: () => undefined },
      markupView: "all-markup",
      ...options.extraDeps?.(rig),
    };
  };

  const apply = (outcome: LayoutOutcome, passSignature: string): void => {
    // The adapters keep the previous layout when a pass produced none.
    if (outcome.layout && outcome.blocks && outcome.measures) {
      committed = { blocks: outcome.blocks, measures: outcome.measures, layout: outcome.layout };
      laidOutSignature = passSignature;
    }
  };

  // Every pass goes through here, whoever asked for it: the product's contract
  // is that the laid-out state is the editor's state at that moment.
  const runPass = (laidOut: EditorState, runOptions: LayoutRunOptions): void => {
    if (laidOut !== state) {
      staleCommits.push({
        reason: runOptions.reason,
        laidOut: docText(laidOut.doc),
        current: docText(state.doc),
      });
    }
    const passSignature = signature();
    apply(
      runLayoutPipeline(
        buildDeps({ layoutSession: session, previousLayout: committed?.layout ?? null }),
        laidOut,
        runOptions,
      ),
      passSignature,
    );
  };

  // What the scheduler reads: the current state, or under the mutant the
  // state the last scheduling edit produced.
  let scheduledState = state;
  const scheduler = createLayoutScheduler({
    readState: () =>
      options.mutant === FRESH_RENDER_MUTANT.scheduledState ? scheduledState : state,
    runLayout: runPass,
    debounceMs: TRANSACTION_LAYOUT_DEBOUNCE_MS,
    leadingFrame: options.leadingFrame,
    maxDelayMs: TRANSACTION_LAYOUT_MAX_DELAY_MS,
    clock: time.clock,
  });

  let notifyTimer: number | null = null;

  const signature = (): string =>
    `${inputs.contentWidth}|${documentRevision}|${options.extraSignature?.(rig) ?? ""}`;

  const rerender = (): void => {
    pendingRender = false;
    // Against what the committed layout was laid out with, as the adapters'
    // layout-input effect compares: not against the last render's inputs.
    if (laidOutSignature === signature() && state.doc === session.lastPmDoc) {
      return;
    }
    runPass(state, { reason: "layout-input" });
  };

  const rig: FreshRenderRig<TExt> = {
    get state() {
      return state;
    },
    get inputs() {
      return inputs;
    },
    ext: options.ext,
    get staleCommits() {
      return staleCommits;
    },
    edit: (build) => {
      const tr = build(state);
      if (!tr) {
        return;
      }
      state = state.apply(tr);
      if (!tr.docChanged) {
        return;
      }
      scheduledState = state;
      scheduler.schedule();
      // The debounced document-change notification re-renders the adapter.
      if (notifyTimer !== null) {
        time.clock.clearTimer(notifyTimer);
      }
      notifyTimer = time.clock.setTimer(() => {
        notifyTimer = null;
        rerender();
      }, DOCUMENT_CHANGE_NOTIFY_DELAY_MS);
    },
    loadDocument: (doc, { layout, document }) => {
      state = createState(doc);
      inputs = { ...inputs, document: document ?? createEmptyDocument() };
      documentRevision += 1;
      pendingRender = true;
      if (layout === LOAD_LAYOUT.immediate) {
        rerender();
      }
    },
    setInputs: (patch) => {
      if (patch.document && patch.document !== inputs.document) {
        documentRevision += 1;
      }
      inputs = { ...inputs, ...patch };
      rerender();
    },
    updateExt: (update, layout = LOAD_LAYOUT.immediate) => {
      rig.ext = update(rig.ext);
      pendingRender = true;
      if (layout === LOAD_LAYOUT.immediate) {
        rerender();
      }
    },
    rerender,
    fontsChanged: () => {
      runPass(state, { reason: "font-ready" });
    },
    get measuredFontSet() {
      return session.lastMeasureInputs?.fontSet ?? null;
    },
    pauseFrames: time.pauseFrames,
    resumeFrames: time.resumeFrames,
    tick: time.tick,
    settle: () => {
      if (pendingRender) {
        rerender();
      }
      // A hidden page keeps running its timers; let them run out before the
      // page is shown, as when a user comes back to a background tab.
      time.tick(SETTLE_TIMERS_MS);
      time.resumeFrames();
      for (let guard = 0; guard < 1000 && !time.isQuiet(); guard += 1) {
        time.tick(FRAME_MS);
      }
    },
    committed: () => (committed ? projectLayout(committed) : null),
    committedArtifacts: () => committed,
    fresh: (override = {}) => {
      resetCanvasContext();
      clearAllCaches();
      const deps = {
        ...buildDeps({ layoutSession: createLayoutSession(), previousLayout: null }),
        ...override.deps,
      };
      const outcome = runLayoutPipeline(deps, override.state ?? state, { reason: "initial" });
      if (!outcome.layout || !outcome.blocks || !outcome.measures) {
        return panic("fresh layout produced no layout");
      }
      return projectLayout({
        blocks: outcome.blocks,
        measures: outcome.measures,
        layout: outcome.layout,
      });
    },
  };

  // The initial layout (the adapters' `handleEditorViewReady`).
  runPass(state, { reason: "initial" });
  return rig;
};

// ============================================================================
// EVENTS
// ============================================================================

/** An event bound to its kind, ready to apply; prints as `kind(event)`. */
export type FreshRenderEvent<TExt> = {
  kind: string;
  apply: (rig: FreshRenderRig<TExt>) => void;
};

export type FreshRenderEventKind<TEvent, TExt> = {
  kind: string;
  arbitrary: fc.Arbitrary<TEvent>;
  apply: (rig: FreshRenderRig<TExt>, event: TEvent) => void;
};

/** Bind an event kind's payloads to its `apply`, so kinds of any payload mix in one sequence. */
export const defineFreshRenderEventKind = <TEvent, TExt>(
  definition: FreshRenderEventKind<TEvent, TExt>,
): fc.Arbitrary<FreshRenderEvent<TExt>> =>
  definition.arbitrary.map((event) => ({
    kind: definition.kind,
    apply: (rig: FreshRenderRig<TExt>) => definition.apply(rig, event),
    [fc.toStringMethod]: () => `${definition.kind}(${fc.stringify(event)})`,
  }));

const WORDS = [
  "the",
  "Buyer",
  "shall",
  "pay",
  "within",
  "thirty",
  "days",
  "of",
  "a",
  "valid",
  "invoice",
];

/** Paragraph text of 0–40 words, so a paragraph spans zero to several lines. */
export const paragraphTextArbitrary = fc
  .array(fc.constantFrom(...WORDS), { maxLength: 40 })
  .map((words) => words.join(" "));

export const docFromParagraphs = (paragraphs: readonly string[]): PMNode =>
  schema.node(
    "doc",
    null,
    paragraphs.map((text) =>
      schema.node("paragraph", null, text.length > 0 ? [schema.text(text)] : []),
    ),
  );

export const initialParagraphsArbitrary = fc.array(paragraphTextArbitrary, {
  minLength: 1,
  maxLength: 6,
});

/** The first position inside the `index`-th top-level paragraph (wrapping), and its end. */
const paragraphBounds = (doc: PMNode, index: number): { start: number; end: number } => {
  const target = index % doc.childCount;
  let pos = 0;
  for (let child = 0; child < target; child += 1) {
    pos += doc.child(child).nodeSize;
  }
  const node = doc.child(target);
  return { start: pos + 1, end: pos + node.nodeSize - 1 };
};

type EditEvent =
  | { type: "insert"; paragraph: number; offset: number; text: string }
  | { type: "delete"; paragraph: number; offset: number; length: number };

/** Type or delete inside one paragraph: the block count stays, as in the incremental path. */
export const editEventKind = <TExt>() =>
  defineFreshRenderEventKind<EditEvent, TExt>({
    kind: "edit",
    arbitrary: fc.oneof(
      fc.record({
        type: fc.constant("insert" as const),
        paragraph: fc.nat(20),
        offset: fc.nat(300),
        text: paragraphTextArbitrary.map((text) => ` ${text} `),
      }),
      fc.record({
        type: fc.constant("delete" as const),
        paragraph: fc.nat(20),
        offset: fc.nat(300),
        length: fc.integer({ min: 1, max: 80 }),
      }),
    ),
    apply: (rig, event) =>
      rig.edit((state) => {
        const { start, end } = paragraphBounds(state.doc, event.paragraph);
        const at = start + (event.offset % (end - start + 1));
        switch (event.type) {
          case "insert":
            return state.tr.insertText(event.text, at);
          case "delete": {
            const to = Math.min(end, at + event.length);
            return to > at ? state.tr.delete(at, to) : null;
          }
          default: {
            const unreachable: never = event;
            return unreachable;
          }
        }
      }),
  });

type LoadEvent = {
  layout: LoadLayout;
  /** Rewrite some paragraphs of the current document, keeping the paragraph count. */
  rewrite: { paragraph: number; text: string }[];
  /** Or load a different document altogether. */
  replacement: string[] | null;
};

/** `ref.loadDocument`: often the next version of the open document (#1142's host). */
export const loadDocumentEventKind = <TExt>() =>
  defineFreshRenderEventKind<LoadEvent, TExt>({
    kind: "loadDocument",
    arbitrary: fc.record({
      layout: fc.constantFrom(LOAD_LAYOUT.immediate, LOAD_LAYOUT.nextRender),
      rewrite: fc.array(fc.record({ paragraph: fc.nat(20), text: paragraphTextArbitrary }), {
        minLength: 1,
        maxLength: 3,
      }),
      replacement: fc.option(initialParagraphsArbitrary, { nil: null, freq: 4 }),
    }),
    apply: (rig, event) => {
      const paragraphs: string[] = [];
      rig.state.doc.forEach((node) => {
        paragraphs.push(node.textContent);
      });
      for (const { paragraph, text } of event.rewrite) {
        paragraphs[paragraph % paragraphs.length] = text;
      }
      rig.loadDocument(docFromParagraphs(event.replacement ?? paragraphs), {
        layout: event.layout,
      });
    },
  });

export const tickEventKind = <TExt>() =>
  defineFreshRenderEventKind<number, TExt>({
    kind: "tick",
    arbitrary: fc.constantFrom(
      0,
      FRAME_MS,
      2 * FRAME_MS,
      100,
      DOCUMENT_CHANGE_NOTIFY_DELAY_MS + 50,
    ),
    apply: (rig, ms) => rig.tick(ms),
  });

/** The page is hidden (frames stop) or shown again. */
export const visibilityEventKind = <TExt>() =>
  defineFreshRenderEventKind<boolean, TExt>({
    kind: "visible",
    arbitrary: fc.boolean(),
    apply: (rig, visible) => (visible ? rig.resumeFrames() : rig.pauseFrames()),
  });

export const pageGeometryEventKind = <TExt>() =>
  defineFreshRenderEventKind<number, TExt>({
    kind: "contentWidth",
    arbitrary: fc.constantFrom(240, 360, 480),
    apply: (rig, contentWidth) => rig.setInputs({ contentWidth }),
  });

export const zoomEventKind = <TExt>() =>
  defineFreshRenderEventKind<number, TExt>({
    kind: "zoom",
    arbitrary: fc.constantFrom(0.5, 1, 1.5),
    apply: (rig, zoom) => rig.setInputs({ zoom }),
  });

export const rerenderEventKind = <TExt>() =>
  defineFreshRenderEventKind<null, TExt>({
    kind: "rerender",
    arbitrary: fc.constant(null),
    apply: (rig) => rig.rerender(),
  });

/** The event kinds every rig understands; an issue adds its own alongside. */
export const baseFreshRenderEventKinds = <TExt>(): fc.Arbitrary<FreshRenderEvent<TExt>>[] => [
  editEventKind<TExt>(),
  loadDocumentEventKind<TExt>(),
  tickEventKind<TExt>(),
  visibilityEventKind<TExt>(),
  pageGeometryEventKind<TExt>(),
  zoomEventKind<TExt>(),
  rerenderEventKind<TExt>(),
];

export type FreshRenderScenario<TExt> = {
  initialParagraphs: string[];
  leadingFrame: boolean;
  events: FreshRenderEvent<TExt>[];
};

export const freshRenderScenarioArbitrary = <TExt>(
  kinds: readonly fc.Arbitrary<FreshRenderEvent<TExt>>[],
  maxEvents = 30,
): fc.Arbitrary<FreshRenderScenario<TExt>> =>
  fc.record({
    initialParagraphs: initialParagraphsArbitrary,
    leadingFrame: fc.boolean(),
    events: fc.array(fc.oneof(...kinds), { maxLength: maxEvents }),
  });

export type FreshRenderVerdict = {
  staleCommits: readonly StaleCommit[];
  committed: LaidOutPage[] | null;
  fresh: LaidOutPage[];
};

type RunFreshRenderScenarioOptions<TExt> = Omit<
  FreshRenderRigOptions<TExt>,
  "initialDoc" | "leadingFrame"
> & {
  scenario: FreshRenderScenario<TExt>;
};

/** Play a scenario, settle, and return what both oracles compare. */
export const runFreshRenderScenario = <TExt>({
  scenario,
  ...rigOptions
}: RunFreshRenderScenarioOptions<TExt>): FreshRenderVerdict => {
  const rig = createFreshRenderRig({
    ...rigOptions,
    initialDoc: docFromParagraphs(scenario.initialParagraphs),
    leadingFrame: scenario.leadingFrame,
  });
  for (const event of scenario.events) {
    event.apply(rig);
  }
  rig.settle();
  return { staleCommits: rig.staleCommits, committed: rig.committed(), fresh: rig.fresh() };
};
