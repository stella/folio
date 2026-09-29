/**
 * A review view lays out the text it shows.
 *
 * Built on the fresh-render harness: the markup view is a layout input the rig
 * carries (`ext`), a view switch is an event kind beside edits, loads, paused
 * frames, geometry, zoom and re-renders, and after every generated sequence
 * the settled layout is judged by three oracles:
 *
 * - no pass laid out a state the editor no longer held;
 * - the committed layout equals a fresh layout of the same state and view;
 * - it equals a fresh All Markup layout of the state the view reads,
 *   materialised on its own: every revision rejected (Original) or accepted
 *   (No Markup, Simple Markup) through the Reject All and Accept All commands.
 *
 * Equality is the painted geometry: every fragment's box and every line's text
 * and width. The documents are generated: justified and ragged paragraphs
 * holding tracked insertions, deletions, run-format changes (a size change, so
 * it moves line breaks), moves, inserted and deleted paragraph marks,
 * paragraph-property changes and a table with tracked rows; edits run in
 * editing or suggesting mode, so typing in Original with tracking on inserts
 * text the view does not show.
 *
 * Three invariants of the committed blocks keep editing in a view sane: every
 * painted character addresses the editor character it shows (so caret,
 * selection and click-to-position resolve), every editor position has a
 * painted place for its caret, and Simple Markup's change bars sit on the
 * paragraphs holding revisions.
 *
 * The mutant check proves the view oracle is not vacuous: laying every view out
 * as All Markup, the way views were a CSS filter before #1143, fails it.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";
import { EditorState } from "prosemirror-state";

import {
  assertProperty,
  propertyConfig,
  propertyTestTimeout,
} from "../../../../test/property-testing";

import { SECTION, WORD_STYLES, buildRawPackage } from "../__tests__/documentShapes";
import {
  EDITOR_MODES,
  createHarnessDoc,
  createHarnessPlugins,
  parseShapeDocument,
  resolveAllChanges,
} from "../__tests__/editorHarness";
import type { EditorMode } from "../__tests__/editorHarness";
import {
  sizeProportionalCharWidth,
  withFakeTextMeasure,
} from "../layout-engine/measure/__tests__/fakeTextMeasure";
import type { FlowBlock, Run } from "../layout-engine/types";
import { DISPLAY_MODES } from "../managers/EditorModeManager";
import type { DisplayMode } from "../managers/EditorModeManager";
import { projectMarkupView, visibleCaretPosition } from "../prosemirror/markupViewProjection";
import { schema } from "../prosemirror/schema";
import type { Document } from "../types/document";
import {
  LOAD_LAYOUT,
  createFreshRenderRig,
  defineFreshRenderEventKind,
  loadDocumentEventKind,
  pageGeometryEventKind,
  rerenderEventKind,
  tickEventKind,
  visibilityEventKind,
  zoomEventKind,
} from "./__tests__/freshRenderHarness";
import type { FreshRenderEvent, FreshRenderRig, LoadLayout } from "./__tests__/freshRenderHarness";
import type { LayoutPipelineDeps } from "./layoutPipeline";

setDefaultTimeout(propertyTestTimeout(120_000));

type PositionRange = { from: number; to: number };

// ============================================================================
// DOCUMENTS
// ============================================================================

const WORDS = [
  "the",
  "Supplier",
  "shall",
  "deliver",
  "goods",
  "within",
  "ten",
  "business",
  "days",
  "of",
  "confirmation",
  "Dodavatel",
  "zboží",
  "kupujícího",
  "notice",
  "in",
  "writing",
] as const;

const phrase = fc
  .array(fc.constantFrom(...WORDS), { minLength: 1, maxLength: 14 })
  .map((words) => `${words.join(" ")} `);

const SEGMENT_KINDS = ["plain", "insertion", "deletion", "format-change"] as const;
type Segment = { kind: (typeof SEGMENT_KINDS)[number]; text: string };

const PARAGRAPH_MARKS = ["plain", "inserted", "deleted"] as const;
type ParagraphSpec = {
  segments: Segment[];
  mark: (typeof PARAGRAPH_MARKS)[number];
  justified: boolean;
  alignmentChange: boolean;
};

type MoveSpec = { from: number; to: number; text: string };

const ROW_KINDS = ["plain", "inserted", "deleted"] as const;
/** A two-column table placed before paragraph `before`; a tracked row's cells carry its revision. */
type TableSpec = {
  before: number;
  rows: { kind: (typeof ROW_KINDS)[number]; cells: [Segment[], Segment[]] }[];
};
type DocumentSpec = { paragraphs: ParagraphSpec[]; move: MoveSpec | null; table: TableSpec | null };

const segmentArb = fc.record({ kind: fc.constantFrom(...SEGMENT_KINDS), text: phrase });

const paragraphArb = fc.record({
  segments: fc.array(segmentArb, { minLength: 1, maxLength: 6 }),
  mark: fc.oneof(
    { weight: 3, arbitrary: fc.constant("plain" as const) },
    { weight: 1, arbitrary: fc.constant("inserted" as const) },
    { weight: 1, arbitrary: fc.constant("deleted" as const) },
  ),
  justified: fc.boolean(),
  alignmentChange: fc.boolean(),
});

const cellArb = fc.array(segmentArb, { minLength: 1, maxLength: 3 });

const documentArb: fc.Arbitrary<DocumentSpec> = fc
  .array(paragraphArb, { minLength: 1, maxLength: 6 })
  .chain((paragraphs) =>
    fc.record({
      paragraphs: fc.constant(paragraphs),
      move: fc.option(
        fc.record({
          from: fc.nat({ max: paragraphs.length - 1 }),
          to: fc.nat({ max: paragraphs.length - 1 }),
          text: phrase,
        }),
        { nil: null },
      ),
      table: fc.option(
        fc.record({
          before: fc.nat({ max: paragraphs.length - 1 }),
          rows: fc.array(
            fc.record({
              kind: fc.constantFrom(...ROW_KINDS),
              cells: fc.tuple(cellArb, cellArb),
            }),
            { minLength: 1, maxLength: 4 },
          ),
        }),
        { nil: null },
      ),
    }),
  );

const FIXED_DATE = "2026-01-01T00:00:00Z";
const escapeText = (text: string): string =>
  text.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");

const buildBody = ({ paragraphs, move, table }: DocumentSpec): string => {
  let nextId = 1;
  const revision = (): string => `w:id="${nextId++}" w:author="Reviewer" w:date="${FIXED_DATE}"`;
  const run = (text: string, rPr = ""): string =>
    `<w:r>${rPr === "" ? "" : `<w:rPr>${rPr}</w:rPr>`}<w:t xml:space="preserve">${escapeText(text)}</w:t></w:r>`;
  const segmentXml = ({ kind, text }: Segment): string => {
    switch (kind) {
      case "plain":
        return run(text);
      case "insertion":
        return `<w:ins ${revision()}>${run(text)}</w:ins>`;
      case "deletion":
        return `<w:del ${revision()}><w:r><w:delText xml:space="preserve">${escapeText(text)}</w:delText></w:r></w:del>`;
      case "format-change":
        return run(text, `<w:sz w:val="32"/><w:rPrChange ${revision()}><w:rPr/></w:rPrChange>`);
      default: {
        const exhaustive: never = kind;
        return exhaustive;
      }
    }
  };
  const moveXml = (side: "From" | "To", text: string): string => {
    const range = nextId++;
    const body =
      side === "From"
        ? `<w:r><w:delText xml:space="preserve">${escapeText(text)}</w:delText></w:r>`
        : run(text);
    return (
      `<w:move${side}RangeStart w:id="${range}" w:name="move1" w:author="Reviewer" w:date="${FIXED_DATE}"/>` +
      `<w:move${side} ${revision()}>${body}</w:move${side}>` +
      `<w:move${side}RangeEnd w:id="${range}"/>`
    );
  };
  const last = paragraphs.length - 1;
  const xml: string[] = paragraphs.map((paragraph, index) => {
    // The body's final paragraph mark stays untracked, as word processors keep it.
    const mark = index === last ? "plain" : paragraph.mark;
    const markXml =
      mark === "plain"
        ? ""
        : `<w:rPr><w:${mark === "inserted" ? "ins" : "del"} ${revision()}/></w:rPr>`;
    const alignment = `<w:jc w:val="${paragraph.justified ? "both" : "left"}"/>`;
    const alignmentChange = paragraph.alignmentChange
      ? `<w:pPrChange ${revision()}><w:pPr><w:jc w:val="${paragraph.justified ? "left" : "both"}"/></w:pPr></w:pPrChange>`
      : "";
    let content = paragraph.segments.map(segmentXml).join("");
    if (move && move.from === index) content += moveXml("From", move.text);
    if (move && move.to === index) content += moveXml("To", move.text);
    return `<w:p><w:pPr>${alignment}${markXml}${alignmentChange}</w:pPr>${content}</w:p>`;
  });
  if (table) {
    const rowXml = ({ kind, cells }: TableSpec["rows"][number]): string => {
      const rowRevision =
        kind === "plain"
          ? ""
          : `<w:trPr><w:${kind === "inserted" ? "ins" : "del"} ${revision()}/></w:trPr>`;
      const cellXml = (segments: Segment[]): string => {
        // A tracked row's text carries the row's revision, as word processors write it.
        const content = segments
          .map((segment) =>
            segmentXml(
              kind === "plain"
                ? segment
                : { kind: kind === "inserted" ? "insertion" : "deletion", text: segment.text },
            ),
          )
          .join("");
        return `<w:tc><w:tcPr><w:tcW w:w="4680" w:type="dxa"/></w:tcPr><w:p>${content}</w:p></w:tc>`;
      };
      return `<w:tr>${rowRevision}${cells.map(cellXml).join("")}</w:tr>`;
    };
    const tableXml =
      '<w:tbl><w:tblPr><w:tblW w:w="9360" w:type="dxa"/></w:tblPr>' +
      '<w:tblGrid><w:gridCol w:w="4680"/><w:gridCol w:w="4680"/></w:tblGrid>' +
      `${table.rows.map(rowXml).join("")}</w:tbl>`;
    xml.splice(table.before, 0, tableXml);
  }
  return xml.join("") + SECTION();
};

const buildDocument = async (spec: DocumentSpec): Promise<Document> =>
  parseShapeDocument(await buildRawPackage({ body: buildBody(spec), styles: WORD_STYLES }));

// ============================================================================
// EVENTS
// ============================================================================

/** The rig's extension state: the review view the markup menu shows. */
type MarkupViewExt = { view: DisplayMode };

/**
 * Choosing a view in the markup menu. The adapter re-renders at once, or on a
 * later render, so a pass an edit already scheduled can run under the new view
 * before the layout-input effect does.
 */
const markupViewEventKind = defineFreshRenderEventKind<
  { view: DisplayMode; layout: LoadLayout },
  MarkupViewExt
>({
  kind: "markupView",
  arbitrary: fc.record({
    view: fc.constantFrom(...DISPLAY_MODES),
    layout: fc.constantFrom(LOAD_LAYOUT.immediate, LOAD_LAYOUT.nextRender),
  }),
  apply: (rig, { view, layout }) => rig.updateExt(() => ({ view }), layout),
});

/** Every position inside a textblock's content, in document order. */
const textPositions = (state: EditorState): number[] => {
  const positions: number[] = [];
  state.doc.descendants((node, pos) => {
    if (!node.isTextblock) return true;
    for (let offset = 0; offset <= node.content.size; offset++) positions.push(pos + 1 + offset);
    return false;
  });
  return positions;
};

type TextblockEdit =
  | { type: "insert"; at: number; text: string }
  | { type: "delete"; at: number; length: number };

/**
 * Type or delete at a textblock position anywhere in the story, table cells
 * included. The state's suggestion plugin tracks what is typed in suggesting
 * mode, as it does for a keystroke.
 */
const textblockEditEventKind = defineFreshRenderEventKind<TextblockEdit, MarkupViewExt>({
  kind: "textblockEdit",
  arbitrary: fc.oneof(
    fc.record({
      type: fc.constant("insert" as const),
      at: fc.double({ min: 0, max: 1, noNaN: true }),
      text: fc.constantFrom("new ", "clause ", "x"),
    }),
    fc.record({
      type: fc.constant("delete" as const),
      at: fc.double({ min: 0, max: 1, noNaN: true }),
      length: fc.integer({ min: 1, max: 6 }),
    }),
  ),
  apply: (rig, event) =>
    rig.edit((state) => {
      const positions = textPositions(state);
      const index = Math.min(positions.length - 1, Math.floor(event.at * positions.length));
      const at = positions.at(index);
      if (at === undefined) return null;
      if (event.type === "insert") return state.tr.insertText(event.text, at);
      const $at = state.doc.resolve(at);
      const to = Math.min($at.end(), at + event.length);
      return to > at ? state.tr.delete(at, to) : null;
    }),
});

const EVENT_KINDS: fc.Arbitrary<FreshRenderEvent<MarkupViewExt>>[] = [
  markupViewEventKind,
  textblockEditEventKind,
  loadDocumentEventKind<MarkupViewExt>(),
  tickEventKind<MarkupViewExt>(),
  visibilityEventKind<MarkupViewExt>(),
  pageGeometryEventKind<MarkupViewExt>(),
  zoomEventKind<MarkupViewExt>(),
  rerenderEventKind<MarkupViewExt>(),
];

type MarkupViewScenario = {
  document: DocumentSpec;
  mode: EditorMode;
  initialView: DisplayMode;
  leadingFrame: boolean;
  events: FreshRenderEvent<MarkupViewExt>[];
};

const scenarioArb: fc.Arbitrary<MarkupViewScenario> = fc.record({
  document: documentArb,
  mode: fc.constantFrom(...EDITOR_MODES),
  initialView: fc.constantFrom(...DISPLAY_MODES),
  leadingFrame: fc.boolean(),
  events: fc.array(fc.oneof(...EVENT_KINDS), { minLength: 1, maxLength: 12 }),
});

// ============================================================================
// ORACLES
// ============================================================================

/** The state a view reads, materialised on its own: every revision resolved as the view shows it. */
const materialise = (state: EditorState, view: DisplayMode): EditorState => {
  switch (view) {
    case "all-markup":
      return state;
    case "simple-markup":
    case "no-markup":
      return resolveAllChanges(state, "accept");
    case "original":
      return resolveAllChanges(state, "reject");
    default: {
      const exhaustive: never = view;
      return exhaustive;
    }
  }
};

const runText = (run: Run): string => (run.kind === "text" ? run.text : `[${run.kind}]`);

type Addressing = { text: string; reads: string; pmSpan: number };

/** Text runs whose editor positions do not read the text they paint. */
const misaddressedRuns = (blocks: readonly FlowBlock[], state: EditorState): Addressing[] => {
  const out: Addressing[] = [];
  const visit = (block: FlowBlock): void => {
    if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) cell.blocks.forEach(visit);
      return;
    }
    if (block.kind !== "paragraph") return;
    for (const run of block.runs) {
      if (run.kind !== "text" || run.pmStart === undefined || run.pmEnd === undefined) continue;
      const reads = state.doc.textBetween(run.pmStart, run.pmEnd);
      const pmSpan = run.pmEnd - run.pmStart;
      if (run.text.length !== pmSpan || reads !== run.text) {
        out.push({ text: run.text, reads, pmSpan });
      }
    }
  };
  blocks.forEach(visit);
  return out;
};

/**
 * Editor positions whose caret has nowhere to paint: the view's caret position
 * for them (`visibleCaretPosition`) lies inside no painted paragraph, so no
 * span, gap or empty-paragraph lookup can place it.
 */
const unpaintableCarets = (
  blocks: readonly FlowBlock[],
  state: EditorState,
  view: DisplayMode,
): number[] => {
  const ranges: PositionRange[] = [];
  const collect = (block: FlowBlock): void => {
    if (block.kind === "table") {
      for (const row of block.rows) for (const cell of row.cells) cell.blocks.forEach(collect);
      return;
    }
    if (block.kind === "paragraph" && block.pmStart !== undefined && block.pmEnd !== undefined) {
      ranges.push({ from: block.pmStart, to: block.pmEnd });
    }
  };
  blocks.forEach(collect);
  const projection = projectMarkupView(state, view);
  return textPositions(state).filter((position) => {
    const caret = visibleCaretPosition(projection, position);
    return !ranges.some(({ from, to }) => caret > from && caret < to);
  });
};

const REVISION_MARKS = new Set(["insertion", "deletion"]);

/**
 * Change bars a view owes: in Simple Markup, every paragraph whose editor text
 * holds an inserted or deleted run (the runs All Markup paints as changes)
 * carries a bar; the other views, and a document with nothing to resolve,
 * carry none. Returns the text of each paragraph that breaks the rule.
 */
const misplacedChangeBars = (
  blocks: readonly FlowBlock[],
  state: EditorState,
  view: DisplayMode,
): string[] => {
  const out: string[] = [];
  for (const block of blocks) {
    if (block.kind !== "paragraph") continue;
    const marked = block.reviewIndicator === "change-bar";
    if (view !== "simple-markup") {
      if (marked) out.push(block.runs.map((run) => runText(run)).join(""));
      continue;
    }
    const { pmStart, pmEnd } = block;
    if (pmStart === undefined || pmEnd === undefined) continue;
    let holdsRevision = false;
    state.doc.nodesBetween(pmStart, pmEnd, (node) => {
      holdsRevision ||=
        node.isText && node.marks.some((mark) => REVISION_MARKS.has(mark.type.name));
      return !holdsRevision;
    });
    if (holdsRevision && !marked) out.push(block.runs.map((run) => runText(run)).join(""));
  }
  return out;
};

// ============================================================================
// PROPERTY
// ============================================================================

type MarkupViewVerdict = {
  staleCommits: readonly unknown[];
  committed: unknown;
  fresh: unknown;
  materialised: unknown;
  misaddressed: Addressing[];
  unpaintable: number[];
  misplacedBars: string[];
};

/**
 * The view the pipeline is told to lay out. `as-input` is the product; the
 * mutant `all-markup` lays every view out as All Markup, which is what hiding
 * runs with CSS amounted to.
 */
type ViewWiring = "as-input" | "all-markup";

const runScenario = async (
  { document: spec, mode, initialView, leadingFrame, events }: MarkupViewScenario,
  wiring: ViewWiring,
): Promise<MarkupViewVerdict> => {
  const document: Document = await buildDocument(spec);
  let verdict: MarkupViewVerdict | null = null;
  withFakeTextMeasure(
    () => {
      const documentDeps = {
        document,
        styles: document.package.styles ?? null,
      } satisfies Partial<LayoutPipelineDeps<null>>;
      const rig: FreshRenderRig<MarkupViewExt> = createFreshRenderRig({
        initialDoc: createHarnessDoc(document),
        leadingFrame,
        ext: { view: initialView },
        plugins: createHarnessPlugins(document, mode),
        extraDeps: ({ ext }) => ({
          ...documentDeps,
          markupView: wiring === "as-input" ? ext.view : "all-markup",
        }),
        extraSignature: ({ ext }) => ext.view,
      });
      for (const event of events) event.apply(rig);
      rig.settle();
      const { view } = rig.ext;
      const blocks = rig.committedArtifacts()?.blocks ?? [];
      verdict = {
        staleCommits: rig.staleCommits,
        committed: rig.committed(),
        fresh: rig.fresh(),
        materialised: rig.fresh({
          state: materialise(rig.state, view),
          deps: { ...documentDeps, markupView: "all-markup" },
        }),
        misaddressed: misaddressedRuns(blocks, rig.state),
        unpaintable: unpaintableCarets(blocks, rig.state, view),
        misplacedBars: misplacedChangeBars(blocks, rig.state, view),
      };
    },
    { charWidth: sizeProportionalCharWidth },
  );
  if (!verdict) throw new Error("The scenario produced no verdict");
  return verdict;
};

/**
 * Counterexamples the property once found, replayed first on every run.
 * Seed 860847627 (path 133:213:0:0:1:1 at ten times the runs): two deleted
 * paragraph marks in a row around a paragraph holding only deleted text, so
 * the accepted story removes both joins back to back and No Markup addressed
 * the last paragraph's text from before the second join.
 */
const PINNED_SCENARIOS: MarkupViewScenario[] = [
  {
    document: {
      paragraphs: [
        {
          segments: [
            {
              kind: "format-change",
              text: "Supplier business shall goods notice kupujícího writing ",
            },
            { kind: "plain", text: "goods deliver Supplier in in Dodavatel " },
          ],
          mark: "plain",
          justified: false,
          alignmentChange: false,
        },
        {
          segments: [
            { kind: "insertion", text: "ten business ten ten notice confirmation " },
            { kind: "insertion", text: "zboží goods business business days the of goods " },
          ],
          mark: "plain",
          justified: false,
          alignmentChange: true,
        },
        {
          segments: [
            { kind: "plain", text: "confirmation notice shall days deliver ten of " },
            {
              kind: "plain",
              text: "confirmation writing writing Supplier confirmation deliver in of goods Supplier ",
            },
          ],
          mark: "deleted",
          justified: false,
          alignmentChange: false,
        },
        {
          segments: [
            {
              kind: "deletion",
              text: "business the ten deliver Dodavatel zboží ten Dodavatel business shall ",
            },
          ],
          mark: "deleted",
          justified: false,
          alignmentChange: true,
        },
        {
          segments: [{ kind: "insertion", text: "days ten shall notice business goods " }],
          mark: "inserted",
          justified: true,
          alignmentChange: false,
        },
      ],
      move: null,
      table: null,
    },
    mode: "editing",
    initialView: "no-markup",
    leadingFrame: false,
    events: [{ kind: "rerender", apply: (rig) => rig.rerender() }],
  },
  // Seed -442636085 (path 6:1:1:1:1:1:1:3:3:3:3:5:5:5:5:5:5:5:5:5:5:0:1:0:2):
  // two adjacent deleted paragraph marks with a deleted-only paragraph between them.
  {
    document: {
      paragraphs: [
        {
          segments: [{ kind: "plain", text: "the " }],
          mark: "plain",
          justified: false,
          alignmentChange: false,
        },
        {
          segments: [{ kind: "plain", text: "the " }],
          mark: "deleted",
          justified: false,
          alignmentChange: false,
        },
        {
          segments: [{ kind: "deletion", text: "the " }],
          mark: "deleted",
          justified: false,
          alignmentChange: false,
        },
        {
          segments: [{ kind: "plain", text: "the " }],
          mark: "deleted",
          justified: false,
          alignmentChange: false,
        },
      ],
      move: null,
      table: null,
    },
    mode: "editing",
    initialView: "all-markup",
    leadingFrame: false,
    events: [
      {
        kind: "markupView",
        apply: (rig) => rig.updateExt(() => ({ view: "no-markup" }), LOAD_LAYOUT.immediate),
      },
      { kind: "visible", apply: (rig) => rig.pauseFrames() },
    ],
  },
];

describe("markup views lay out the text they show", () => {
  test("addressing oracle rejects a text run whose editor span is too short", () => {
    const doc = schema.nodes.doc.create(null, [
      schema.nodes.paragraph.create(null, schema.text("abc")),
    ]);
    const state = EditorState.create({ doc });
    const blocks = [
      {
        kind: "paragraph",
        id: "generated",
        runs: [{ kind: "text", text: "abc", pmStart: 1, pmEnd: 3 }],
      },
    ] satisfies FlowBlock[];

    expect(misaddressedRuns(blocks, state)).toEqual([{ text: "abc", reads: "ab", pmSpan: 2 }]);
  });

  test("every view settles on the layout of the text it reads, through any event sequence", async () => {
    await assertProperty(
      fc.asyncProperty(scenarioArb, async (scenario) => {
        const verdict = await runScenario(scenario, "as-input");
        expect(verdict.staleCommits).toEqual([]);
        expect(verdict.committed).toEqual(verdict.fresh);
        expect(verdict.committed).toEqual(verdict.materialised);
        expect(verdict.misaddressed).toEqual([]);
        expect(verdict.unpaintable).toEqual([]);
        expect(verdict.misplacedBars).toEqual([]);
      }),
      { numRuns: 40, examples: PINNED_SCENARIOS.map((scenario) => [scenario]) },
    );
  });

  test("oracle self-check: a view that is not a layout input fails the view oracle", async () => {
    const result = await fc.check(
      fc.asyncProperty(scenarioArb, async (scenario) => {
        const verdict = await runScenario(scenario, "all-markup");
        return JSON.stringify(verdict.committed) === JSON.stringify(verdict.materialised);
      }),
      propertyConfig({ numRuns: 100 }),
    );
    expect(result.failed).toBe(true);
  });
});
