/**
 * Regression coverage for the layout pipeline's commit-vs-discard contract.
 *
 * `runLayoutPipeline` wraps compute + paint + the session-memory commit in one
 * try/catch. The hardening under test: the session is committed only after BOTH
 * layout and paint succeed, and a throw anywhere in the body (notably the paint
 * phase) discards the partial outcome AND leaves the session memory untouched,
 * so the next run re-lays-out instead of skipping on a layout it never painted.
 *
 * Headless setup: a fake DOM (createElement -> FakeElement whose `getContext`
 * returns a deterministic `measureText`) backs BOTH the canvas measurement seam
 * the pipeline installs and the paint phase, so compute and paint run without a
 * browser. Globals are swapped per-test and restored, and the canvas context +
 * measure caches are reset so nothing leaks across files.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { EditorState } from "prosemirror-state";

import {
  EXPECTED_CANCELLED_STYLE_TOGGLES,
  findStyleToggleRun,
  makeStyleToggleDefinitions,
  makeStyleToggleParagraph,
} from "../__tests__/styleToggleFlowFixture";
import type { LayoutInstrumentation } from "../layout-engine/layoutInstrumentation";
import { clearAllCaches } from "../layout-engine/measure/cache";
import {
  type HyphenationDictionaryId,
  resetHyphenationDictionaries,
} from "../layout-engine/measure/hyphenationDictionaries";
import type {
  FlowBlock,
  FootnoteContent,
  HeaderFooterContent,
  TextBoxBlock,
} from "../layout-engine/types";
import { resetCanvasContext } from "../layout-engine/measure/measureContainer";
import { convertHeaderFooterToContent } from "../layout-bridge/convert/headerFooterLayout";
import { LayoutPainter } from "../layout-painter";
import { LayoutSelectionGate } from "../paged-layout/LayoutSelectionGate";
import { twipsToPixels } from "../paged-layout/sectionGeometry";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import {
  createTemplatePreviewValuesPlugin,
  templatePreviewValuesKey,
} from "../prosemirror/plugins/templatePreviewValues";
import { schema } from "../prosemirror/schema";
import type { Footnote } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { runLayoutPipeline } from "./layoutPipeline";
import type { LayoutPipelineDeps } from "./layoutPipeline";
import { createLayoutSession } from "./layoutSession";
import type { LayoutSession } from "./layoutSession";

// --- Minimal fake DOM (paragraph rendering only) ---------------------------
// Adapted from the proven fake element in renderPage-watermark-incremental;
// the pipeline's paint path only renders plain paragraphs here.

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  parentElement: FakeElement | null = null;
  private ownText = "";
  readonly attributes = new Map<string, string>();
  readonly classList: { add: (...names: string[]) => void };
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
    this.classList = {
      add: (...names: string[]) => {
        const classes = new Set(this.className.split(/\s+/u).filter(Boolean));
        for (const name of names) {
          classes.add(name);
        }
        this.className = Array.from(classes).join(" ");
      },
    };
  }

  get firstChild(): FakeElement | null {
    return this.children.at(0) ?? null;
  }

  get innerHTML(): string {
    return this.textContent;
  }

  set innerHTML(_value: string) {
    for (const child of this.children) {
      child.parentElement = null;
    }
    this.children = [];
    this.ownText = "";
  }

  get textContent(): string {
    return this.ownText + this.children.map((child) => child.textContent).join("");
  }

  set textContent(value: string) {
    this.ownText = value;
    this.children = [];
  }

  append(...children: FakeElement[]): void {
    for (const child of children) {
      this.attach(child);
      this.children.push(child);
    }
  }

  prepend(...children: FakeElement[]): void {
    for (let index = children.length - 1; index >= 0; index--) {
      const child = children[index];
      if (!child) {
        continue;
      }
      this.attach(child);
      this.children.unshift(child);
    }
  }

  appendChild(child: FakeElement): FakeElement {
    this.append(child);
    return child;
  }

  insertBefore(child: FakeElement, before: FakeElement | null): FakeElement {
    if (!before) {
      this.append(child);
      return child;
    }
    const index = this.children.indexOf(before);
    if (index === -1) {
      this.append(child);
      return child;
    }
    this.attach(child);
    this.children.splice(index, 0, child);
    return child;
  }

  remove(): void {
    const parent = this.parentElement;
    if (!parent) {
      return;
    }
    const index = parent.children.indexOf(this);
    if (index !== -1) {
      parent.children.splice(index, 1);
    }
    this.parentElement = null;
  }

  setAttribute(name: string, value: string): void {
    this.attributes.set(name, value);
  }

  dispatchEvent(_event: Event): boolean {
    return true;
  }

  getContext(_contextId: "2d"): CanvasRenderingContext2D {
    return {
      font: "",
      measureText(text: string) {
        return {
          width: text.length * 7,
          actualBoundingBoxAscent: 8,
          actualBoundingBoxDescent: 2,
        };
      },
    } as CanvasRenderingContext2D;
  }

  getBoundingClientRect(): DOMRect {
    return {
      bottom: 0,
      height: 0,
      left: 0,
      right: 0,
      top: 0,
      width: 0,
      x: 0,
      y: 0,
      toJSON: () => ({}),
    } as DOMRect;
  }

  querySelector(selector: string): FakeElement | null {
    return findByClass(this, classFromSelector(selector));
  }

  querySelectorAll(selector: string): FakeElement[] {
    const out: FakeElement[] = [];
    collectByClass(this, classFromSelector(selector), out);
    return out;
  }

  private attach(child: FakeElement): void {
    child.remove();
    child.parentElement = this;
  }
}

const CLASS_SELECTOR_RE = /\.(?<cls>[\w-]+)/u;

const classFromSelector = (selector: string): string =>
  CLASS_SELECTOR_RE.exec(selector)?.groups?.["cls"] ?? "";

const hasClass = (element: FakeElement, className: string): boolean =>
  element.className.split(/\s+/u).includes(className);

function findByClass(root: FakeElement, className: string): FakeElement | null {
  for (const child of root.children) {
    if (hasClass(child, className)) {
      return child;
    }
    const inner = findByClass(child, className);
    if (inner) {
      return inner;
    }
  }
  return null;
}

function collectByClass(root: FakeElement, className: string, out: FakeElement[]): void {
  for (const child of root.children) {
    if (hasClass(child, className)) {
      out.push(child);
    }
    collectByClass(child, className, out);
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
  createTextNode(text: string): FakeElement {
    const node = new FakeElement("#text");
    node.textContent = text;
    return node;
  },
};

// The fake element structurally covers the subset of the DOM the paint phase
// touches; the pipeline only forwards the container to `renderPages`.
const asContainer = (value: object): HTMLDivElement =>
  // oxlint-disable-next-line typescript/no-unsafe-type-assertion -- test fake DOM container
  value as unknown as HTMLDivElement;

class FakeIntersectionObserver {
  observe(): void {}
  unobserve(): void {}
  disconnect(): void {}
}

class FakeCustomEvent<T> {
  readonly bubbles: boolean;
  readonly cancelable: boolean;
  readonly detail: T | null;
  readonly type: string;

  constructor(type: string, init: CustomEventInit<T> = {}) {
    this.type = type;
    this.detail = init.detail ?? null;
    this.bubbles = init.bubbles ?? false;
    this.cancelable = init.cancelable ?? false;
  }
}

// --- Global swap plumbing ---------------------------------------------------

const originalGlobalDescriptors = new Map<string, PropertyDescriptor | undefined>();

function setGlobal(name: string, value: unknown): void {
  if (!originalGlobalDescriptors.has(name)) {
    originalGlobalDescriptors.set(name, Object.getOwnPropertyDescriptor(globalThis, name));
  }
  Object.defineProperty(globalThis, name, {
    configurable: true,
    value,
    writable: true,
  });
}

function restoreGlobals(): void {
  for (const [name, descriptor] of originalGlobalDescriptors) {
    if (descriptor) {
      Object.defineProperty(globalThis, name, descriptor);
    } else {
      Reflect.deleteProperty(globalThis, name);
    }
  }
  originalGlobalDescriptors.clear();
}

// --- Pipeline deps + state builders ----------------------------------------

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS = {
  top: 72,
  right: 72,
  bottom: 72,
  left: 72,
  header: 36,
  footer: 36,
};

const makeState = (): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Hello world")]),
      schema.node("paragraph", null, [schema.text("Second paragraph here.")]),
    ]),
  });

// An inline `{{ terms }}` marker with running text on both sides, plus a
// paragraph under it whose position proves the marker's paragraph took the
// height its value needs.
const MULTILINE_PREVIEW_VALUE = "First line.\nSecond line.\nThird line.";

const makeMultilinePreviewState = (value: string): EditorState => {
  const state = EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Intro {{ terms }} end.")]),
      schema.node("paragraph", null, [schema.text("Following paragraph.")]),
    ]),
    plugins: [createTemplatePreviewValuesPlugin()],
  });
  return state.apply(
    state.tr.setMeta(templatePreviewValuesKey, {
      preview: { values: { terms: value }, mode: "plain" },
    }),
  );
};

const makeTwoPageState = (): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node("paragraph", null, [schema.text("Odd page")]),
      schema.node("pageBreak"),
      schema.node("paragraph", null, [schema.text("Even page")]),
    ]),
  });

const makeSectionedState = (): EditorState =>
  EditorState.create({
    doc: schema.node("doc", null, [
      schema.node(
        "paragraph",
        {
          _sectionProperties: {
            sectionStart: "continuous",
            marginTop: 1080,
            marginRight: 1080,
            marginBottom: 1080,
            marginLeft: 1080,
            headerDistance: 540,
            footerDistance: 540,
          },
        },
        [schema.text("First section")],
      ),
      schema.node("paragraph", null, [schema.text("Second section")]),
    ]),
  });

const makeImplicitSectionBoundaryState = (
  finalSectionStart?: "continuous" | "nextPage",
): EditorState =>
  EditorState.create({
    doc: schema.node("doc", { _finalSectionStart: finalSectionStart ?? null }, [
      schema.node(
        "paragraph",
        {
          _sectionProperties: {},
        },
        [schema.text("First section")],
      ),
      schema.node("paragraph", null, [schema.text("Second section")]),
    ]),
  });

const makeSectionBoundaryDocument = (sectionStart: "continuous" | "nextPage") => {
  const document = createEmptyDocument();
  document.package.document.sections = [
    { properties: {}, content: [] },
    { properties: { sectionStart }, content: [] },
  ];
  return document;
};

// A template document whose middle block is conditional, plus the host's
// verdict on it. The preview plugin turns the verdict into the hidden span the
// pipeline reads off the editor state.
const CONDITIONAL_PARAGRAPHS = ["{% if premium %}", "Premium terms.", "{% endif %}", "Tail."];
// PM position of the "Tail." paragraph node: every paragraph before it spans
// its text plus the node's open and close tokens.
const TAIL_PM_START = CONDITIONAL_PARAGRAPHS.slice(0, -1).reduce(
  (pos, text) => pos + text.length + 2,
  0,
);

const makeConditionalPreviewState = (
  premium: boolean,
  mode: "highlighted" | "plain",
): EditorState => {
  const state = EditorState.create({
    doc: schema.node(
      "doc",
      null,
      CONDITIONAL_PARAGRAPHS.map((text) => schema.node("paragraph", null, [schema.text(text)])),
    ),
    plugins: [createTemplatePreviewValuesPlugin()],
  });
  return state.apply(
    state.tr.setMeta(templatePreviewValuesKey, {
      preview: { values: {}, mode, conditions: { premium } },
    }),
  );
};

const paragraphText = (block: FlowBlock): string =>
  block.kind === "paragraph"
    ? block.runs.map((run) => (run.kind === "text" ? run.text : "")).join("")
    : block.kind;

const makePreparedHeaderFooter = (height: number): HeaderFooterContent => ({
  blocks: [],
  measures: [],
  height,
  visualTop: 0,
  visualBottom: height,
  marginPushTop: 0,
  marginPushBottom: height,
});

type DepsOverrides = Partial<LayoutPipelineDeps<null>>;

const makeDeps = (
  session: LayoutSession,
  overrides: DepsOverrides = {},
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
  layout: null,
  hfPMs: null,
  painter: null,
  pagesContainer: null,
  session,
  renderHfFromContentOrPm: () => undefined,
  renderHeaderFooterContentByRId: () => undefined,
  documentFontsAreLoaded: () => true,
  buildFootnoteRenderItems: () => new Map(),
  describeInvalidHighlightMarks: () => "",
  emptyTemplatePreviewEntries: [],
  emptyTemplatePreviewHidden: [],
  hyphenationReadiness: { track: () => undefined, cancel: () => undefined },
  ...overrides,
});

// Capture the instrumentation callbacks the pipeline fires on
// complete/error without spying on a directly-imported module.
let layoutCompletes: { reason: string }[] = [];
let layoutErrors: { message: string; reason: string }[] = [];

describe("runLayoutPipeline", () => {
  beforeEach(() => {
    setGlobal("document", fakeDocument);
    setGlobal("HTMLElement", FakeElement);
    setGlobal("IntersectionObserver", FakeIntersectionObserver);
    setGlobal("CustomEvent", FakeCustomEvent);
    setGlobal("window", { innerHeight: 900 });
    resetCanvasContext();
    clearAllCaches();
    layoutCompletes = [];
    layoutErrors = [];
    const instrumentation: LayoutInstrumentation = {
      onLayoutComplete: (event) => {
        layoutCompletes.push(event);
      },
      onLayoutError: (event) => {
        layoutErrors.push(event);
      },
    };
    globalThis.__folioLayoutInstrumentation = instrumentation;
  });

  afterEach(() => {
    globalThis.__folioLayoutInstrumentation = undefined;
    restoreGlobals();
    resetCanvasContext();
    clearAllCaches();
  });

  test("commits the session memory and returns a painted outcome on success", () => {
    const session = createLayoutSession();
    const state = makeState();
    const container = fakeDocument.createElement("div");
    const deps = makeDeps(session, {
      painter: new LayoutPainter(),
      pagesContainer: asContainer(container),
    });

    const outcome = runLayoutPipeline(deps, state);

    // Outcome is fully populated, including the block lookup the paint phase
    // builds when a painter is attached.
    expect(outcome.blocks?.length ?? 0).toBeGreaterThan(0);
    expect(outcome.measures?.length ?? 0).toBeGreaterThan(0);
    expect(outcome.layout?.pages.length ?? 0).toBeGreaterThan(0);
    expect(outcome.blockLookup).toBeInstanceOf(Map);
    expect(outcome.blockLookup?.size ?? 0).toBeGreaterThan(0);

    // Session memory is committed only after layout AND paint succeed.
    expect(session.artifacts).not.toBeNull();
    expect(session.artifacts?.blocks.length ?? 0).toBeGreaterThan(0);
    expect(session.artifacts?.measures.length ?? 0).toBeGreaterThan(0);
    expect(session.lastEditorState).toBe(state);
    expect(session.lastPmDoc).toBe(state.doc);
    expect(session.usedLoadedFonts).toBe(true);
    expect(session.lastTemplatePreview).toEqual({ entries: [], hidden: [], mode: "plain" });

    expect(layoutCompletes).toHaveLength(1);
    expect(layoutErrors).toHaveLength(0);
  });

  test("drops a hidden conditional block from the pages and lifts what follows", () => {
    // Same document and same layout inputs twice; only the host's verdict on
    // the `{% if premium %}` block differs, so any difference in the pages is
    // the hiding and nothing else.
    const laidOut = (premium: boolean, mode: "highlighted" | "plain" = "plain") => {
      const outcome = runLayoutPipeline(
        makeDeps(createLayoutSession()),
        makeConditionalPreviewState(premium, mode),
      );
      const fragments = (outcome.layout?.pages ?? []).flatMap((page) => page.fragments);
      return { blocks: outcome.blocks ?? [], fragments };
    };

    const hidden = laidOut(false);
    // `highlighted` keeps every directive tag, so this is the document as authored.
    const shown = laidOut(true, "highlighted");

    // The opener, its body and the closer leave the flow; the tail stays.
    expect(hidden.blocks.map(paragraphText)).toEqual(["Tail."]);
    expect(shown.blocks.map(paragraphText)).toEqual([
      "{% if premium %}",
      "Premium terms.",
      "{% endif %}",
      "Tail.",
    ]);

    // Nothing is painted for a dropped block: no fragment addresses a PM
    // position inside the hidden span.
    expect(hidden.fragments.map((fragment) => fragment.pmStart)).toEqual([TAIL_PM_START]);

    // The tail keeps its PM position and moves up the page by the three
    // paragraphs that are no longer above it.
    const yOf = (fragments: { pmStart?: number; y: number }[]) =>
      fragments.find((fragment) => fragment.pmStart === TAIL_PM_START)?.y;
    expect(yOf(hidden.fragments)).toBeLessThan(yOf(shown.fragments) ?? 0);

    // A block that does apply keeps its body in `plain` mode and loses only its
    // tag paragraphs, so the pages read as the generated document.
    const tagsDropped = laidOut(true);
    expect(tagsDropped.blocks.map(paragraphText)).toEqual(["Premium terms.", "Tail."]);
    expect(yOf(tagsDropped.fragments)).toBeLessThan(yOf(shown.fragments) ?? 0);
  });

  test("gives a multi-line preview value a line per line instead of overlapping", () => {
    const laidOut = (value: string) => {
      const outcome = runLayoutPipeline(
        makeDeps(createLayoutSession()),
        makeMultilinePreviewState(value),
      );
      return outcome.layout?.pages ?? [];
    };

    const pages = laidOut(MULTILINE_PREVIEW_VALUE);
    expect(pages.length).toBeGreaterThan(0);

    // Nothing is stacked on top of anything else: two fragments sharing a y on
    // one page is the overlap this guards against.
    for (const page of pages) {
      const ys = page.fragments.map((fragment) => fragment.y);
      expect(new Set(ys).size).toBe(ys.length);
    }

    // The marker's paragraph is three lines tall, so the paragraph under it
    // sits a full three lines lower than it would with a one-line value. Before
    // the fix the value measured as one line and painted as three, and this gap
    // stayed at one line while the text ran over the paragraph below.
    const gapUnder = (value: string): number => {
      const fragments = laidOut(value).flatMap((page) => page.fragments);
      const [marker, following] = fragments;
      if (!marker || !following) {
        throw new Error("expected both paragraphs to be laid out");
      }
      return following.y - marker.y;
    };
    const singleLineGap = gapUnder("Only line.");
    expect(gapUnder(MULTILINE_PREVIEW_VALUE)).toBeGreaterThanOrEqual(singleLineGap * 3);
  });

  test("commits the session without a block lookup when no painter is attached", () => {
    const session = createLayoutSession();
    const state = makeState();
    const deps = makeDeps(session); // painter + pagesContainer default to null

    const outcome = runLayoutPipeline(deps, state);

    expect(outcome.layout?.pages.length ?? 0).toBeGreaterThan(0);
    // Paint phase is skipped, so no block lookup is produced.
    expect(outcome.blockLookup).toBeUndefined();
    // The layout still succeeded, so the session is committed.
    expect(session.artifacts).not.toBeNull();
    expect(session.lastEditorState).toBe(state);
    expect(layoutCompletes).toHaveLength(1);
  });

  test("resolves the complete character-style toggle cascade in body layout", () => {
    const styles = makeStyleToggleDefinitions();
    const document = createEmptyDocument();
    document.package.styles = styles;
    document.package.document.content = [makeStyleToggleParagraph()];
    const state = EditorState.create({
      doc: schema.nodeFromJSON(toProseDoc(document, { styles }).toJSON()),
    });

    const outcome = runLayoutPipeline(makeDeps(createLayoutSession(), { document, styles }), state);

    expect(findStyleToggleRun(outcome.blocks ?? [])).toMatchObject(
      EXPECTED_CANCELLED_STYLE_TOGGLES,
    );
    expect(layoutErrors).toHaveLength(0);
  });

  test("forwards styles through top-level header, footer, and footnote layout", () => {
    const styles = makeStyleToggleDefinitions();
    const document = createEmptyDocument();
    document.package.styles = styles;
    document.package.footnotes = [
      {
        type: "footnote",
        id: 1,
        noteType: "normal",
        content: [makeStyleToggleParagraph()],
      },
    ];
    const footnoteRef = schema.mark("footnoteRef", { id: "1", noteType: "footnote" });
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [schema.text("Body"), schema.text("1", [footnoteRef])]),
      ]),
    });
    const preparedBySection = new Map<"header" | "footer", HeaderFooterContent>();
    let capturedFootnotes: Map<number, FootnoteContent> | undefined;
    const container = fakeDocument.createElement("div");

    runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document,
        styles,
        headerContent: {
          type: "header",
          hdrFtrType: "default",
          content: [makeStyleToggleParagraph()],
        },
        footerContent: {
          type: "footer",
          hdrFtrType: "default",
          content: [makeStyleToggleParagraph()],
        },
        painter: new LayoutPainter(),
        pagesContainer: asContainer(container),
        renderHfFromContentOrPm: (content, _rId, _hfPMs, contentWidth, metrics, options) => {
          if (!content) {
            return undefined;
          }
          const prepared = convertHeaderFooterToContent(content, contentWidth, metrics, options);
          if (prepared) {
            preparedBySection.set(metrics.section, prepared);
          }
          return prepared;
        },
        buildFootnoteRenderItems: (_pageMap, contentMap) => {
          capturedFootnotes = contentMap;
          return new Map();
        },
      }),
      state,
    );

    expect(findStyleToggleRun(preparedBySection.get("header")?.blocks ?? [])).toMatchObject(
      EXPECTED_CANCELLED_STYLE_TOGGLES,
    );
    expect(findStyleToggleRun(preparedBySection.get("footer")?.blocks ?? [])).toMatchObject(
      EXPECTED_CANCELLED_STYLE_TOGGLES,
    );
    expect(findStyleToggleRun(capturedFootnotes?.get(1)?.blocks ?? [])).toMatchObject(
      EXPECTED_CANCELLED_STYLE_TOGGLES,
    );
    expect(layoutErrors).toHaveLength(0);
  });

  test("remaps body note markers to sequential reference-order display numbers", () => {
    // Non-contiguous, out-of-order footnote ids (5, 2, 9 referenced in that
    // order) must display 1, 2, 3 in both the body marker and the footnote
    // area; the positive-id continuationNotice must not shift the numbering.
    const session = createLayoutSession();
    const footnoteMark = (id: number, noteType: "footnote" | "endnote") =>
      schema.mark("footnoteRef", { id: String(id), noteType });
    const state = EditorState.create({
      doc: schema.node("doc", null, [
        schema.node("paragraph", null, [
          schema.text("Alpha"),
          schema.text("5", [footnoteMark(5, "footnote")]),
          schema.text(" beta"),
          schema.text("2", [footnoteMark(2, "footnote")]),
          schema.text(" gamma"),
          schema.text("9", [footnoteMark(9, "footnote")]),
          schema.text(" delta"),
          schema.text("8", [footnoteMark(8, "endnote")]),
        ]),
      ]),
    });

    const noteParagraph = (text: string): Footnote["content"] => [
      { type: "paragraph", content: [{ type: "run", content: [{ type: "text", text }] }] },
    ];
    const doc = createEmptyDocument();
    doc.package.footnotes = [
      { type: "footnote", id: 1, noteType: "continuationNotice", content: noteParagraph("cont") },
      { type: "footnote", id: 2, noteType: "normal", content: noteParagraph("two") },
      { type: "footnote", id: 5, noteType: "normal", content: noteParagraph("five") },
      { type: "footnote", id: 9, noteType: "normal", content: noteParagraph("nine") },
    ];
    doc.package.endnotes = [
      { type: "endnote", id: 8, noteType: "normal", content: noteParagraph("end") },
    ];

    let capturedContentMap: Map<number, FootnoteContent> | undefined;
    const container = fakeDocument.createElement("div");
    const deps = makeDeps(session, {
      document: doc,
      painter: new LayoutPainter(),
      pagesContainer: asContainer(container),
      buildFootnoteRenderItems: (_pageMap, contentMap) => {
        capturedContentMap = contentMap;
        return new Map();
      },
    });

    const outcome = runLayoutPipeline(deps, state);

    const bodyMarkers: { noteId: number; text: string }[] = [];
    for (const block of outcome.blocks ?? []) {
      if (block.kind !== "paragraph") {
        continue;
      }
      for (const run of block.runs) {
        if (run.kind !== "text") {
          continue;
        }
        const noteId = run.footnoteRefId ?? run.endnoteRefId;
        if (noteId !== undefined) {
          bodyMarkers.push({ noteId, text: run.text });
        }
      }
    }
    expect(bodyMarkers).toEqual([
      { noteId: 5, text: "1" },
      { noteId: 2, text: "2" },
      { noteId: 9, text: "3" },
      { noteId: 8, text: "i" },
    ]);

    // The footnote area numbering comes from the same map as the body markers.
    expect(capturedContentMap?.get(5)?.displayNumber).toBe(1);
    expect(capturedContentMap?.get(2)?.displayNumber).toBe(2);
    expect(capturedContentMap?.get(9)?.displayNumber).toBe(3);
    expect(capturedContentMap?.has(1)).toBe(false);
    expect(layoutErrors).toHaveLength(0);
  });

  test("moves body content above an overflowing default footer", () => {
    const state = makeState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        footerContent: { type: "footer", hdrFtrType: "default", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "footer"
            ? {
                blocks: [],
                measures: [],
                height: 120,
                visualTop: 0,
                visualBottom: 120,
                marginPushTop: 0,
                marginPushBottom: 120,
              }
            : undefined,
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.bottom).toBe(MARGINS.footer + 120);
  });

  test("keeps authored body margin when a default footer stays below it", () => {
    const state = makeState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        footerContent: { type: "footer", hdrFtrType: "default", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "footer"
            ? {
                blocks: [],
                measures: [],
                height: 20,
                visualTop: 0,
                visualBottom: 20,
                marginPushTop: 0,
                marginPushBottom: 20,
              }
            : undefined,
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.bottom).toBe(MARGINS.bottom);
  });

  test("moves body content below an overflowing default header", () => {
    const state = makeState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        headerContent: { type: "header", hdrFtrType: "default", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "header"
            ? {
                blocks: [],
                measures: [],
                height: 120,
                visualTop: 0,
                visualBottom: 120,
                marginPushTop: 0,
                marginPushBottom: 120,
              }
            : undefined,
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.top).toBe(MARGINS.header + 120);
    expect(outcome.layout?.pages.at(0)?.fragments.at(0)?.y).toBe(MARGINS.header + 120);
  });

  test("keeps authored body margin when a default header stays above it", () => {
    const state = makeState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        headerContent: { type: "header", hdrFtrType: "default", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "header"
            ? {
                blocks: [],
                measures: [],
                height: 20,
                visualTop: 0,
                visualBottom: 20,
                marginPushTop: 0,
                marginPushBottom: 20,
              }
            : undefined,
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.top).toBe(MARGINS.top);
    expect(outcome.layout?.pages.at(0)?.fragments.at(0)?.y).toBe(MARGINS.top);
  });

  test("moves body content below a page-top header wrap band", () => {
    const state = makeState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        headerContent: { type: "header", hdrFtrType: "default", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "header"
            ? {
                blocks: [],
                measures: [],
                height: 0,
                marginPushTop: 0,
                marginPushBottom: 0,
                bodyTopClearance: 160,
              }
            : undefined,
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.top).toBe(160);
    expect(outcome.layout?.pages.at(0)?.fragments.at(0)?.y).toBe(160);
  });

  describe("page-frame anchored text boxes under header/footer clearance", () => {
    // A header or footer taller than its margin pushes the body box in, and the
    // margin frames an anchor can be positioned in follow that box: `topMargin`
    // spans the page top to the pushed body top, `bottomMargin` the pushed body
    // bottom to the page bottom, and `margin` the pushed body box itself. Only
    // `page` stays on the sheet.
    type Vertical = NonNullable<NonNullable<TextBoxBlock["position"]>["vertical"]>;
    type Furniture = "header" | "footer";

    const BOX_HEIGHT = 30;
    const FURNITURE_HEIGHT = 150;
    const tallFurniture = makePreparedHeaderFooter(FURNITURE_HEIGHT);
    const PUSHED_TOP = MARGINS.header + FURNITURE_HEIGHT;
    const PUSHED_BOTTOM = MARGINS.footer + FURNITURE_HEIGHT;
    // 1/4 inch: 24px at 96 DPI.
    const QUARTER_INCH_EMU = 228_600;

    const makeAnchoredBoxState = (vertical: Vertical) =>
      EditorState.create({
        doc: schema.node("doc", null, [
          schema.node("paragraph", null, [schema.text("Before the box.")]),
          schema.node(
            "textBox",
            {
              width: 200,
              height: BOX_HEIGHT,
              displayMode: "float",
              wrapType: "inFront",
              position: { horizontal: { relativeTo: "page", posOffset: 0 }, vertical },
            },
            [schema.node("paragraph", null, [schema.text("Boxed")])],
          ),
          schema.node("paragraph", null, [schema.text("After the box.")]),
        ]),
      });

    const boxY = (vertical: Vertical, furniture: Furniture): number | undefined => {
      const outcome = runLayoutPipeline(
        makeDeps(createLayoutSession(), {
          ...(furniture === "header"
            ? { headerContent: { type: "header", hdrFtrType: "default", content: [] } }
            : { footerContent: { type: "footer", hdrFtrType: "default", content: [] } }),
          renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
            hf && metrics.section === furniture ? tallFurniture : undefined,
        }),
        makeAnchoredBoxState(vertical),
      );
      expect(layoutErrors.map((error) => error.message)).toEqual([]);
      const page = outcome.layout?.pages.at(0);
      if (furniture === "header") {
        expect(page?.margins.top).toBe(PUSHED_TOP);
      } else {
        expect(page?.margins.bottom).toBe(PUSHED_BOTTOM);
      }
      return page?.fragments.find((fragment) => fragment.kind === "textBox")?.y;
    };

    test.each<[string, Vertical, Furniture, number]>([
      [
        "topMargin, bottom-aligned",
        { relativeTo: "topMargin", align: "bottom" },
        "header",
        PUSHED_TOP - BOX_HEIGHT,
      ],
      ["topMargin, offset", { relativeTo: "topMargin", posOffset: QUARTER_INCH_EMU }, "header", 24],
      ["margin, offset 0", { relativeTo: "margin", posOffset: 0 }, "header", PUSHED_TOP],
      ["page, offset", { relativeTo: "page", posOffset: QUARTER_INCH_EMU }, "header", 24],
      [
        "bottomMargin, offset 0",
        { relativeTo: "bottomMargin", posOffset: 0 },
        "footer",
        PAGE_SIZE.h - PUSHED_BOTTOM,
      ],
    ])("%s", (_label, vertical, furniture, expectedY) => {
      expect(boxY(vertical, furniture)).toBe(expectedY);
    });
  });

  test("uses authored margins for a blank even-page header and footer", () => {
    const content = { type: "header", hdrFtrType: "default", content: [] } as const;
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        headerContent: content,
        footerContent: { ...content, type: "footer" },
        sectionHeaderFooterRefs: [
          {
            evenAndOddHeaders: true,
            headerDefault: "odd-header",
            footerDefault: "odd-footer",
          },
        ],
        renderHfFromContentOrPm: (hf) => (hf ? makePreparedHeaderFooter(120) : undefined),
        renderHeaderFooterContentByRId: (_source, _hfPMs, _width, metrics) =>
          new Map([
            [
              metrics.section === "header" ? "odd-header" : "odd-footer",
              makePreparedHeaderFooter(120),
            ],
          ]),
      }),
      makeTwoPageState(),
    );

    expect(outcome.layout?.pages).toHaveLength(2);
    expect(outcome.layout?.pages[0]?.margins.top).toBe(MARGINS.header + 120);
    expect(outcome.layout?.pages[0]?.margins.bottom).toBe(MARGINS.footer + 120);
    expect(outcome.layout?.pages[1]?.margins).toEqual(MARGINS);
  });

  test("moves a section body between its referenced default header and footer", () => {
    const state = makeSectionedState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        sectionHeaderFooterRefs: [
          { headerDefault: "first-header", footerDefault: "first-footer" },
          { headerDefault: "second-header", footerDefault: "second-footer" },
        ],
        renderHeaderFooterContentByRId: (_source, _hfPMs, _contentWidth, metrics) => {
          if (metrics.section === "header") {
            return new Map([
              ["first-header", makePreparedHeaderFooter(120)],
              ["second-header", makePreparedHeaderFooter(20)],
            ]);
          }
          return new Map([
            ["first-footer", makePreparedHeaderFooter(120)],
            ["second-footer", makePreparedHeaderFooter(20)],
          ]);
        },
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.top).toBe(MARGINS.header + 120);
    expect(outcome.layout?.pages.at(0)?.margins.bottom).toBe(MARGINS.footer + 120);
    expect(outcome.layout?.pages.at(0)?.fragments.at(0)?.y).toBe(MARGINS.header + 120);
  });

  test("keeps a section's authored margin when its referenced header fits above it", () => {
    const state = makeSectionedState();
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        sectionHeaderFooterRefs: [{ headerDefault: "compact-header" }, {}],
        renderHeaderFooterContentByRId: () =>
          new Map([["compact-header", makePreparedHeaderFooter(20)]]),
      }),
      state,
    );

    expect(outcome.layout?.pages.at(0)?.margins.top).toBe(MARGINS.top);
    expect(outcome.layout?.pages.at(0)?.fragments.at(0)?.y).toBe(MARGINS.top);
  });

  test("moves only the title-page body below an overflowing first-page header", () => {
    const state = makeState();
    const baseline = runLayoutPipeline(makeDeps(createLayoutSession()), state);
    const withTallFirstHeader = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        sectionProperties: { titlePg: true },
        firstPageHeaderContent: { type: "header", hdrFtrType: "first", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "header"
            ? {
                blocks: [],
                measures: [],
                height: 120,
                visualTop: 0,
                visualBottom: 120,
                marginPushTop: 0,
                marginPushBottom: 120,
              }
            : undefined,
      }),
      state,
    );

    const baselineFirstPage = baseline.layout?.pages.at(0);
    const pushedFirstPage = withTallFirstHeader.layout?.pages.at(0);
    expect(pushedFirstPage?.margins.top).toBe(MARGINS.header + 120);
    expect(pushedFirstPage?.fragments.at(0)?.y).toBeGreaterThan(
      baselineFirstPage?.fragments.at(0)?.y ?? 0,
    );
  });

  test("moves only the title-page body above an overflowing first-page footer", () => {
    const state = makeState();
    const baseline = runLayoutPipeline(makeDeps(createLayoutSession()), state);
    const withTallFirstFooter = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        sectionProperties: { titlePg: true },
        firstPageFooterContent: { type: "footer", hdrFtrType: "first", content: [] },
        renderHfFromContentOrPm: (hf, _rId, _hfPMs, _contentWidth, metrics) =>
          hf && metrics.section === "footer"
            ? {
                blocks: [],
                measures: [],
                height: 120,
                visualTop: 0,
                visualBottom: 120,
                marginPushTop: 0,
                marginPushBottom: 120,
              }
            : undefined,
      }),
      state,
    );

    const baselineFirstPage = baseline.layout?.pages.at(0);
    const pushedFirstPage = withTallFirstFooter.layout?.pages.at(0);
    expect(pushedFirstPage?.margins.bottom).toBe(MARGINS.footer + 120);
    expect(pushedFirstPage?.margins.bottom).toBeGreaterThan(baselineFirstPage?.margins.bottom ?? 0);
  });

  test("threads the final section geometry into layout", () => {
    const document = createEmptyDocument();
    document.package.document.sections = [
      {
        properties: {
          pageWidth: 16_860,
          pageHeight: 11_920,
          orientation: "landscape",
          marginTop: 1_134,
          marginRight: 1_134,
          marginBottom: 1_134,
          marginLeft: 1_134,
        },
        content: [],
      },
    ];

    const outcome = runLayoutPipeline(makeDeps(createLayoutSession(), { document }), makeState());

    expect(outcome.layout?.pages.at(0)?.size).toEqual({
      w: twipsToPixels(16_860),
      h: twipsToPixels(11_920),
    });
  });

  test("threads document automatic-hyphenation settings into paragraph layout", () => {
    const document = createEmptyDocument();
    document.package.settings = {
      defaultTabStop: 720,
      autoHyphenation: true,
      doNotHyphenateCaps: true,
      consecutiveHyphenLimit: 2,
      hyphenationZoneTwips: 720,
    };

    const outcome = runLayoutPipeline(makeDeps(createLayoutSession(), { document }), makeState());
    const paragraph = outcome.blocks?.find((block) => block.kind === "paragraph");

    expect(paragraph?.attrs?.automaticHyphenation).toEqual({
      enabled: true,
      doNotHyphenateCaps: true,
      consecutiveLineLimit: 2,
      hyphenationZoneTwips: 720,
    });
  });

  test("hands the hyphenation dictionaries a run lacked to its editor", () => {
    resetHyphenationDictionaries();
    const document = createEmptyDocument();
    document.package.settings = { autoHyphenation: true };
    document.package.document.content = [
      {
        type: "paragraph",
        content: [
          {
            type: "run",
            formatting: { language: { val: "sk-SK" } },
            content: [{ type: "text", text: "najneobhospodarovávateľnejší ".repeat(12).trim() }],
          },
        ],
      },
    ];
    const tracked: HyphenationDictionaryId[][] = [];
    const state = EditorState.create({
      doc: schema.nodeFromJSON(toProseDoc(document).toJSON()),
    });

    runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document,
        hyphenationReadiness: {
          track: (missing) => tracked.push([...missing]),
          cancel: () => undefined,
        },
      }),
      state,
    );

    expect(tracked).toEqual([["sk"]]);
  });

  test("paints package-owned picture watermarks", () => {
    const imageTarget = "word/media/watermark.png";
    const imageSrc = "data:image/png;base64,iVBORw0KGgo=";
    const document = createEmptyDocument();
    document.package.media = new Map([
      [
        imageTarget,
        {
          path: imageTarget,
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: imageSrc,
        },
      ],
    ]);
    document.package.headers = new Map([
      [
        "rId1",
        {
          type: "header",
          hdrFtrType: "default",
          content: [],
          watermark: {
            kind: "picture",
            imageRId: "rId1",
            imageTarget,
          },
        },
      ],
    ]);
    const pagesContainer = new FakeElement("div");

    runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document,
        painter: new LayoutPainter(),
        pagesContainer: asContainer(pagesContainer),
      }),
      makeState(),
    );

    const watermarkLayer = pagesContainer.querySelector(".layout-page-watermark");
    expect(watermarkLayer?.children.at(0)).toMatchObject({
      src: imageSrc,
      tagName: "img",
    });
  });

  test("does not resolve linked picture watermarks through package media", () => {
    const imageTarget = "https://example.com/watermark.png";
    const document = createEmptyDocument();
    document.package.media = new Map([
      [
        imageTarget,
        {
          path: imageTarget,
          mimeType: "image/png",
          data: new ArrayBuffer(0),
          dataUrl: "data:image/png;base64,iVBORw0KGgo=",
        },
      ],
    ]);
    document.package.headers = new Map([
      [
        "rId1",
        {
          type: "header",
          hdrFtrType: "default",
          content: [],
          watermark: {
            kind: "picture",
            imageRId: "rId1",
            imageTarget,
            imageTargetExternal: true,
          },
        },
      ],
    ]);
    const pagesContainer = new FakeElement("div");

    runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document,
        painter: new LayoutPainter(),
        pagesContainer: asContainer(pagesContainer),
      }),
      makeState(),
    );

    expect(pagesContainer.querySelector(".layout-page-watermark")).toBeNull();
  });

  test("applies the final section start mode at its preceding implicit boundary", () => {
    const continuous = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document: makeSectionBoundaryDocument("continuous"),
        sectionProperties: {},
      }),
      makeImplicitSectionBoundaryState("continuous"),
    );
    const nextPage = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document: makeSectionBoundaryDocument("nextPage"),
        sectionProperties: {},
      }),
      makeImplicitSectionBoundaryState("nextPage"),
    );

    expect(continuous.layout?.pages).toHaveLength(1);
    expect(nextPage.layout?.pages).toHaveLength(2);
  });

  test("uses the supplied section start when document metadata is unavailable", () => {
    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document: null,
        sectionProperties: { sectionStart: "continuous" },
      }),
      makeImplicitSectionBoundaryState(),
    );

    expect(outcome.layout?.pages).toHaveLength(1);
  });

  test("uses the supplied vertical alignment when document metadata is unavailable", () => {
    const top = runLayoutPipeline(
      makeDeps(createLayoutSession(), { document: null, sectionProperties: {} }),
      makeState(),
    );
    const centered = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document: null,
        sectionProperties: { verticalAlign: "center" },
      }),
      makeState(),
    );

    expect(centered.layout?.pages[0]?.fragments[0]?.y).toBeGreaterThan(
      top.layout?.pages[0]?.fragments[0]?.y ?? Number.POSITIVE_INFINITY,
    );
  });

  test("uses the referenced final-section footer for body clearance", () => {
    const document = createEmptyDocument();
    document.package.document.sections = [
      {
        properties: {
          pageWidth: 12_240,
          pageHeight: 15_840,
          marginTop: 1_080,
          marginRight: 1_080,
          marginBottom: 1_440,
          marginLeft: 1_080,
          footerDistance: 720,
        },
        content: [],
      },
    ];

    const outcome = runLayoutPipeline(
      makeDeps(createLayoutSession(), {
        document,
        sectionHeaderFooterRefs: [{ footerDefault: "final-footer" }],
        renderHeaderFooterContentByRId: (_source, _hfPMs, _contentWidth, metrics) =>
          metrics.section === "footer"
            ? new Map([
                [
                  "final-footer",
                  {
                    blocks: [],
                    measures: [],
                    height: 120,
                    visualTop: 0,
                    visualBottom: 120,
                    marginPushTop: 0,
                    marginPushBottom: 120,
                  },
                ],
              ])
            : undefined,
      }),
      makeState(),
    );

    expect(outcome.layout?.pages.at(0)?.margins.bottom).toBe(168);
  });

  test("discards the outcome and leaves the session unmutated when the paint phase throws", () => {
    const session = createLayoutSession();
    const state = makeState();

    // The throw lands in the render-pages (paint) phase: `renderPages` is a
    // direct import (not injectable), so we force it to throw on its very first
    // container access. This is after the session is STAGED (measure step) but
    // before it is COMMITTED, exactly the window the hardening protects.
    const throwingContainer = asContainer(
      new Proxy(
        {},
        {
          get() {
            throw new Error("paint phase exploded");
          },
        },
      ),
    );
    const deps = makeDeps(session, {
      painter: new LayoutPainter(),
      pagesContainer: throwingContainer,
    });

    // The catch swallows the throw — the pipeline must not rethrow.
    const outcome = runLayoutPipeline(deps, state);

    // Nothing applied: the partial outcome is dropped.
    expect(outcome).toEqual({});

    // The session keeps its pre-call values, so the next run re-lays-out
    // instead of skipping on pages that were never painted.
    expect(session.artifacts).toBeNull();
    expect(session.lastEditorState).toBeNull();
    expect(session.lastPmDoc).toBeNull();
    expect(session.usedLoadedFonts).toBe(false);
    expect(session.lastTemplatePreview).toEqual({ entries: [], hidden: [], mode: "plain" });

    // The error recorder ran; no completion was recorded.
    expect(layoutErrors).toHaveLength(1);
    expect(layoutCompletes).toHaveLength(0);
  });
});

describe("createLayoutSession", () => {
  test("returns the documented empty defaults", () => {
    expect(createLayoutSession()).toEqual({
      artifacts: null,
      lastEditorState: null,
      lastPmDoc: null,
      usedLoadedFonts: false,
      lastTemplatePreview: { entries: [], hidden: [], mode: "plain" },
    });
  });
});
