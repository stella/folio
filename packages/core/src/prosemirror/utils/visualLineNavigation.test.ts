import { afterEach, beforeEach, describe, expect, test } from "bun:test";
import { EditorState, TextSelection } from "prosemirror-state";
import type { Transaction } from "prosemirror-state";
import type { EditorView } from "prosemirror-view";

import { schema } from "../schema";
import {
  createVisualLineState,
  findPositionOnLineAtClientX,
  getCaretClientX,
  handleVisualLineKeyDown,
} from "./visualLineNavigation";

type FakeElement = {
  closest: () => FakeElement | null;
  dataset: Record<string, string>;
  querySelector: () => FakeElement | null;
  querySelectorAll: () => FakeElement[];
};

const element = (dataset: Record<string, string> = {}): FakeElement => ({
  closest: () => null,
  dataset,
  querySelector: () => null,
  querySelectorAll: () => [],
});

describe("findPositionOnLineAtClientX", () => {
  test("uses an empty run's own position for a trailing blank line", () => {
    const paragraph = element({ pmStart: "10" });
    const emptyRun = element({ pmStart: "24", pmEnd: "25" });
    emptyRun.closest = () => paragraph;
    const line = element();
    line.querySelector = () => emptyRun;
    line.querySelectorAll = () => [emptyRun];

    expect(findPositionOnLineAtClientX(line as unknown as HTMLElement, 800)).toBe(24);
  });

  test("falls back to the paragraph content start for a structural empty line", () => {
    const paragraph = element({ pmStart: "30" });
    const emptyRun = element();
    emptyRun.closest = () => paragraph;
    const line = element();
    line.querySelector = () => emptyRun;

    expect(findPositionOnLineAtClientX(line as unknown as HTMLElement, 200)).toBe(31);
  });
});

/**
 * Painted-layout stand-in for the vertical-navigation tests: enough of the DOM
 * surface that `visualLineNavigation` reads (pm ranges, class names, rects,
 * ranges inside the text stream) to exercise the real geometry. Character
 * advance is uniform, so a client X maps to a character offset exactly.
 */
const CHAR_WIDTH = 8;
const LINE_HEIGHT = 18;
const SPACE_ADVANCE = 4;
const CONTENT_LEFT = 100;

type FakeText = { nodeType: 3; length: number; owner: FakeHTMLElement; offsetInOwner: number };

const rect = (values: Partial<DOMRect>): DOMRect =>
  ({
    x: values.left ?? 0,
    y: values.top ?? 0,
    left: values.left ?? 0,
    top: values.top ?? 0,
    right: values.right ?? 0,
    bottom: values.bottom ?? 0,
    width: values.width ?? 0,
    height: values.height ?? 0,
    toJSON: () => values,
  }) as DOMRect;

class FakeHTMLElement {
  dataset: Record<string, string> = {};
  readonly childNodes: unknown[] = [];
  readonly classList = {
    contains: (className: string) => this.classes.has(className),
  };
  private readonly classes: Set<string>;
  private readonly children: FakeHTMLElement[] = [];
  private readonly rect: Partial<DOMRect>;
  private parent: FakeHTMLElement | null = null;

  constructor(classes: string[] = [], rectInput: Partial<DOMRect> = {}) {
    this.classes = new Set(classes);
    this.rect = rectInput;
  }

  get ownerDocument(): { createRange: () => Range } {
    return fakeDocument;
  }

  get previousElementSibling(): FakeHTMLElement | null {
    const siblings = this.parent?.children ?? [];
    const index = siblings.indexOf(this);
    return index > 0 ? (siblings.at(index - 1) ?? null) : null;
  }

  get nextElementSibling(): FakeHTMLElement | null {
    const siblings = this.parent?.children ?? [];
    const index = siblings.indexOf(this);
    return index === -1 ? null : (siblings.at(index + 1) ?? null);
  }

  append(...children: FakeHTMLElement[]): void {
    for (const child of children) {
      child.parent = this;
      this.children.push(child);
      this.childNodes.push(child);
    }
  }

  appendText(text: string): void {
    this.childNodes.push({
      nodeType: 3,
      length: text.length,
      owner: this,
      offsetInOwner: 0,
    } satisfies FakeText);
  }

  querySelectorAll(selector: string): FakeHTMLElement[] {
    const out: FakeHTMLElement[] = [];
    this.collect(selector, out);
    return out;
  }

  querySelector(selector: string): FakeHTMLElement | null {
    return this.querySelectorAll(selector).at(0) ?? null;
  }

  closest(selector: string): FakeHTMLElement | null {
    return this.matches(selector) ? this : (this.parent?.closest(selector) ?? null);
  }

  getBoundingClientRect(): DOMRect {
    return rect(this.rect);
  }

  getAttribute(name: string): string | null {
    if (name === "data-collapsed-leading-spaces") {
      return this.dataset["collapsedLeadingSpaces"] ?? null;
    }
    if (name === "data-collapsed-trailing-spaces") {
      return this.dataset["collapsedTrailingSpaces"] ?? null;
    }
    return null;
  }

  private collect(selector: string, out: FakeHTMLElement[]): void {
    for (const child of this.children) {
      if (child.matches(selector)) out.push(child);
      child.collect(selector, out);
    }
  }

  private matches(selector: string): boolean {
    if (selector.endsWith("span[data-pm-start][data-pm-end]")) {
      return (
        this.classes.has("layout-run-text") &&
        this.dataset["pmStart"] !== undefined &&
        this.dataset["pmEnd"] !== undefined
      );
    }
    if (selector.endsWith(".layout-empty-run")) return this.classes.has("layout-empty-run");
    if (selector.endsWith(".layout-line")) return this.classes.has("layout-line");
    if (selector.endsWith(".layout-paragraph")) return this.classes.has("layout-paragraph");
    return false;
  }
}

/** Collapsed line-edge spaces paint at zero width, so every offset inside one
 *  measures at the span's own left edge — the browser behaviour the painted
 *  caret geometry compensates for. */
const fakeDocument = {
  createRange: (): Range => {
    let start: FakeText | null = null;
    let startOffset = 0;
    return {
      setStart: (node: unknown, offset: number) => {
        start = node as FakeText;
        startOffset = offset;
      },
      setEnd: () => undefined,
      getBoundingClientRect: () => {
        if (!start) return rect({});
        const owner = start.owner;
        const left = owner.getBoundingClientRect().left;
        const collapsed = owner.dataset["collapsedTrailingSpaces"] === "true";
        return rect({
          left: collapsed ? left : left + startOffset * CHAR_WIDTH,
          top: 0,
          height: LINE_HEIGHT,
        });
      },
    } as unknown as Range;
  },
};

type LineSpec = { text: string; pmStart: number; left: number };

/** One wrapped paragraph: each line carries its visible text plus, where the
 *  line ends in a space, the painter's collapsed zero-width trailing-space run
 *  that keeps the space addressable. */
const buildContainer = (lines: readonly LineSpec[]): HTMLElement => {
  const container = new FakeHTMLElement();
  const page = new FakeHTMLElement(["layout-page"]);
  const content = new FakeHTMLElement(["layout-page-content"]);
  const paragraph = new FakeHTMLElement(["layout-paragraph"]);
  paragraph.dataset["pmStart"] = "0";
  container.append(page);
  page.append(content);
  content.append(paragraph);

  for (const [index, spec] of lines.entries()) {
    const top = index * LINE_HEIGHT;
    const lineEl = new FakeHTMLElement(["layout-line"], {
      left: spec.left,
      top,
      right: spec.left + spec.text.length * CHAR_WIDTH,
      bottom: top + LINE_HEIGHT,
      height: LINE_HEIGHT,
    });
    paragraph.append(lineEl);

    const visible = spec.text.replace(/ +$/u, "");
    const visibleEl = new FakeHTMLElement(["layout-run-text"], {
      left: spec.left,
      top,
      right: spec.left + visible.length * CHAR_WIDTH,
      bottom: top + LINE_HEIGHT,
      width: visible.length * CHAR_WIDTH,
      height: LINE_HEIGHT,
    });
    visibleEl.dataset["pmStart"] = String(spec.pmStart);
    visibleEl.dataset["pmEnd"] = String(spec.pmStart + visible.length);
    visibleEl.appendText(visible);
    lineEl.append(visibleEl);

    const spaces = spec.text.slice(visible.length);
    if (spaces.length === 0) continue;
    const spacesLeft = spec.left + visible.length * CHAR_WIDTH;
    const spacesEl = new FakeHTMLElement(["layout-run-text"], {
      left: spacesLeft,
      top,
      right: spacesLeft,
      bottom: top + LINE_HEIGHT,
      width: 0,
      height: 0,
    });
    spacesEl.dataset["pmStart"] = String(spec.pmStart + visible.length);
    spacesEl.dataset["pmEnd"] = String(spec.pmStart + spec.text.length);
    spacesEl.dataset["collapsedTrailingSpaces"] = "true";
    spacesEl.dataset["collapsedSpaceAdvance"] = String(SPACE_ADVANCE);
    spacesEl.appendText(spaces);
    lineEl.append(spacesEl);
  }

  return container as unknown as HTMLElement;
};

/** "one two three four" wrapped after "two " and after "three ". */
const WRAPPED_LINES: readonly LineSpec[] = [
  { text: "one two ", pmStart: 1, left: CONTENT_LEFT },
  { text: "three ", pmStart: 9, left: CONTENT_LEFT },
  { text: "four", pmStart: 15, left: CONTENT_LEFT },
];
/** The caret position shared by the end of line 0 and the start of line 1. */
const WRAP_BOUNDARY = 9;

const createView = (): EditorView => {
  const doc = schema.node("doc", null, [
    schema.node("paragraph", null, [schema.text("one two three four")]),
  ]);
  const view = {
    state: EditorState.create({ doc }),
    dispatch(tr: Transaction) {
      view.state = view.state.apply(tr);
    },
  };
  return view as unknown as EditorView;
};

const placeCaret = (view: EditorView, pos: number): void => {
  view.dispatch(view.state.tr.setSelection(TextSelection.create(view.state.doc, pos)));
};

const arrow = (key: "ArrowDown" | "ArrowUp"): KeyboardEvent =>
  ({ key, ctrlKey: false, metaKey: false, shiftKey: false }) as KeyboardEvent;

let originalHTMLElement: unknown;

beforeEach(() => {
  originalHTMLElement = globalThis.HTMLElement;
  Object.assign(globalThis, { HTMLElement: FakeHTMLElement });
});

afterEach(() => {
  Object.assign(globalThis, { HTMLElement: originalHTMLElement });
});

describe("getCaretClientX", () => {
  test("measures a caret after a collapsed trailing space where the caret is painted", () => {
    const container = buildContainer(WRAPPED_LINES);
    const visibleRight = CONTENT_LEFT + "one two".length * CHAR_WIDTH;

    expect(getCaretClientX(container, WRAP_BOUNDARY)).toBe(visibleRight + SPACE_ADVANCE);
  });
});

describe("handleVisualLineKeyDown", () => {
  test("re-resolves the visual line when the caret moved since the last step", () => {
    const container = buildContainer(WRAPPED_LINES);
    const state = createVisualLineState();
    const view = createView();

    placeCaret(view, WRAP_BOUNDARY);
    expect(handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container)).toBe(true);
    const firstStep = view.state.selection.from;
    expect(firstStep).toBe(14);

    // A click (or find result, or agent edit) puts the caret back without any
    // key the handler sees: the next step must start from where the caret is.
    placeCaret(view, WRAP_BOUNDARY);
    expect(handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container)).toBe(true);

    expect(view.state.selection.from).toBe(firstStep);
  });

  test("re-measures the sticky column when the caret moved since the last step", () => {
    const container = buildContainer(WRAPPED_LINES);
    const state = createVisualLineState();
    const view = createView();

    placeCaret(view, WRAP_BOUNDARY);
    handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container);

    // Second character of line 1, far left of the column the step above kept.
    placeCaret(view, 10);
    handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container);

    expect(view.state.selection.from).toBe(16);
  });

  test("keeps stepping down when a step lands on a position two lines share", () => {
    const indented: LineSpec[] = [
      { text: "one two ", pmStart: 1, left: CONTENT_LEFT },
      { text: "three ", pmStart: 9, left: CONTENT_LEFT + 200 },
      { text: "four", pmStart: 15, left: CONTENT_LEFT },
    ];
    const container = buildContainer(indented);
    const state = createVisualLineState();
    const view = createView();

    placeCaret(view, WRAP_BOUNDARY);
    handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container);
    // The indented line starts right of the sticky column, so the step lands on
    // its first position — the same position that ends the line above.
    expect(view.state.selection.from).toBe(WRAP_BOUNDARY);

    handleVisualLineKeyDown(state, view, arrow("ArrowDown"), container);

    expect(view.state.selection.from).toBe(19);
  });
});
