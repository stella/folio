import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import {
  findPositionInSpan,
  getCaretPositionFromDom,
  getSelectionRectsFromDom,
} from "./clickToPositionDom";

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

class FakeText {
  readonly nodeType = 3;
  readonly childNodes: never[] = [];

  constructor(readonly data: string) {}

  get length(): number {
    return this.data.length;
  }
}

class FakeHTMLElement {
  dataset: Record<string, string> = {};
  readonly childNodes: Array<FakeHTMLElement | FakeText> = [];
  private parent: FakeHTMLElement | null = null;
  readonly classList = { contains: (className: string) => this.classes.includes(className) };

  constructor(
    private readonly classes: string[] = [],
    private readonly elementRect: Partial<DOMRect> = {},
  ) {}

  get offsetHeight(): number {
    return this.elementRect.height ?? 0;
  }

  get firstChild(): FakeHTMLElement | FakeText | null {
    return this.childNodes.at(0) ?? null;
  }

  get ownerDocument(): { createRange: () => Range } {
    return fakeDocument;
  }

  append(...children: Array<FakeHTMLElement | FakeText>): void {
    for (const child of children) {
      if (child instanceof FakeHTMLElement) child.parent = this;
      this.childNodes.push(child);
    }
  }

  querySelectorAll(selector: string): FakeHTMLElement[] {
    const found: FakeHTMLElement[] = [];
    for (const child of this.childNodes) {
      if (!(child instanceof FakeHTMLElement)) continue;
      if (selector.includes(".layout-page-content") && child.classes.includes("layout-run-text")) {
        found.push(child);
      }
      found.push(...child.querySelectorAll(selector));
    }
    return found;
  }

  closest(selector: string): FakeHTMLElement | null {
    if (selector === ".layout-page" && this.classes.includes("layout-page")) return this;
    return this.parent?.closest(selector) ?? null;
  }

  getBoundingClientRect(): DOMRect {
    return rect(this.elementRect);
  }
}

let rangeStart: { node: FakeText; offset: number } | null;
let rangeEnd: { node: FakeText; offset: number } | null;
const fakeDocument = {
  caretPositionFromPoint: () => nativeCaret,
  createRange: () =>
    ({
      setStart: (node: FakeText, offset: number) => {
        rangeStart = { node, offset };
      },
      setEnd: (node: FakeText, offset: number) => {
        rangeEnd = { node, offset };
      },
      getClientRects: () => [rect({ left: 20, top: 30, width: 80, height: 16 })],
      getBoundingClientRect: () => rect({ left: 40, top: 30, width: 0, height: 16 }),
    }) as unknown as Range,
};

const buildNestedRun = (): {
  container: HTMLElement;
  first: FakeText;
  middle: FakeText;
  last: FakeText;
  run: HTMLElement;
} => {
  const container = new FakeHTMLElement();
  const page = new FakeHTMLElement(["layout-page"]);
  page.dataset["pageNumber"] = "2";
  const content = new FakeHTMLElement(["layout-page-content"]);
  const run = new FakeHTMLElement(["layout-run-text"]);
  run.dataset["pmStart"] = "10";
  run.dataset["pmEnd"] = "32";
  const first = new FakeText("Alpha beta ");
  const middle = new FakeText("gamma ");
  const last = new FakeText("delta");
  const firstFragment = new FakeHTMLElement();
  const middleFragment = new FakeHTMLElement();
  const nestedMiddleFragment = new FakeHTMLElement();
  const lastFragment = new FakeHTMLElement();
  firstFragment.append(first);
  nestedMiddleFragment.append(middle);
  middleFragment.append(new FakeHTMLElement(), nestedMiddleFragment);
  lastFragment.append(last);
  run.append(firstFragment, middleFragment, lastFragment);
  container.append(page);
  page.append(content);
  content.append(run);
  return {
    container: container as unknown as HTMLElement,
    first,
    middle,
    last,
    run: run as unknown as HTMLElement,
  };
};

const buildTabRun = (): HTMLElement => {
  const container = new FakeHTMLElement();
  const page = new FakeHTMLElement(["layout-page"]);
  page.dataset["pageNumber"] = "1";
  const content = new FakeHTMLElement(["layout-page-content"]);
  const run = new FakeHTMLElement(["layout-run-text", "layout-run-tab"], {
    left: 25,
    top: 35,
    width: 48,
    height: 16,
  });
  run.dataset["pmStart"] = "4";
  run.dataset["pmEnd"] = "5";
  run.append(new FakeText("\u00a0"));
  container.append(page);
  page.append(content);
  content.append(run);
  return container as unknown as HTMLElement;
};

let originalHTMLElement: unknown;
let originalNode: unknown;
let nativeCaret: { offsetNode: FakeText; offset: number } | null;

beforeEach(() => {
  rangeStart = null;
  rangeEnd = null;
  nativeCaret = null;
  originalHTMLElement = globalThis.HTMLElement;
  originalNode = globalThis.Node;
  Object.assign(globalThis, {
    HTMLElement: FakeHTMLElement,
    Node: { TEXT_NODE: 3 },
  });
});

afterEach(() => {
  Object.assign(globalThis, {
    HTMLElement: originalHTMLElement,
    Node: originalNode,
  });
});

describe("getSelectionRectsFromDom", () => {
  test("paints a run whose text is split into nested spacing fragments", () => {
    const { container, first, last } = buildNestedRun();

    const selection = getSelectionRectsFromDom(container, 10, 32, rect({}));

    expect(selection).toEqual([{ x: 20, y: 30, width: 80, height: 16, pageIndex: 1 }]);
    expect(rangeStart).toEqual({ node: first, offset: 0 });
    expect(rangeEnd).toEqual({ node: last, offset: 5 });
  });

  test("maps partial selections across nested fragment boundaries", () => {
    const { container, first, last } = buildNestedRun();

    getSelectionRectsFromDom(container, 16, 29, rect({}));

    expect(rangeStart).toEqual({ node: first, offset: 6 });
    expect(rangeEnd).toEqual({ node: last, offset: 2 });
  });

  test("retains a tab run's complete visual width", () => {
    expect(getSelectionRectsFromDom(buildTabRun(), 4, 5, rect({}))).toEqual([
      { x: 25, y: 35, width: 48, height: 16, pageIndex: 0 },
    ]);
  });

  test("maps native hit-testing in a nested fragment to the logical run offset", () => {
    const { middle, run } = buildNestedRun();
    nativeCaret = { offsetNode: middle, offset: 2 };

    expect(findPositionInSpan(run, 50, 30)).toBe(23);
  });

  test("places a caret inside a nested fragment", () => {
    const { container, middle } = buildNestedRun();

    const caret = getCaretPositionFromDom(container, 23, rect({}));

    expect(caret).toEqual({ x: 40, y: 30, height: 16, pageIndex: 1 });
    expect(rangeStart).toEqual({ node: middle, offset: 2 });
    expect(rangeEnd).toEqual({ node: middle, offset: 2 });
  });
});
