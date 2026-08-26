import { describe, expect, test } from "bun:test";

import type { FolioAIEditSnapshot } from "../ai-edits/types";
import { onPaintedLayoutChange, readBlockRects } from "./blockGeometry";

type RectInit = {
  top: number;
  left: number;
  width: number;
  height: number;
};

const domRect = ({ top, left, width, height }: RectInit): DOMRect => ({
  x: left,
  y: top,
  top,
  left,
  width,
  height,
  right: left + width,
  bottom: top + height,
  toJSON: () => ({}),
});

class FakeElement {
  readonly children: FakeElement[] = [];
  readonly dataset: Record<string, string> = {};
  readonly listeners = new Map<string, Set<() => void>>();
  clientLeft = 0;
  clientTop = 0;
  rectReads = 0;
  scrollLeft = 0;
  scrollTop = 0;

  constructor(private readonly rect: DOMRect) {}

  append(...elements: FakeElement[]): void {
    this.children.push(...elements);
  }

  addEventListener(type: string, listener: () => void): void {
    const listeners = this.listeners.get(type);
    if (listeners) {
      listeners.add(listener);
    } else {
      this.listeners.set(type, new Set([listener]));
    }
  }

  removeEventListener(type: string, listener: () => void): void {
    this.listeners.get(type)?.delete(listener);
  }

  dispatch(type: string): void {
    for (const listener of this.listeners.get(type) ?? []) {
      listener();
    }
  }

  getBoundingClientRect(): DOMRect {
    this.rectReads += 1;
    return this.rect;
  }

  querySelectorAll(selector: string): FakeElement[] {
    const matches = (element: FakeElement): boolean => {
      if (selector === "[data-page-number]") {
        return element.dataset["pageNumber"] !== undefined;
      }
      if (selector === "[data-block-id][data-pm-start]") {
        return element.dataset["blockId"] !== undefined && element.dataset["pmStart"] !== undefined;
      }
      return false;
    };
    const found: FakeElement[] = [];
    const visit = (element: FakeElement): void => {
      if (matches(element)) {
        found.push(element);
      }
      for (const child of element.children) {
        visit(child);
      }
    };
    for (const child of this.children) {
      visit(child);
    }
    return found;
  }
}

const block = (pmStart: number, rect: RectInit): FakeElement => {
  const element = new FakeElement(domRect(rect));
  element.dataset["blockId"] = `renderer-${String(pmStart)}`;
  element.dataset["pmStart"] = String(pmStart);
  return element;
};

const page = (pageNumber: number, blocks: readonly FakeElement[]): FakeElement => {
  const element = new FakeElement(domRect({ top: 0, left: 0, width: 600, height: 800 }));
  element.dataset["pageNumber"] = String(pageNumber);
  element.append(...blocks);
  return element;
};

const snapshot = (...entries: readonly [blockId: string, from: number][]): FolioAIEditSnapshot => ({
  blocks: entries.map(([id]) => ({ id, kind: "paragraph", text: id })),
  anchors: Object.fromEntries(
    entries.map(([id, from]) => [
      id,
      {
        id,
        from,
        to: from + 1,
        text: id,
        normalizedText: id,
        textHash: `hash-${id}`,
        hashOccurrenceCount: 1,
      },
    ]),
  ),
});

const htmlElement = (element: FakeElement): HTMLElement => {
  // SAFETY: the geometry reader uses only the DOM members implemented above.
  return element as unknown as HTMLElement;
};

describe("block geometry", () => {
  test("measures snapshot ids in order and omits an unpainted page", () => {
    const first = block(10, { top: 140, left: 90, width: 300, height: 20 });
    const second = block(20, { top: 180, left: 90, width: 320, height: 24 });
    const pages = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    pages.append(page(1, [first, second]), page(2, []));
    const scrollRoot = new FakeElement(domRect({ top: 100, left: 50, width: 700, height: 500 }));
    scrollRoot.clientLeft = 2;
    scrollRoot.clientTop = 3;
    scrollRoot.scrollLeft = 5;
    scrollRoot.scrollTop = 200;

    const rects = readBlockRects({
      blockIds: ["first", "second", "unpainted", "missing"],
      snapshot: snapshot(["first", 10], ["second", 20], ["unpainted", 30]),
      pagesContainer: htmlElement(pages),
      scrollRoot: htmlElement(scrollRoot),
    });

    expect(rects.get("first")).toEqual({
      blockId: "first",
      page: 1,
      top: 237,
      left: 43,
      width: 300,
      height: 20,
    });
    expect(rects.get("second")?.top).toBeGreaterThan(rects.get("first")?.top ?? Infinity);
    expect(rects.has("unpainted")).toBe(false);
    expect(rects.has("missing")).toBe(false);
    expect(scrollRoot.rectReads).toBe(1);
    expect(first.rectReads).toBe(1);
    expect(second.rectReads).toBe(1);
  });

  test("resolves a positional seq id the way scrollToBlock does", () => {
    // A paraId-less paragraph is cited by the server as `seq-NNNN`, while the
    // live snapshot keys the same block by the hex paraId the editor minted.
    const first = block(10, { top: 140, left: 90, width: 300, height: 20 });
    const second = block(20, { top: 180, left: 90, width: 320, height: 24 });
    const pages = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    pages.append(page(1, [first, second]));
    const scrollRoot = new FakeElement(domRect({ top: 0, left: 0, width: 700, height: 500 }));

    const rects = readBlockRects({
      blockIds: ["seq-0002", "seq-0009"],
      snapshot: snapshot(["1a2b3c4d", 10], ["5e6f7a8b", 20]),
      pagesContainer: htmlElement(pages),
      scrollRoot: htmlElement(scrollRoot),
    });

    expect(rects.get("seq-0002")).toMatchObject({
      blockId: "seq-0002",
      page: 1,
      top: 180,
      width: 320,
      height: 24,
    });
    expect(rects.has("seq-0009")).toBe(false);
  });

  test("never reads a matching block from another editor instance", () => {
    const blockA = block(10, { top: 120, left: 20, width: 100, height: 10 });
    const blockB = block(10, { top: 420, left: 20, width: 200, height: 30 });
    const pagesA = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    const pagesB = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    pagesA.append(page(1, [blockA]));
    pagesB.append(page(1, [blockB]));
    const rootA = new FakeElement(domRect({ top: 100, left: 0, width: 500, height: 400 }));
    const rootB = new FakeElement(domRect({ top: 400, left: 0, width: 500, height: 400 }));

    const rectA = readBlockRects({
      blockIds: ["same-id"],
      snapshot: snapshot(["same-id", 10]),
      pagesContainer: htmlElement(pagesA),
      scrollRoot: htmlElement(rootA),
    }).get("same-id");
    const rectB = readBlockRects({
      blockIds: ["same-id"],
      snapshot: snapshot(["same-id", 10]),
      pagesContainer: htmlElement(pagesB),
      scrollRoot: htmlElement(rootB),
    }).get("same-id");

    expect(rectA).toMatchObject({ top: 20, width: 100, height: 10 });
    expect(rectB).toMatchObject({ top: 20, width: 200, height: 30 });
  });

  test("layout subscriptions are instance-scoped and disposable", () => {
    const pagesA = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    const pagesB = new FakeElement(domRect({ top: 0, left: 0, width: 0, height: 0 }));
    let changes = 0;
    const unsubscribe = onPaintedLayoutChange(htmlElement(pagesA), () => {
      changes += 1;
    });

    pagesB.dispatch("painter:painted");
    expect(changes).toBe(0);
    pagesA.dispatch("painter:painted");
    expect(changes).toBe(1);
    unsubscribe();
    pagesA.dispatch("painter:painted");
    expect(changes).toBe(1);
  });
});
