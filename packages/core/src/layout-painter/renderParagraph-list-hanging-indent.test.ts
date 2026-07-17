// eigenpal #730 (#729) — a numbered list whose direct paragraph indent has
// `hanging` greater than `left` must hang its marker into the left margin (as
// Word does), not clamp it to the content edge. The marker line keeps
// `text-indent: 0` and the hang comes from padding-left; CSS padding can't be
// negative, so when `left - hanging < 0` the negative portion rides on the
// marker's own `margin-left`. Without it the old `Math.max(0, left - hanging)`
// clamp pinned the marker to the content edge, shifting the numbers right of
// the text above.

import { describe, expect, test } from "bun:test";

import { clearTextWidthCache } from "../layout-engine/measure/cache";
import { resetCanvasContext } from "../layout-engine/measure/measureContainer";
import type { ParagraphBlock, ParagraphFragment, ParagraphMeasure } from "../layout-engine/types";
import { renderParagraphFragment } from "./renderParagraph";

function createFakeStyle(): Record<string, string> {
  const store: Record<string, string> = {};
  return new Proxy(store, {
    get(target, prop: string) {
      if (prop === "setProperty") {
        return (key: string, value: string) => {
          target[key] = value;
        };
      }
      if (prop === "getPropertyValue") {
        return (key: string) => target[key] ?? "";
      }
      return target[prop];
    },
    set(target, prop: string, value: string) {
      target[prop] = value;
      return true;
    },
  }) as unknown as Record<string, string>;
}

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  innerHTML = "";
  dir = "";
  style: Record<string, string> = createFakeStyle();
  children: FakeElement[] = [];
  classList = {
    add: (...tokens: string[]) => {
      this.className = [this.className, ...tokens].filter(Boolean).join(" ");
    },
  };
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
  }

  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }
  prepend(...children: FakeElement[]): void {
    this.children.unshift(...children);
  }
  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }
  getContext(): {
    font: string;
    measureText: (text: string) => { width: number };
  } | null {
    if (this.tagName !== "canvas") {
      return null;
    }
    return { font: "", measureText: (t: string) => ({ width: t.length * 7 }) };
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
} as unknown as Document;

function renderListItem(
  indent: { left: number; hanging: number },
  listMarkerAlignment?: "left" | "center" | "right",
  listMarkerBold?: boolean,
): {
  line: HTMLElement;
  marker: HTMLElement | undefined;
} {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "p1",
    runs: [{ kind: "text", text: "TEST1" }],
    attrs: { listMarker: "1.", indent, listMarkerAlignment, listMarkerBold },
  };
  const measure: ParagraphMeasure = {
    kind: "paragraph",
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 5,
        width: 40,
        ascent: 10,
        descent: 3,
        lineHeight: 14,
      },
    ],
    totalHeight: 14,
  };
  const fragment: ParagraphFragment = {
    kind: "paragraph",
    blockId: "p1",
    x: 0,
    y: 0,
    width: 400,
    height: 14,
    fromLine: 0,
    toLine: 1,
  };

  const originalDocument = globalThis.document;
  Object.defineProperty(globalThis, "document", {
    value: fakeDocument,
    configurable: true,
  });
  clearTextWidthCache();
  resetCanvasContext();
  try {
    const fragmentEl = renderParagraphFragment(
      fragment,
      block,
      measure,
      { pageNumber: 1, totalPages: 1, section: "body" },
      { document: fakeDocument },
    );
    const line = fragmentEl.children[0] as HTMLElement;
    const marker = line.children[0] as HTMLElement | undefined;
    return { line, marker };
  } finally {
    clearTextWidthCache();
    resetCanvasContext();
    Object.defineProperty(globalThis, "document", {
      value: originalDocument,
      configurable: true,
    });
  }
}

describe("Issue #729 — list hanging indent exceeding left indent", () => {
  test("marker bold state reaches the painted element", () => {
    const bold = renderListItem({ left: 48, hanging: 24 }, undefined, true).marker;
    const regular = renderListItem({ left: 48, hanging: 24 }, undefined, false).marker;

    expect(bold?.style.fontWeight).toBe("700");
    expect(regular?.style.fontWeight).toBe("normal");
  });

  test("right-aligned marker paints before its anchor without moving body text", () => {
    const { marker } = renderListItem({ left: 48, hanging: 24 }, "right");

    expect(marker?.style.width).toBe("24px");
    expect(Number.parseFloat(marker?.style.transform?.slice(11) ?? "")).toBeCloseTo(-14, 1);
  });

  test("hanging > left: marker hangs into the margin via negative margin-left", () => {
    // 15px left, 38px hanging — marker should start at 15 - 38 = -23px.
    const { line, marker } = renderListItem({ left: 15, hanging: 38 });
    expect(marker?.className).toContain("layout-list-marker");
    expect(Number.parseFloat(marker?.style.marginLeft ?? "")).toBeCloseTo(-23, 1);
    // padding clamps to 0 (can't be negative); text-indent stays 0.
    expect(line.style.paddingLeft).toBe("0px");
    expect(line.style.textIndent).toBe("0");
  });

  test("hanging <= left: existing path unchanged (padding, no marker margin)", () => {
    // 48px left, 24px hanging — marker starts at 48 - 24 = 24px via padding.
    const { line, marker } = renderListItem({ left: 48, hanging: 24 });
    expect(marker?.style.marginLeft).toBeFalsy();
    expect(line.style.paddingLeft).toBe("24px");
    expect(line.style.textIndent).toBe("0");
  });

  test("left == 0 with hanging: marker hangs left and body stays at the content edge", () => {
    const { line, marker } = renderListItem({ left: 0, hanging: 24 });
    expect(marker?.style.marginLeft).toBe("-24px");
    expect(line.style.paddingLeft).toBe("0px");
  });

  test("left == 0 with hanging: continuation lines stay at the content edge", () => {
    const { continuation } = renderMultiLineListItem({ left: 0, hanging: 24 });
    expect(continuation.style.paddingLeft).toBeFalsy();
    expect(continuation.style.marginLeft).toBeFalsy();
  });
});

describe("negative left indent with hanging (w:ind w:left<0 w:hanging)", () => {
  test("marker hangs to indentLeft - hanging, accounting for the line's margin-left", () => {
    // `w:ind w:left="-180" w:hanging="360"` -> left = -9px, hanging = 18px.
    // Word puts the marker at left - hanging = -27px (deep in the left margin)
    // and the body at left = -9px. The line element already carries
    // margin-left = min(left, 0) = -9px, so the marker only needs the
    // remaining -18px on its own margin-left.
    const { line, marker } = renderListItem({ left: -9, hanging: 18 });
    expect(marker?.className).toContain("layout-list-marker");
    expect(Number.parseFloat(marker?.style.marginLeft ?? "")).toBeCloseTo(-18, 1);
    // padding clamps to 0 (markerStart is negative); the -9px negative left
    // indent rides on the line's own margin-left, not padding.
    expect(line.style.paddingLeft).toBe("0px");
    expect(line.style.textIndent).toBe("0");
    expect(Number.parseFloat(line.style.marginLeft ?? "")).toBeCloseTo(-9, 1);
  });

  test("continuation line sits at the negative left indent, not left + hanging", () => {
    // Guard against double-shifting: the body/continuation line must land at
    // `left` (-9px, via the line margin-left), NOT get an extra `hanging`
    // padding on top of it.
    const { continuation } = renderMultiLineListItem({ left: -9, hanging: 18 });
    // No hanging padding added; the negative indent is realized by margin-left.
    expect(continuation.style.paddingLeft).toBeFalsy();
    expect(Number.parseFloat(continuation.style.marginLeft ?? "")).toBeCloseTo(-9, 1);
  });
});

function renderMultiLineListItem(indent: { left: number; hanging: number }): {
  first: HTMLElement;
  continuation: HTMLElement;
} {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "p1",
    runs: [{ kind: "text", text: "TEST1TEST2" }],
    attrs: { listMarker: "1.", indent },
  };
  const line = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: 5,
    width: 40,
    ascent: 10,
    descent: 3,
    lineHeight: 14,
  };
  const measure: ParagraphMeasure = {
    kind: "paragraph",
    lines: [line, { ...line, fromChar: 5, toChar: 10 }],
    totalHeight: 28,
  };
  const fragment: ParagraphFragment = {
    kind: "paragraph",
    blockId: "p1",
    x: 0,
    y: 0,
    width: 400,
    height: 28,
    fromLine: 0,
    toLine: 2,
  };

  const originalDocument = globalThis.document;
  Object.defineProperty(globalThis, "document", {
    value: fakeDocument,
    configurable: true,
  });
  clearTextWidthCache();
  resetCanvasContext();
  try {
    const fragmentEl = renderParagraphFragment(
      fragment,
      block,
      measure,
      { pageNumber: 1, totalPages: 1, section: "body" },
      { document: fakeDocument },
    );
    return {
      first: fragmentEl.children[0] as HTMLElement,
      continuation: fragmentEl.children[1] as HTMLElement,
    };
  } finally {
    clearTextWidthCache();
    resetCanvasContext();
    Object.defineProperty(globalThis, "document", {
      value: originalDocument,
      configurable: true,
    });
  }
}
