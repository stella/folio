// Regression eigenpal #566 (renderer half): a right-aligned tab whose stop
// sits at the line's right edge with no trailing tab should promote the line
// to a flex row — tab gets `flex: 1 1 0`, trailing text/field sits flush
// against the line's right edge. Canvas-measured widths and DOM layout drift
// by sub-pixels under accumulation, so geometry alone leaves TOC page numbers
// one pixel short of the margin; flex layout pins them.

import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock, TabStop } from "../layout-engine/types";
import { renderLine } from "./renderParagraph";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  innerHTML = "";
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  classList = {
    add: (...tokens: string[]) => {
      this.className = [this.className, ...tokens].filter(Boolean).join(" ");
    },
  };
  height = 0;
  width = 0;
  src = "";
  readonly tagName: string;
  textContent = "";

  constructor(tagName: string) {
    this.tagName = tagName;
  }

  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }

  get firstElementChild(): FakeElement | null {
    return this.children.at(0) ?? null;
  }

  getContext(): {
    font: string;
    measureText: (text: string) => { width: number };
  } | null {
    if (this.tagName !== "canvas") {
      return null;
    }
    const context = {
      font: "",
      measureText(text: string) {
        const width = context.font.includes("FolioCsTestFont") ? 30 : 7;
        return { width: text.length * width };
      },
    };
    return context;
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
} as unknown as Document;

function findTabEl(lineEl: FakeElement): FakeElement | undefined {
  return lineEl.children.find((c) => c.className.includes("layout-run-tab"));
}

function findFieldOrTextEls(lineEl: FakeElement): FakeElement[] {
  return lineEl.children.filter((c) => c.className.includes("layout-run-text"));
}

describe("renderLine right-tab flex anchor", () => {
  test("tracks Arabic advance with complex-script typography before a tab", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "arabic-before-tab",
      runs: [
        {
          kind: "text",
          text: "ع",
          fontFamily: "FolioBaseTestFont",
          complexScriptFontFamily: "FolioCsTestFont",
        },
        { kind: "tab" },
        { kind: "text", text: "x" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 107,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 200,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "start", pos: 1500 }],
      leftIndentPx: 0,
      lineRightEdgePx: 200,
    }) as unknown as FakeElement;

    expect(findTabEl(lineEl)?.style["width"]).toBe("70px");
  });

  test("does NOT clamp a leading left tab to fit following text at the right edge", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "leading-left-tab",
      runs: [{ kind: "tab" }, { kind: "text", text: "wide trailing text" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 1,
      toChar: 18,
      width: 150,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 150,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "start", pos: 1500 }],
      leftIndentPx: 0,
      lineRightEdgePx: 150,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl?.style["width"]).toBe("100px");
  });

  test("clamps a non-leading left tab when tabbed label text would overshoot", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "label-left-tab",
      runs: [
        { kind: "text", text: "(a)" },
        { kind: "tab" },
        { kind: "text", text: "wide trailing text" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 18,
      width: 150,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 150,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "start", pos: 1500 }],
      leftIndentPx: 0,
      lineRightEdgePx: 150,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl?.style["width"]).toBe("3px");
  });

  test("does not clamp the tab that anchors a hanging-indent body", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "hanging-body-left-tab",
      runs: [
        { kind: "text", text: "abc" },
        { kind: "tab" },
        { kind: "text", text: "abcdefghijklmnopqr" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 18,
      width: 150,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 150,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      firstLineIndentPx: -30,
      leftIndentPx: 30,
      lineRightEdgePx: 150,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl?.style["width"]).toBe("9px");
  });

  test("does not clamp a non-leading left tab on a non-final line", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "multi-line-label-left-tab",
      runs: [
        { kind: "text", text: "(i)" },
        { kind: "tab" },
        { kind: "text", text: "wide trailing text" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 18,
      width: 150,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 150,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "start", pos: 1500 }],
      leftIndentPx: 0,
      lineRightEdgePx: 150,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl?.style["width"]).toBe("79px");
  });

  // TOC1-style line: title text + right-aligned tab + page-number field.
  // Tab stop sits at the line's right edge; with no trailing tab and a tab
  // alignment of "end", the painter must promote the line to flex layout.
  test("promotes a TOC line to flex when right-aligned tab sits at the edge", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "toc1",
      runs: [
        { kind: "text", text: "Chapter One" },
        { kind: "tab" },
        { kind: "field", fieldType: "PAGE", fallback: "5" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };
    const tabStops: TabStop[] = [
      // Right-aligned ("end") tab stop at ~600px (9000 twips = 9000/15 ≈ 600px).
      { val: "end", pos: 9000, leader: "dot" },
    ];

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 600,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops,
      leftIndentPx: 0,
      lineRightEdgePx: 600,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBe("true");
    expect(lineEl.style["display"]).toBe("flex");
    expect(lineEl.style["alignItems"]).toBe("baseline");
    // `pre` (not `nowrap`) so consecutive XML-preserved spaces in TOC titles
    // survive the flex switch; `pre` also disallows mid-line wrap.
    expect(lineEl.style["whiteSpace"]).toBe("pre");

    const tabEl = findTabEl(lineEl);
    expect(tabEl).toBeDefined();
    // flex: 1 1 0 lets the tab grow to fill the remaining space.
    expect(tabEl?.style["flex"]).toBe("1 1 0");

    // The trailing field run lands AFTER the tab in flex order so layout
    // pushes it flush right.
    const trailing = findFieldOrTextEls(lineEl);
    // Title + page number — both are text-class spans (field renders via
    // renderTextRun's "text" path with field-resolved content).
    expect(trailing.length).toBeGreaterThanOrEqual(1);
  });

  test("keeps an RTL TOC end tab logical instead of flex-anchoring it physically", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "rtl-toc",
      attrs: { bidi: true },
      runs: [{ kind: "text", text: "عنوان عربي" }, { kind: "tab" }, { kind: "text", text: "1" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 528,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      leftIndentPx: 48,
      firstLineIndentPx: -48,
      isRtl: true,
      contentWidthPx: 624,
      lineRightEdgePx: 576,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(lineEl.style["display"]).toBeUndefined();
    expect(findTabEl(lineEl)?.style["width"]).toBe("523px");
  });

  test("uses the authored tab origin when RTL swaps asymmetric physical indents", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "rtl-toc-level-2",
      attrs: { bidi: true },
      runs: [{ kind: "text", text: "عنوان عربي" }, { kind: "tab" }, { kind: "text", text: "1" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 552,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 528,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      // Physical left is zero after RTL swaps w:left=1440 to the right edge.
      leftIndentPx: 0,
      // Tab stops still use the authored logical w:left coordinate.
      tabLeftIndentPx: 96,
      firstLineIndentPx: -48,
      isRtl: true,
      contentWidthPx: 624,
      lineRightEdgePx: 528,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(findTabEl(lineEl)?.style["width"]).toBe("475px");
  });

  test("bounds an RTL end tab authored beyond the content box", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "rtl-toc-out-of-bounds",
      attrs: { bidi: true },
      runs: [{ kind: "text", text: "عنوان عربي" }, { kind: "tab" }, { kind: "text", text: "1" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 528,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 100_000_000, leader: "dot" }],
      leftIndentPx: 48,
      firstLineIndentPx: -48,
      isRtl: true,
      contentWidthPx: 624,
      lineRightEdgePx: 576,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(Number.parseFloat(findTabEl(lineEl)?.style["width"] ?? "Infinity")).toBeLessThan(624);
  });

  test("bounds an RTL end tab at a right-side floating exclusion", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "rtl-toc-right-float",
      attrs: { bidi: true },
      runs: [{ kind: "text", text: "عنوان عربي" }, { kind: "tab" }, { kind: "text", text: "1" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 276,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
      rightOffset: 300,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 228,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      leftIndentPx: 48,
      firstLineIndentPx: -48,
      isRtl: true,
      contentWidthPx: 624,
      floatingMargins: { leftMargin: 0, rightMargin: 300 },
      lineRightEdgePx: 276,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(findTabEl(lineEl)?.style["width"]).toBe("199px");
  });

  test("includes a left-side floating exclusion in the RTL tab coordinate", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "rtl-toc-left-float",
      attrs: { bidi: true },
      runs: [{ kind: "text", text: "عنوان عربي" }, { kind: "tab" }, { kind: "text", text: "1" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 300,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
      leftOffset: 300,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 228,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      leftIndentPx: 48,
      firstLineIndentPx: -48,
      isRtl: true,
      contentWidthPx: 624,
      floatingMargins: { leftMargin: 300, rightMargin: 0 },
      lineRightEdgePx: 576,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(findTabEl(lineEl)?.style["width"]).toBe("223px");
  });

  test("does NOT promote to flex when the right-aligned tab is not the last on the line", () => {
    // Two end-aligned tabs on one line — the first tab must NOT fire the
    // anchor (it has a following tab), so its trailing content lays out
    // naturally up to the next tab. Only the LAST tab is eligible for the
    // right-edge anchor.
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "two-tabs",
      runs: [
        { kind: "text", text: "A" },
        { kind: "tab" },
        { kind: "text", text: "B" },
        { kind: "tab" },
        { kind: "text", text: "C" },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 4,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };
    // Both stops are NOT at the right edge (line edge = 600px = 9000 twips).
    // The first stop sits at pos 1500 (~100px), the second at pos 3000 (~200px) —
    // currentX + tab + trailing stays well below 600, so neither tab reaches
    // the anchor's right-edge threshold even though both are end-aligned.
    const tabStops: TabStop[] = [
      { val: "end", pos: 1500, leader: "dot" },
      { val: "end", pos: 3000, leader: "dot" },
    ];

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 600,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops,
      leftIndentPx: 0,
      lineRightEdgePx: 600,
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
    expect(lineEl.style["display"]).not.toBe("flex");
  });

  test("does NOT promote to flex when no lineRightEdgePx is provided", () => {
    // Backwards-compat: callers that don't pass lineRightEdgePx (older code
    // paths, table cells before the rewrite) keep the existing non-flex
    // tab rendering. The anchor must opt in via the new option.
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "no-edge",
      runs: [{ kind: "text", text: "title" }, { kind: "tab" }, { kind: "text", text: "5" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 600,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      leftIndentPx: 0,
      // No lineRightEdgePx.
    }) as unknown as FakeElement;

    expect(lineEl.dataset["flexLine"]).toBeUndefined();
  });

  test("uses cached fallback text for non-live fields even with render context", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "cached-pageref",
      runs: [
        {
          kind: "field",
          fieldType: "OTHER",
          instruction: " PAGEREF _Toc1 \\h ",
          fallback: "2",
        },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 7,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 600,
      context: {
        pageNumber: 1,
        totalPages: 10,
        section: "body",
        bookmarkPages: new Map([["_Toc1", 6]]),
      },
    }) as unknown as FakeElement;

    expect(findFieldOrTextEls(lineEl).at(0)?.textContent).toBe("2");
  });
});

describe("renderTabRun leader rendering", () => {
  // Regression: the SVG background-image leader sat at the line's bottom
  // edge and broke under flex layout. The new pattern uses an absolutely
  // positioned inner span over a zero-width-space, so the outer span keeps
  // its baseline aligned with surrounding text while the leader clips
  // horizontally inside.
  test("renders dot leader as an absolutely-positioned inner span", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "leader-only",
      runs: [{ kind: "text", text: "x" }, { kind: "tab" }, { kind: "text", text: "y" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 600,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 600,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      tabStops: [{ val: "end", pos: 9000, leader: "dot" }],
      leftIndentPx: 0,
      lineRightEdgePx: 600,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl).toBeDefined();
    // The outer tab span is position: relative; the inner leader span is
    // position: absolute and clips horizontally.
    expect(tabEl?.style["position"]).toBe("relative");
    const inner = tabEl?.children.at(0);
    expect(inner).toBeDefined();
    expect(inner?.style["position"]).toBe("absolute");
    expect(inner?.style["overflow"]).toBe("hidden");
    expect(inner?.style["whiteSpace"]).toBe("nowrap");
    // Leader is repeated to fill the box; the exact count is an implementation
    // detail (LEADER_FILL_COUNT). What matters is that it's many dots, not one.
    const innerText = inner?.textContent ?? "";
    expect(innerText.length).toBeGreaterThan(100);
    expect(innerText.startsWith(".")).toBe(true);
  });
});
