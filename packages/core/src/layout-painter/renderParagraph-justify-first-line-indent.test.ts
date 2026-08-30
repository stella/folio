// eigenpal/docx-editor#868 — a justified paragraph with a first-line indent
// must justify its first line to the FULL content width. The first-line shift
// is realized purely by `text-indent`; narrowing the justify box by `firstLine`
// as well double-counted the indent, leaving the first line's right edge short
// of the right margin while body lines reached it.

import { describe, expect, test } from "bun:test";

import { clearTextWidthCache } from "../layout-engine/measure/cache";
import { resetCanvasContext } from "../layout-engine/measure/measureContainer";
import type {
  MeasuredLine,
  ParagraphBlock,
  ParagraphFragment,
  ParagraphMeasure,
} from "../layout-engine/types";
import { renderLine, renderParagraphFragment } from "./renderParagraph";

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
  textContent = "";
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

function paintedVisualSpaceContraction(element: FakeElement, ancestorScale = 1): number {
  const transformScale = /^scaleX\(([^)]+)\)$/u.exec(element.style.transform ?? "")?.at(1);
  const effectiveScale = ancestorScale * (transformScale ? Number(transformScale) : 1);
  const wordSpacing = Number.parseFloat(element.style.wordSpacing ?? "0");
  const asciiSpaces = [...element.textContent].filter((character) => character === " ").length;
  let contraction = -wordSpacing * asciiSpaces * effectiveScale;
  for (const child of element.children) {
    contraction += paintedVisualSpaceContraction(child, effectiveScale);
  }
  return contraction;
}

function renderJustifiedFirstLine(firstLine: number): {
  firstLineEl: HTMLElement;
} {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "p1",
    runs: [{ kind: "text", text: "first second third fourth fifth sixth" }],
    attrs: { alignment: "justify", indent: { firstLine } },
  };
  const line = (toChar: number, fromChar: number, width: number) => ({
    fromRun: 0,
    fromChar,
    toRun: 0,
    toChar,
    width,
    ascent: 10,
    descent: 3,
    lineHeight: 14,
  });
  const measure: ParagraphMeasure = {
    kind: "paragraph",
    // Two lines: the first is justified, the second is the (left-aligned) last.
    // The first line's measured capacity is 400 - 30 = 370px. The painter
    // keeps the outer line box at 400px and lets text-indent reserve those
    // 30px; the measured content itself cannot occupy the full 400px.
    lines: [line(18, 0, 370), line(36, 18, 380)],
    totalHeight: 28,
  };
  const fragment: ParagraphFragment = {
    kind: "paragraph",
    blockId: "p1",
    x: 0,
    y: 0,
    width: 400, // no left/right indent → availableWidth === 400
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
    return { firstLineEl: fragmentEl.children[0] as HTMLElement };
  } finally {
    clearTextWidthCache();
    resetCanvasContext();
    Object.defineProperty(globalThis, "document", {
      value: originalDocument,
      configurable: true,
    });
  }
}

describe("Issue #868 — justify first line to full content width on indented paragraphs", () => {
  test("first line justify box is the full content width, not narrowed by firstLine", () => {
    const { firstLineEl } = renderJustifiedFirstLine(30);
    // availableWidth is the full 400px content width; the first-line shift is
    // applied via text-indent, NOT by shrinking the justify box to 370px.
    expect(firstLineEl.style.width).toBe("400px");
    expect(firstLineEl.style.textIndent).toBe("30px");
    expect(firstLineEl.style.textAlign).toBe("justify");
  });

  test("overfull justified lines use negative word spacing instead of browser expansion", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-overfull",
      runs: [{ kind: "text", text: "alpha beta gamma" }],
      attrs: { alignment: "justify" },
    };
    const line = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 16,
      width: 110,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
    });

    expect(lineEl.style.width).toBe("100px");
    expect(lineEl.style.textAlign).toBe("left");
    expect(lineEl.style.textAlignLast).toBe("auto");
    expect(lineEl.style.wordSpacing).toBe("0");
    expect((lineEl as unknown as FakeElement).children.at(0)?.style.wordSpacing).toBe("-5px");
  });

  test("compresses an overfull final justified line without expanding an underfull one", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-contraction",
      runs: [{ kind: "text", text: "alpha beta gamma" }],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line = (width: number, admission: "admitted" | "unadmitted"): MeasuredLine => {
      const measuredLine = {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 16,
        width,
        ascent: 10,
        descent: 3,
        lineHeight: 14,
      } satisfies MeasuredLine;
      if (admission === "unadmitted") {
        return measuredLine;
      }
      return {
        ...measuredLine,
        justificationPaint: { type: "space-contraction", contractionPx: 2 },
      };
    };
    const options = {
      availableWidth: 100,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    };

    const overfull = renderLine(block, line(102, "admitted"), "justify", fakeDocument, options);
    const unadmitted = renderLine(block, line(103, "unadmitted"), "justify", fakeDocument, options);
    const underfull = renderLine(block, line(90, "admitted"), "justify", fakeDocument, options);

    expect(overfull.style.width).toBe("100px");
    expect(overfull.style.wordSpacing).toBe("0");
    expect((overfull as unknown as FakeElement).children.at(0)?.style.wordSpacing).toBe("-1px");
    expect(unadmitted.style.width).toBeUndefined();
    expect(unadmitted.style.wordSpacing).toBeUndefined();
    expect(underfull.style.width).toBeUndefined();
    expect(underfull.style.wordSpacing).toBeUndefined();
  });

  test("contracts only measured ASCII spaces when the final line also contains NBSP", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-fixed-space",
      runs: [{ kind: "text", text: "alpha beta\u00a0gamma" }],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 16,
      width: 102,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 2 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    });

    expect(lineEl.style.wordSpacing).toBe("0");
    const paintedRun = (lineEl as unknown as FakeElement).children.at(0);
    expect(paintedRun?.children.at(0)?.style.wordSpacing).toBe("-2px");
    expect(paintedRun?.children.at(1)?.textContent).toBe("\u00a0");
    expect(paintedRun?.children.at(1)?.style.wordSpacing).toBe("0");
  });

  test("contracts beside hanging CJK punctuation without pulling the hanging glyph inward", () => {
    const text = `${"a ".repeat(50)}bb。`;
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-cjk-hanging",
      runs: [{ kind: "text", text, language: { eastAsia: "zh-CN" } }],
      attrs: { alignment: "justify", listMarker: "1.", overflowPunctuation: true },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: text.length,
      width: 103,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 2 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    });

    expect(lineEl.style.wordSpacing).toBe("0");
    expect((lineEl as unknown as FakeElement).children.at(0)?.style.wordSpacing).toBe("-0.04px");
  });

  test("compensates final-list space contraction inside 50% scaled text", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-half-scale",
      runs: [{ kind: "text", text: "alpha beta gamma", horizontalScale: 50 }],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 16,
      width: 204,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 4 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 200,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    }) as unknown as FakeElement;

    expect(lineEl.children.at(0)?.style.wordSpacing).toBe("-4px");
    expect(paintedVisualSpaceContraction(lineEl)).toBeCloseTo(4);
  });

  test("compensates final-list space contraction through a 150% scaled field wrapper", () => {
    const fieldText = "甲 alpha beta";
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-field-one-half-scale",
      runs: [
        {
          kind: "field",
          fieldType: "OTHER",
          instruction: "REF target",
          fallback: fieldText,
          pmStart: 1,
          horizontalScale: 150,
          eastAsiaFontFamily: "Noto Sans CJK SC",
        },
      ],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 204,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 4 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 200,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
      context: {
        pageNumber: 1,
        totalPages: 1,
        section: "body",
        bookmarkText: new Map([["target", fieldText]]),
      },
    }) as unknown as FakeElement;
    const fieldWrapper = lineEl.children.at(0);
    const latinSegment = fieldWrapper?.children.at(1);

    expect(fieldWrapper?.style.transform).toBe("scaleX(1.5)");
    expect(fieldWrapper?.style.width).toBe("122px");
    expect(latinSegment?.style.wordSpacing).toBe("-1.3333333333333333px");
    expect(paintedVisualSpaceContraction(lineEl)).toBeCloseTo(4);
  });

  test("allocates mixed-scale final-list contraction in visual coordinates", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-mixed-scale",
      runs: [
        { kind: "text", text: "a b", horizontalScale: 50 },
        { kind: "text", text: "c d", horizontalScale: 150 },
        { kind: "text", text: "e f" },
      ],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 3,
      width: 306,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 6 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 300,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    }) as unknown as FakeElement;

    expect(lineEl.children.map((child) => child.style.wordSpacing)).toEqual([
      "-4px",
      "-1.3333333333333333px",
      "-2px",
    ]);
    expect(lineEl.children.map((child) => child.style.width)).toEqual([
      "8.5px",
      "29.5px",
      undefined,
    ]);
    expect(paintedVisualSpaceContraction(lineEl)).toBeCloseTo(6);
  });

  test("uses contracted scaled-run advance when resolving a following tab", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-scaled-tab-advance",
      runs: [
        { kind: "text", text: "a b", horizontalScale: 50 },
        { kind: "tab" },
        { kind: "text", text: "c d" },
      ],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 3,
      width: 104,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 4 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    }) as unknown as FakeElement;
    const scaledRun = lineEl.children.at(0);
    const tab = lineEl.children.find((child) => child.className.includes("layout-run-tab"));

    expect(scaledRun?.style.width).toBe("8.5px");
    expect(tab?.style.width).toBe("39.5px");
  });

  test("excludes zero-scale spaces from final-list contraction", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-zero-scale",
      runs: [
        { kind: "text", text: "a b", horizontalScale: 0 },
        { kind: "text", text: "c d" },
        {
          kind: "field",
          fieldType: "OTHER",
          fallback: "甲 e f",
          eastAsiaFontFamily: "Noto Sans CJK SC",
          horizontalScale: 0,
        },
      ],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 1,
      width: 204,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 4 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 200,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
      context: { pageNumber: 1, totalPages: 1, section: "body" },
    }) as unknown as FakeElement;

    expect(lineEl.children.at(0)?.style.transform).toBe("scaleX(0)");
    expect(lineEl.children.at(0)?.style.wordSpacing).toBeUndefined();
    expect(lineEl.children.at(1)?.style.wordSpacing).toBe("-4px");
    expect(lineEl.children.at(2)?.style.transform).toBe("scaleX(0)");
    expect(paintedVisualSpaceContraction(lineEl)).toBeCloseTo(4);
  });

  test("keeps MathML whitespace neutral while final text opts into contraction", () => {
    const finalText = " alpha beta";
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-final-line-math-neutral-spacing",
      runs: [
        {
          kind: "math",
          display: "inline",
          ommlXml:
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>m n\u00a0o\u2003p</m:t></m:r></m:oMath>',
          plainText: "m n\u00a0o\u2003p",
        },
        { kind: "text", text: finalText },
      ],
      attrs: { alignment: "justify", listMarker: "1." },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 1,
      toChar: finalText.length,
      width: 102,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
      justificationPaint: { type: "space-contraction", contractionPx: 2 },
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
    });
    const children = (lineEl as unknown as FakeElement).children;

    expect(lineEl.style.wordSpacing).toBe("0");
    expect(children.at(0)?.style.wordSpacing).not.toBe("-1px");
    expect(children.at(1)?.style.wordSpacing).toBe("-1px");
  });

  test("counts resolved field text when compressing an overfull line", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-resolved-field-spaces",
      runs: [
        {
          kind: "field",
          fieldType: "OTHER",
          instruction: "REF target",
          fallback: "fallback",
          pmStart: 1,
        },
      ],
      attrs: { alignment: "justify" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 110,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      context: {
        pageNumber: 1,
        totalPages: 1,
        section: "body",
        bookmarkText: new Map([["target", "alpha beta gamma"]]),
      },
    });

    expect(lineEl.style.wordSpacing).toBe("0");
    expect((lineEl as unknown as FakeElement).children.at(0)?.style.wordSpacing).toBe("-5px");
  });

  test("does not treat native math spaces as compressible line spaces", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-native-math-spaces",
      runs: [
        {
          kind: "math",
          display: "inline",
          ommlXml:
            '<m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math"><m:r><m:t>a b</m:t></m:r></m:oMath>',
          plainText: "a b",
        },
      ],
      attrs: { alignment: "justify" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 110,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
    });

    expect(lineEl.style.wordSpacing).toBeUndefined();
  });

  test("underfull justified tab lines distribute their remaining width explicitly", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-underfull-tab",
      runs: [
        { kind: "text", text: "6.3.2" },
        { kind: "tab" },
        { kind: "text", text: "alpha beta" },
      ],
      attrs: { alignment: "justify" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 10,
      width: 80,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
    });

    expect(lineEl.style.width).toBe("100px");
    expect(lineEl.style.textAlign).toBe("left");
    expect(lineEl.style.textAlignLast).toBe("auto");
    expect(lineEl.style.wordSpacing).toBe("20px");
  });

  test("does not distribute space into a positive first-line indent", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-first-line-tab-capacity",
      runs: [{ kind: "tab" }, { kind: "text", text: "alpha beta" }],
      attrs: { alignment: "justify", indent: { firstLine: 30 } },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 1,
      toChar: 10,
      // The measurer subtracts the 30px first-line shift from the 100px
      // paragraph width. This line already fills that complete 70px budget.
      width: 70,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      firstLineIndentPx: 30,
    });

    // The line box stays at the full paragraph width so CSS text-indent can
    // place the first line. Explicit word spacing must use the remaining 70px
    // capacity, or it expands the text by the indent and overruns the margin.
    expect(lineEl.style.width).toBe("100px");
    expect(lineEl.style.wordSpacing).toBeUndefined();
    expect(lineEl.style.textAlign).toBe("justify");
  });

  test("includes a hanging first-line region in justified tab capacity", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-underfull-hanging-tab",
      runs: [
        { kind: "text", text: "6.3.2" },
        { kind: "tab" },
        { kind: "text", text: "alpha beta" },
      ],
      attrs: { alignment: "justify", indent: { left: 20, hanging: 20 } },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 10,
      width: 110,
      ascent: 10,
      descent: 3,
      lineHeight: 14,
    };

    const lineEl = renderLine(block, line, "justify", fakeDocument, {
      availableWidth: 100,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      firstLineIndentPx: -20,
    });

    expect(lineEl.style.width).toBe("100px");
    expect(lineEl.style.wordSpacing).toBe("10px");
  });

  test("justifies a hanging-list marker line through the full right edge", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-list-justify-width",
      runs: [{ kind: "text", text: "Item text" }],
      attrs: { alignment: "justify", listMarker: "1.", indent: { left: 36, hanging: 36 } },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 9,
      width: 90,
      ascent: 10,
      descent: 2,
      lineHeight: 12,
    };

    const originalDocument = globalThis.document;
    Object.defineProperty(globalThis, "document", { value: fakeDocument, configurable: true });
    resetCanvasContext();
    try {
      const lineEl = renderLine(block, line, "justify", fakeDocument, {
        availableWidth: 100,
        isLastLine: false,
        isFirstLine: true,
        paragraphEndsWithLineBreak: false,
        firstLineIndentPx: -36,
        leftIndentPx: 36,
      }) as unknown as FakeElement;

      expect(lineEl.style["width"]).toBe("136px");
    } finally {
      resetCanvasContext();
      Object.defineProperty(globalThis, "document", {
        value: originalDocument,
        configurable: true,
      });
    }
  });

  test("does not widen a zero-left hanging-list line past the right edge", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "p-zero-left-list-justify-width",
      runs: [{ kind: "text", text: "Item text" }],
      attrs: { alignment: "justify", listMarker: "1.", indent: { left: 0, hanging: 36 } },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 9,
      width: 90,
      ascent: 10,
      descent: 2,
      lineHeight: 12,
    };

    const originalDocument = globalThis.document;
    Object.defineProperty(globalThis, "document", { value: fakeDocument, configurable: true });
    resetCanvasContext();
    try {
      const lineEl = renderLine(block, line, "justify", fakeDocument, {
        availableWidth: 100,
        isLastLine: false,
        isFirstLine: true,
        paragraphEndsWithLineBreak: false,
        firstLineIndentPx: -36,
        leftIndentPx: 0,
      }) as unknown as FakeElement;

      expect(lineEl.style["width"]).toBe("100px");
    } finally {
      resetCanvasContext();
      Object.defineProperty(globalThis, "document", {
        value: originalDocument,
        configurable: true,
      });
    }
  });
});
