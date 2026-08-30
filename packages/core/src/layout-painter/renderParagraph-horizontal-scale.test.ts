import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock, Run } from "../layout-engine/types";
import { renderLine } from "./renderParagraph";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  textContent = "";
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

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }

  getContext(): { font: string; measureText: (text: string) => { width: number } } | null {
    if (this.tagName !== "canvas") {
      return null;
    }
    return {
      font: "",
      measureText: (text: string) => ({ width: text.length * 7 }),
    };
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
  // SAFETY: renderSingleRun uses only Document.createElement.
} as unknown as Document;

const renderSingleRun = (run: Run): FakeElement => {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "horizontal-scale",
    runs: [run],
  };
  const line: MeasuredLine = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: run.kind === "text" ? run.text.length : 1,
    width: 28,
    ascent: 10,
    descent: 2,
    lineHeight: 12,
  };

  const lineElement = renderLine(block, line, undefined, fakeDocument, {
    availableWidth: 360,
    context: { pageNumber: 1, totalPages: 1, section: "body" },
  }) as unknown as FakeElement;
  const runElement = lineElement.children.at(0);
  if (!runElement) {
    throw new Error("Expected painted run");
  }
  return runElement;
};

describe("horizontal text scale painting", () => {
  test.each([
    [0, "scaleX(0)", "0px"],
    [600, "scaleX(6)", "168px"],
  ])("preserves valid boundary scale %p", (horizontalScale, transform, width) => {
    const runElement = renderSingleRun({ kind: "text", text: "abcd", horizontalScale });

    expect(runElement.style["transform"]).toBe(transform);
    expect(runElement.style["width"]).toBe(width);
  });

  test.each([-50, 601, Number.NaN, Number.POSITIVE_INFINITY, Number.NEGATIVE_INFINITY])(
    "does not emit CSS for malformed scale %p",
    (horizontalScale) => {
      const runElement = renderSingleRun({ kind: "text", text: "abcd", horizontalScale });

      expect(runElement.style["transform"]).toBeUndefined();
      expect(runElement.style["width"]).toBeUndefined();
    },
  );

  test("normalizes the field wrapper and reserved advance together", () => {
    const runElement = renderSingleRun({
      kind: "field",
      fieldType: "OTHER",
      fallback: "A合",
      eastAsiaFontFamily: "Noto Sans CJK",
      horizontalScale: 0,
    });

    expect(runElement.style["transform"]).toBe("scaleX(0)");
    expect(runElement.style["width"]).toBe("0px");
  });
});
