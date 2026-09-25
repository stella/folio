// `w:smallCaps` paints a mixed-case run as alternating full-size and
// shrunken-capital segments (see smallCapsCasing.ts) instead of asking a
// browser for its own `font-variant: small-caps` synthesis, which scales at
// a different, uncontrollable ratio.

import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock, Run } from "../layout-engine/types";
import { renderLine } from "./renderParagraph";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  private text = "";
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

  get textContent(): string {
    return this.children.length > 0
      ? this.children.map((child) => child.textContent).join("")
      : this.text;
  }

  set textContent(value: string) {
    this.text = value;
    this.children = [];
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
  // SAFETY: renderLine uses only Document.createElement.
} as unknown as Document;

const renderSingleRun = (run: Run): FakeElement => {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "small-caps",
    runs: [run],
  };
  const line: MeasuredLine = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: run.kind === "text" ? run.text.length : 1,
    width: 100,
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

describe("small-caps painting", () => {
  test("splits mixed-case text into full-size and shrunken-capital spans", () => {
    const runElement = renderSingleRun({
      kind: "text",
      text: "Alert",
      smallCaps: true,
      fontSize: 12,
    });

    // "A" full size, "lert" synthesized at 0.8 * 16px (12pt -> 16px at 96dpi).
    expect(runElement.children.map((child) => child.textContent)).toEqual(["A", "LERT"]);
    expect(runElement.children[0]?.style["fontSize"]).toBeUndefined();
    expect(runElement.children[1]?.style["fontSize"]).toBe("12.8px");
    expect(runElement.textContent).toBe("ALERT");
  });

  test("an already-uppercase run paints as one plain span (no nested spans)", () => {
    const runElement = renderSingleRun({
      kind: "text",
      text: "HEADING",
      smallCaps: true,
      fontSize: 12,
    });

    expect(runElement.children).toHaveLength(0);
    expect(runElement.textContent).toBe("HEADING");
  });

  test("a punctuation mark between two synthesized words stays full size", () => {
    const runElement = renderSingleRun({
      kind: "text",
      text: "alpha / bravo",
      smallCaps: true,
      fontSize: 12,
    });

    expect(runElement.children.map((child) => child.textContent)).toEqual([
      "ALPHA ",
      "/ ",
      "BRAVO",
    ]);
    expect(runElement.children[0]?.style["fontSize"]).toBe("12.8px");
    expect(runElement.children[1]?.style["fontSize"]).toBeUndefined();
    expect(runElement.children[2]?.style["fontSize"]).toBe("12.8px");
  });

  test("w:caps wins over w:smallCaps: no synthesized-capital segmentation", () => {
    const runElement = renderSingleRun({
      kind: "text",
      text: "Alert",
      smallCaps: true,
      allCaps: true,
      fontSize: 12,
    });

    expect(runElement.children).toHaveLength(0);
    expect(runElement.style["textTransform"]).toBe("uppercase");
    // The DOM text keeps its authored case; `text-transform` paints the caps.
    expect(runElement.textContent).toBe("Alert");
  });

  test("a run with no smallCaps formatting is unaffected", () => {
    const runElement = renderSingleRun({ kind: "text", text: "Alert", fontSize: 12 });

    expect(runElement.children).toHaveLength(0);
    expect(runElement.textContent).toBe("Alert");
  });
});
