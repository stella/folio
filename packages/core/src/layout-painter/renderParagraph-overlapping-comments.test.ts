// A run can sit inside more than one comment range: Word lets ranges overlap
// and nest freely. Advertising only the first id made the run belong to one of
// them, so hover and active styling, and the sidebar anchor, answered for one
// comment and denied the other.

import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock, TextRun } from "../layout-engine/types";
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
    return {
      font: "",
      measureText(text: string) {
        return { width: text.length * 7 };
      },
    };
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
} as unknown as Document;

const paintRuns = (runs: TextRun[]): FakeElement[] => {
  const block: ParagraphBlock = { kind: "paragraph", id: "p", runs };
  const lastRun = runs.at(-1);
  if (!lastRun) {
    throw new Error("paintRuns requires at least one run");
  }
  const line: MeasuredLine = {
    fromRun: 0,
    fromChar: 0,
    toRun: runs.length - 1,
    toChar: lastRun.text.length,
    width: 200,
    ascent: 12,
    descent: 3,
    lineHeight: 15,
  };
  const lineEl = renderLine(block, line, undefined, fakeDocument, {
    availableWidth: 600,
    isLastLine: true,
    isFirstLine: true,
    paragraphEndsWithLineBreak: false,
    tabStops: [],
    leftIndentPx: 0,
    lineRightEdgePx: 600,
  }) as unknown as FakeElement;
  return lineEl.children.filter((child) => child.className.includes("layout-run-text"));
};

describe("renderParagraph comment anchors", () => {
  test("a run inside overlapping ranges advertises every comment it is in", () => {
    const [runEl] = paintRuns([{ kind: "text", text: "both", commentIds: [7, 9] }]) as [
      FakeElement,
    ];

    expect(runEl.dataset["commentIds"]).toBe("7 9");
    // The first id stays where every existing reader already looks for it.
    expect(runEl.dataset["commentId"]).toBe("7");
  });

  test("a run inside one range paints exactly what it painted before", () => {
    const [runEl] = paintRuns([{ kind: "text", text: "one", commentIds: [7] }]) as [FakeElement];

    expect(runEl.dataset).toEqual({ commentId: "7" });
  });
});
