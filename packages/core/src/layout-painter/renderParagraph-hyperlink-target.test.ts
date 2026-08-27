// Run hyperlink targets go through the same narrowing as image hyperlinks
// (`renderImage`): only the protocols the painter navigates to reach `href`.
// Bookmark targets (`#name`) scroll inside the document and stay verbatim.

import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock, TextRun } from "../layout-engine/types";
import { renderLine } from "./renderParagraph";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  innerHTML = "";
  style: Record<string, string> & { setProperty: (name: string, value: string) => void };
  children: FakeElement[] = [];
  classList = {
    add: (...tokens: string[]) => {
      this.className = [this.className, ...tokens].filter(Boolean).join(" ");
    },
  };
  href = "";
  target = "";
  rel = "";
  title = "";
  dir = "";
  readonly tagName: string;
  textContent = "";

  constructor(tagName: string) {
    this.tagName = tagName;
    this.style = { setProperty: () => undefined } as FakeElement["style"];
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

  getContext(): null {
    return null;
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
} as unknown as Document;

const LINE_OPTIONS = {
  availableWidth: 600,
  isLastLine: true,
  isFirstLine: true,
  paragraphEndsWithLineBreak: false,
  tabStops: [],
  leftIndentPx: 0,
  lineRightEdgePx: 600,
};

const renderRun = (href: string): { anchors: FakeElement[]; text: string } => {
  const run: TextRun = { kind: "text", text: "link", hyperlink: { href } };
  const block: ParagraphBlock = { kind: "paragraph", id: "p", runs: [run] };
  const line: MeasuredLine = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: run.text.length,
    width: 200,
    ascent: 12,
    descent: 3,
    lineHeight: 15,
  };

  const lineEl = renderLine(
    block,
    line,
    undefined,
    fakeDocument,
    LINE_OPTIONS,
  ) as unknown as FakeElement;
  const anchors: FakeElement[] = [];
  const walk = (element: FakeElement) => {
    if (element.tagName === "a") {
      anchors.push(element);
    }
    for (const child of element.children) {
      walk(child);
    }
  };
  walk(lineEl);
  const texts: string[] = [];
  const collectText = (element: FakeElement) => {
    if (element.textContent) {
      texts.push(element.textContent);
    }
    for (const child of element.children) {
      collectText(child);
    }
  };
  collectText(lineEl);
  return { anchors, text: texts.join("") };
};

const renderAnchor = (href: string): FakeElement => {
  const anchor = renderRun(href).anchors.at(0);
  if (!anchor) {
    throw new Error("expected an anchor element");
  }
  return anchor;
};

describe("run hyperlink targets", () => {
  test("keeps an http target and opens it in a new tab", () => {
    const anchor = renderAnchor("https://example.com/a?b=1#c");

    expect(anchor.href).toBe("https://example.com/a?b=1#c");
    expect(anchor.target).toBe("_blank");
    expect(anchor.rel).toBe("noopener noreferrer");
  });

  test("keeps a bookmark target verbatim and in-document", () => {
    const anchor = renderAnchor("#_Toc12345");

    expect(anchor.href).toBe("#_Toc12345");
    expect(anchor.target).toBe("");
  });

  test("renders no anchor at all for a target outside the navigable protocols", () => {
    // An empty `href` still resolves to the current document, so the run must
    // carry no anchor rather than an anchor with a blank target.
    for (const href of ["javascript:void 0", "data:text/html,<p>x</p>", "ftp://example.com/f"]) {
      const rendered = renderRun(href);

      expect(rendered.anchors).toHaveLength(0);
      expect(rendered.text).toContain("link");
    }
  });
});
