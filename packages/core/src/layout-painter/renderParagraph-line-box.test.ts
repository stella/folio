import { describe, expect, test } from "bun:test";

import type { MeasuredLine, ParagraphBlock } from "../layout-engine/types";
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

function findTabEl(lineEl: FakeElement): FakeElement | undefined {
  return lineEl.children.find((child) => child.className.includes("layout-run-tab"));
}

function findTabEls(lineEl: FakeElement): FakeElement[] {
  return lineEl.children.filter((child) => child.className.includes("layout-run-tab"));
}

describe("renderLine box model", () => {
  test("does not create a text measurer for an ordinary unscaled line", () => {
    let canvasCreations = 0;
    const countingDocument = {
      createElement(tagName: string): FakeElement {
        if (tagName === "canvas") {
          canvasCreations += 1;
        }
        return new FakeElement(tagName);
      },
    } as unknown as Document;
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "unscaled-line",
      runs: [{ kind: "text", text: "No tabs or scaling" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "No tabs or scaling".length,
      width: 112,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, countingDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;

    expect(lineEl.children.at(0)?.textContent).toBe("No tabs or scaling");
    expect(canvasCreations).toBe(0);
  });

  test("measures a horizontally scaled run to reserve its painted advance", () => {
    let canvasCreations = 0;
    const countingDocument = {
      createElement(tagName: string): FakeElement {
        if (tagName === "canvas") {
          canvasCreations += 1;
        }
        return new FakeElement(tagName);
      },
    } as unknown as Document;
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "scaled-line",
      runs: [{ kind: "text", text: "scaled", horizontalScale: 50 }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "scaled".length,
      width: 21,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, countingDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;

    expect(lineEl.children.at(0)?.style["width"]).toBe("21px");
    expect(canvasCreations).toBe(1);
  });

  test("paints no-break hyphens without changing document positions", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "no-break-hyphen",
      runs: [{ kind: "text", text: "non\u2011breaking", pmStart: 10, pmEnd: 22 }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "non\u2011breaking".length,
      width: 84,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 84,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const text = lineEl.children.at(0);

    expect(text?.textContent).toBe("non-breaking");
    expect(text?.dataset["pmStart"]).toBe("10");
    expect(text?.dataset["pmEnd"]).toBe("22");
  });

  test("paints a discretionary hyphen without assigning it a document position", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "discretionary-hyphen",
      runs: [
        {
          kind: "text",
          text: "hyphenation",
          pmStart: 10,
          pmEnd: 21,
          bold: true,
        },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 6,
      width: 49,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
      discretionaryHyphen: { runIndex: 0 },
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 49,
      isLastLine: false,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [text, hyphen] = lineEl.children;

    expect(text?.textContent).toBe("hyphen");
    expect(text?.dataset["pmStart"]).toBe("10");
    expect(text?.dataset["pmEnd"]).toBe("16");
    expect(hyphen?.textContent).toBe("-");
    expect(hyphen?.dataset["discretionaryHyphen"]).toBe("true");
    expect(hyphen?.dataset["pmStart"]).toBeUndefined();
    expect(hyphen?.dataset["pmEnd"]).toBeUndefined();
    expect(hyphen?.style["fontWeight"]).toBe("700");
  });

  test("paints pair kerning only when the authored threshold is met", () => {
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 4,
      width: 28,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };
    const render = (kerningMinPt?: number) =>
      renderLine(
        {
          kind: "paragraph",
          id: "kerning",
          runs: [
            {
              kind: "text",
              text: "AVAV",
              fontSize: 11,
              ...(kerningMinPt !== undefined ? { kerningMinPt } : {}),
            },
          ],
        },
        line,
        undefined,
        fakeDocument,
        {
          availableWidth: 360,
          isLastLine: true,
          isFirstLine: true,
          paragraphEndsWithLineBreak: false,
          leftIndentPx: 0,
        },
      ) as unknown as FakeElement;

    expect(render().children.at(0)?.style["fontKerning"]).toBe("none");
    expect(render(12).children.at(0)?.style["fontKerning"]).toBe("none");
    expect(render(10).children.at(0)?.style["fontKerning"]).toBe("normal");
  });

  test("uses content-box and visible overflow so highlighted text is not clipped", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "highlighted-placeholder",
      runs: [
        {
          kind: "text",
          text: "COMPANY NAME",
          highlight: "yellow",
        },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "COMPANY NAME".length,
      width: 100,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 288,
    }) as unknown as FakeElement;

    expect(lineEl.style["boxSizing"]).toBe("content-box");
    expect(lineEl.style["overflow"]).toBe("visible");
  });

  test("collapses paintless trailing spaces without losing their document positions", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "right-aligned-reference",
      runs: [{ kind: "text", text: "Reference   ", pmStart: 10, pmEnd: 22 }],
      attrs: { alignment: "right" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "Reference   ".length,
      width: 63,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [text, spaces] = lineEl.children;

    expect(text?.textContent).toBe("Reference");
    expect(text?.dataset["pmStart"]).toBe("10");
    expect(text?.dataset["pmEnd"]).toBe("19");
    expect(spaces?.textContent).toBe("   ");
    expect(spaces?.dataset["pmStart"]).toBe("19");
    expect(spaces?.dataset["pmEnd"]).toBe("22");
    expect(spaces?.dataset["collapsedTrailingSpaces"]).toBe("true");
    expect(spaces?.style["fontSize"]).toBe("0");
  });

  test("collapses trailing spaces split across authored runs", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "split-trailing-spaces",
      runs: [
        { kind: "text", text: "Reference", pmStart: 10, pmEnd: 19 },
        { kind: "text", text: " ", pmStart: 19, pmEnd: 20 },
        { kind: "text", text: "  ", pmStart: 20, pmEnd: 22 },
      ],
      attrs: { alignment: "right" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 2,
      toChar: 2,
      width: 63,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;

    expect(lineEl.children.map((child) => child.textContent)).toEqual(["Reference", " ", "  "]);
    expect(lineEl.children.map((child) => child.style["fontSize"])).toEqual([undefined, "0", "0"]);
  });

  test("keeps decorated trailing spaces paintable", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "underlined-signature",
      runs: [{ kind: "text", text: "Signature   ", underline: true }],
      attrs: { alignment: "right" },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "Signature   ".length,
      width: 84,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, "right", fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [text] = lineEl.children;

    expect(lineEl.children).toHaveLength(1);
    expect(text?.textContent).toBe("Signature   ");
    expect(text?.dataset["collapsedTrailingSpaces"]).toBeUndefined();
    expect(text?.style["fontSize"]).toBeUndefined();
  });

  test("collapses ordinary spaces at the start of a soft-wrapped line", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "soft-wrapped-leading-spaces",
      runs: [{ kind: "text", text: "alpha   Reference", pmStart: 10, pmEnd: 27 }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: "alpha".length,
      toRun: 0,
      toChar: "alpha   Reference".length,
      width: 63,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [spaces, text] = lineEl.children;

    expect(spaces?.textContent).toBe("   ");
    expect(spaces?.dataset["pmStart"]).toBe("15");
    expect(spaces?.dataset["pmEnd"]).toBe("18");
    expect(spaces?.dataset["collapsedLeadingSpaces"]).toBe("true");
    expect(spaces?.style["fontSize"]).toBe("0");
    expect(spaces?.style["letterSpacing"]).toBe("0");
    expect(spaces?.style["wordSpacing"]).toBe("0");
    expect(text?.textContent).toBe("Reference");
    expect(text?.dataset["pmStart"]).toBe("18");
    expect(text?.dataset["pmEnd"]).toBe("27");
  });

  test("keeps authored paragraph-leading spaces paintable", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "authored-leading-spaces",
      runs: [{ kind: "text", text: "   Reference" }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: "   Reference".length,
      width: 84,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [text] = lineEl.children;

    expect(lineEl.children).toHaveLength(1);
    expect(text?.textContent).toBe("   Reference");
    expect(text?.dataset["collapsedLeadingSpaces"]).toBeUndefined();
    expect(text?.style["fontSize"]).toBeUndefined();
  });

  test("keeps spaces after a manual line break paintable", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "manual-break-leading-spaces",
      runs: [{ kind: "lineBreak" }, { kind: "text", text: "   Reference" }],
    };
    const line: MeasuredLine = {
      fromRun: 1,
      fromChar: 0,
      toRun: 1,
      toChar: "   Reference".length,
      width: 84,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: false,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const [text] = lineEl.children;

    expect(lineEl.children).toHaveLength(1);
    expect(text?.textContent).toBe("   Reference");
    expect(text?.dataset["collapsedLeadingSpaces"]).toBeUndefined();
    expect(text?.style["fontSize"]).toBeUndefined();
  });

  test("renders underlined tabs as a continuous rule without a text underline mark", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "signature-line",
      runs: [
        { kind: "text", text: "By:" },
        { kind: "tab", underline: { style: "single" } },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 1,
      toChar: 1,
      width: 200,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;

    const tabEl = findTabEl(lineEl);
    expect(tabEl).toBeDefined();
    expect(tabEl?.style["borderBottom"]).toBe("1px solid currentColor");
    expect(tabEl?.style["textDecorationLine"]).toBe("");
  });

  test("paints consecutive tabs on the document default grid", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "authored-default-tabs",
      runs: [
        { kind: "text", text: "1234567" },
        { kind: "tab" },
        { kind: "tab" },
        { kind: "text", text: "x" },
      ],
      attrs: { defaultTabStopTwips: 600 },
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 3,
      toChar: 1,
      width: 127,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 200,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;

    const tabs = findTabEls(lineEl);
    expect(tabs).toHaveLength(2);
    expect(tabs[0]?.style["width"]).toBe("31px");
    expect(tabs[1]?.style["width"]).toBe("40px");
  });

  test("does not underline raised footnote reference markers", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "footnote-marker",
      runs: [
        {
          kind: "text",
          text: "9",
          footnoteRefId: 9,
          superscript: true,
          underline: { style: "single" },
        },
        {
          kind: "text",
          text: "4",
          endnoteRefId: 4,
          superscript: true,
        },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 1,
      toChar: 1,
      width: 10,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const noteMarker = lineEl.children.at(0);

    // eigenpal/docx-editor#846: raised via a paint-only `position`/`top`
    // offset, never `vertical-align` (which would inflate the line box).
    expect(noteMarker?.style["verticalAlign"]).not.toBe("super");
    expect(noteMarker?.style["position"]).toBe("relative");
    expect(noteMarker?.style["top"]).toBe("-0.4em");
    expect(noteMarker?.style["textDecorationLine"]).toBeUndefined();
    expect(noteMarker?.dataset["noteKind"]).toBe("footnote");
    expect(noteMarker?.dataset["noteId"]).toBe("9");
    expect(lineEl.children.at(1)?.dataset).toMatchObject({ noteKind: "endnote", noteId: "4" });
  });

  test("sizes superscript markers from the run font instead of the parent line", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "footnote-marker",
      runs: [
        {
          kind: "text",
          text: "8",
          fontFamily: "Times New Roman",
          fontSize: 10,
          superscript: true,
        },
      ],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 1,
      width: 10,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const noteMarker = lineEl.children.at(0);

    // eigenpal/docx-editor#846: paint-only offset, font size still derived
    // from the run (folio keeps `getRaisedRunFontSize`, not a hardcoded ratio).
    expect(noteMarker?.style["verticalAlign"]).not.toBe("super");
    expect(noteMarker?.style["position"]).toBe("relative");
    expect(noteMarker?.style["top"]).toBe("-0.4em");
    expect(noteMarker?.style["fontFamily"]).toContain("Times New Roman");
    expect(noteMarker?.style["fontSize"]).toBe("10px");
  });

  test("renders underlined whitespace as a rule segment", () => {
    const block: ParagraphBlock = {
      kind: "paragraph",
      id: "underlined-space",
      runs: [{ kind: "text", text: "  ", underline: { style: "single" } }],
    };
    const line: MeasuredLine = {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 2,
      width: 20,
      ascent: 12,
      descent: 3,
      lineHeight: 15,
    };

    const lineEl = renderLine(block, line, undefined, fakeDocument, {
      availableWidth: 360,
      isLastLine: true,
      isFirstLine: true,
      paragraphEndsWithLineBreak: false,
      leftIndentPx: 0,
    }) as unknown as FakeElement;
    const space = lineEl.children.at(0);

    expect(space?.style["borderBottom"]).toBe("1px solid currentColor");
    expect(space?.style["textDecorationLine"]).toBe("");
  });
});
