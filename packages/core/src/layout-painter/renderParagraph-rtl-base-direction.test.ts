// eigenpal #723 (#719) — a paragraph that carries right-to-left runs (`w:rtl`)
// but no explicit paragraph `w:bidi` flag must still lay out right-to-left.
// The painter renders each run as its own `dir`-marked, bidi-isolated span, so
// without a base direction on the fragment the runs stay in logical (LTR) order
// and reversed Hebrew/Arabic reads backwards. We set the fragment direction
// from first-strong base-direction detection, gated to paragraphs that actually
// contain RTL runs so pure-LTR content is untouched.

import { describe, expect, test } from "bun:test";

import type {
  FieldRun,
  ParagraphAttrs,
  ParagraphBlock,
  ParagraphFragment,
  ParagraphMeasure,
  TextRun,
} from "../layout-engine/types";
import { setHyperlinkInstanceIndex } from "../layout-engine/measure/hyperlinkInstance";
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

function render(runs: (TextRun | FieldRun)[], attrs?: ParagraphAttrs): HTMLElement {
  const block: ParagraphBlock = { kind: "paragraph", id: "p1", runs, attrs };
  const totalChars = runs.reduce((n, r) => n + ("text" in r ? r.text.length : 0), 0);
  const measure: ParagraphMeasure = {
    kind: "paragraph",
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: Math.max(0, runs.length - 1),
        toChar: totalChars,
        width: 100,
        ascent: 10,
        descent: 3,
        lineHeight: 13,
      },
    ],
    totalHeight: 13,
  };
  const fragment: ParagraphFragment = {
    kind: "paragraph",
    blockId: "p1",
    x: 0,
    y: 0,
    width: 200,
    height: 13,
    fromLine: 0,
    toLine: 1,
  };
  return renderParagraphFragment(
    fragment,
    block,
    measure,
    { pageNumber: 1, totalPages: 1, section: "body" },
    { document: fakeDocument },
  );
}

function findByTag(element: FakeElement, tagName: string): FakeElement | undefined {
  if (element.tagName === tagName) {
    return element;
  }
  for (const child of element.children) {
    const match = findByTag(child, tagName);
    if (match) {
      return match;
    }
  }
  return undefined;
}

function findAllByTag(element: FakeElement, tagName: string): FakeElement[] {
  const matches = element.tagName === tagName ? [element] : [];
  for (const child of element.children) {
    matches.push(...findAllByTag(child, tagName));
  }
  return matches;
}

const text = (value: string, rtl?: boolean): TextRun => ({
  kind: "text",
  text: value,
  ...(rtl === undefined ? {} : { rtl }),
});

describe("Issue #719 — RTL base direction detection", () => {
  test("Hebrew-led paragraph with rtl runs renders dir=rtl", () => {
    expect(render([text("בדיקה 1", true)]).dir).toBe("rtl");
  });

  test("explicit w:bidi paragraph still renders dir=rtl", () => {
    expect(render([text("hello")], { bidi: true }).dir).toBe("rtl");
  });

  test("applies an asymmetric hanging indent at the RTL logical start edge", () => {
    const paragraph = render([text("عنوان عربي", true)], {
      bidi: true,
      indent: { left: 96, hanging: 48 },
    }) as unknown as FakeElement;
    const line = paragraph.children.at(0);

    expect(line?.style["paddingRight"]).toBe("96px");
    expect(line?.style["paddingLeft"]).toBeUndefined();
    expect(line?.style["textIndent"]).toBe("-48px");
  });

  test("mirrors a negative logical-start indent into the physical right margin", () => {
    const paragraph = render([text("عنوان عربي", true)], {
      bidi: true,
      indent: { left: -9 },
    }) as unknown as FakeElement;

    expect(paragraph.children.at(0)?.style["marginRight"]).toBe("-9px");
  });

  test("mirrors explicit bidi justification onto the physical page", () => {
    expect(render([text("13")], { alignment: "right", bidi: true }).style.textAlign).toBe("left");
    expect(render([text("13")], { alignment: "left", bidi: true }).style.textAlign).toBe("right");
  });

  test("does not mirror explicit alignment or indents for inferred RTL", () => {
    const paragraph = render([text("مرحبا", true)], {
      alignment: "right",
      indent: { left: 96, right: 24 },
    }) as unknown as FakeElement;
    const line = paragraph.children.at(0);

    expect(paragraph.style["textAlign"]).toBe("right");
    expect(line?.style["paddingLeft"]).toBe("96px");
    expect(line?.style["paddingRight"]).toBe("24px");
  });

  test("explicit w:bidi=false wins over rtl runs (stays LTR)", () => {
    // `<w:bidi w:val="0"/>` is an explicit LTR override; first-strong detection
    // must not re-enable RTL for a Hebrew run inside it.
    expect(render([text("בדיקה", true)], { bidi: false }).dir).toBe("");
  });

  test("English-led paragraph with an embedded rtl word stays LTR (no dir)", () => {
    expect(render([text("Hello "), text("שלום", true)]).dir).toBe("");
  });

  test("pure-LTR paragraph is untouched (no dir)", () => {
    expect(render([text("plain text")]).dir).toBe("");
  });

  test("detected-RTL paragraph with no explicit alignment defaults to right-align", () => {
    // Detection must drive the same alignment path as an explicit w:bidi
    // paragraph, not just the `dir` attribute.
    const el = render([text("בדיקה", true)]);
    expect(el.dir).toBe("rtl");
    expect(el.style.textAlign).toBe("right");
  });

  test("explicit bidi marks (RLM/ALM/LRM) decide before letters", () => {
    // U+200F RLM and U+061C ALM are strong RTL; U+200E LRM is strong LTR.
    expect(render([text("‏Hello", true)]).dir).toBe("rtl"); // RLM-led
    expect(render([text("؜Hello", true)]).dir).toBe("rtl"); // ALM-led
    expect(render([text("‎שלום", true)]).dir).toBe(""); // LRM overrides
  });

  test("rtl runs with only digits/punctuation honour w:rtl (no strong char)", () => {
    expect(render([text("123 .", true)]).dir).toBe("rtl");
  });

  test("Arabic-led paragraph with rtl runs renders dir=rtl", () => {
    expect(render([text("مرحبا", true)]).dir).toBe("rtl");
  });

  test("RTL scripts outside the BMP (Adlam) are detected", () => {
    expect(render([text("\u{1E900}\u{1E921}", true)]).dir).toBe("rtl");
  });

  test("a field-result run contributes to base-direction detection", () => {
    // A field result (e.g. a cross-reference) renders as text, so its first
    // strong letter counts.
    const field: FieldRun = {
      kind: "field",
      fieldType: "OTHER",
      fallback: "שלום",
      rtl: true,
    };
    expect(render([field]).dir).toBe("rtl");
  });

  test("a CJK-only rtl run resolves LTR (CJK is strong left-to-right)", () => {
    // CJK, Devanagari, Thai, Hangul and kana are Unicode bidi class L, so a
    // w:rtl run containing only such text must lay out LTR — not fall through to
    // the digits/punctuation `honor w:rtl` path. Also guards the RTL char-class
    // against the pasted-glyph corruption that once swallowed most of the BMP.
    expect(render([text("中文", true)]).dir).toBe("");
    expect(render([text("अआ", true)]).dir).toBe("");
  });

  test("weak chars (Arabic-Indic digits) are skipped; the first letter decides", () => {
    // Arabic-Indic digits are bidi class AN (weak), not strong R — leading them
    // must not trigger RTL. A run led by them whose first *letter* is Latin is
    // LTR.
    expect(render([text("١٢٣ Hello", true)]).dir).toBe("");
  });

  test("isolates a displayed URL left-to-right inside an RTL paragraph", () => {
    const paragraph = render(
      [
        {
          kind: "text",
          text: "https://www.jobaccess.gov.au/find-a-provider",
          hyperlink: { href: "https://www.jobaccess.gov.au/find-a-provider" },
        },
      ],
      { bidi: true },
    ) as unknown as FakeElement;

    expect(findByTag(paragraph, "a")?.dir).toBe("ltr");
  });

  test("isolates every formatting fragment of a displayed URL", () => {
    const href = "https://www.jobaccess.gov.au/find-a-provider";
    const paragraph = render(
      [
        { kind: "text", text: "https://www.jobaccess.", hyperlink: { href } },
        { kind: "text", text: "gov.au/find-a-provider", bold: true, hyperlink: { href } },
      ],
      { bidi: true },
    ) as unknown as FakeElement;

    expect(findAllByTag(paragraph, "a").map((anchor) => anchor.dir)).toEqual(["ltr", "ltr"]);
  });

  test("does not merge adjacent same-target hyperlink instances", () => {
    const href = "https://example.test";
    const urlHyperlink = { href };
    const labelHyperlink = { href };
    setHyperlinkInstanceIndex(urlHyperlink, 0);
    setHyperlinkInstanceIndex(labelHyperlink, 1);
    const paragraph = render(
      [
        { kind: "text", text: href, hyperlink: urlHyperlink },
        { kind: "text", text: "رابط", hyperlink: labelHyperlink },
      ],
      { bidi: true },
    ) as unknown as FakeElement;

    expect(findAllByTag(paragraph, "a").map((anchor) => anchor.dir)).toEqual(["ltr", ""]);
  });

  test("does not force a link label or Arabic link text left-to-right", () => {
    const renderLink = (value: string): FakeElement =>
      render(
        [
          {
            kind: "text",
            text: value,
            hyperlink: { href: "https://example.com" },
          },
        ],
        { bidi: true },
      ) as unknown as FakeElement;

    expect(findByTag(renderLink("معلومات إضافية"), "a")?.dir).toBe("");
    expect(findByTag(renderLink("Job Access"), "a")?.dir).toBe("");
  });

  test("does not force a split link label left-to-right", () => {
    const href = "https://example.com";
    const paragraph = render(
      [
        { kind: "text", text: "Job", hyperlink: { href } },
        { kind: "text", text: " Access", italic: true, hyperlink: { href } },
      ],
      { bidi: true },
    ) as unknown as FakeElement;

    expect(findAllByTag(paragraph, "a").map((anchor) => anchor.dir)).toEqual(["", ""]);
  });
});
