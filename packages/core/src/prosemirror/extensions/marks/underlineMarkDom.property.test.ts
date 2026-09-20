/**
 * The editor's DOM and the page's DOM paint an underline the same way.
 *
 * `UnderlineExtension.toDOM` kept a private four-entry underline → CSS table
 * (`double`, `dotted`, `dash`, `wave`) while the painters read the total one,
 * so a `dottedHeavy`, `dashLongHeavy`, `wavyHeavy` or `words` run drew a plain
 * line in the editor and its own pattern on the page. Both now derive from
 * `underlineDecorationCss`, and these hold them to it over the whole
 * enumeration rather than over the members someone thought to list.
 *
 * The member list is derived twice over: from the one CSS table, which is
 * `satisfies Record<UnderlineStyle, …>`, and from the display list's record,
 * which `outlineDash.property.test.ts` pins to `ST_Underline` in the committed
 * schema graph.
 */

import { describe, expect, test } from "bun:test";
import type { ParseRule, StyleParseRule } from "prosemirror-model";

import { UNDERLINE_STROKES } from "../../../display-list/build/strokes";
import type { MeasuredLine, ParagraphBlock, TextRun } from "../../../layout-engine/types";
import { renderLine } from "../../../layout-painter/renderParagraph";
import type { UnderlineStyle } from "../../../types/document";
import {
  PLAIN_UNDERLINE,
  UNDERLINE_DECORATION_STYLES,
  underlineDecorationCss,
  underlineStyleFromCssDecoration,
} from "../../../utils/formatToStyle";
import { schema } from "../../schema";

const isUnderlineStyle = (value: string): value is UnderlineStyle =>
  Object.hasOwn(UNDERLINE_DECORATION_STYLES, value);

const UNDERLINE_STYLES = Object.keys(UNDERLINE_DECORATION_STYLES).filter(isUnderlineStyle);

/** What a backend decided to paint, in the terms both backends share. */
type PaintedUnderline = {
  underlines: boolean;
  decorationStyle: string | undefined;
  decorationThickness: string | undefined;
};

// ---------------------------------------------------------------------------
// The editor's DOM
// ---------------------------------------------------------------------------

const underlineMarkType = schema.marks["underline"];
if (!underlineMarkType) {
  throw new Error("the schema declares no underline mark");
}

const { toDOM, parseDOM } = underlineMarkType.spec;
if (!toDOM) {
  throw new Error("the underline mark declares no toDOM");
}

const isStyleRule = (rule: ParseRule): rule is StyleParseRule => rule.style !== undefined;

const textDecorationRule = (parseDOM ?? [])
  .filter(isStyleRule)
  .find((rule) => rule.style === "text-decoration");

const declarationsOf = (inlineStyle: string): ReadonlyMap<string, string> =>
  new Map(
    inlineStyle
      .split(";")
      .map((declaration) => declaration.split(":"))
      .filter((parts): parts is [string, string] => parts.length === 2)
      .map(([property, value]) => [property.trim(), value.trim()]),
  );

const editorDeclarations = (style: UnderlineStyle): ReadonlyMap<string, string> => {
  const output = toDOM(underlineMarkType.create({ style }), true);
  const attrs = Array.isArray(output) ? output[1] : undefined;
  const inlineStyle =
    attrs !== null && typeof attrs === "object" && "style" in attrs ? attrs["style"] : undefined;
  return declarationsOf(typeof inlineStyle === "string" ? inlineStyle : "");
};

/** The `text-decoration` shorthand the editor writes: the line and the style. */
const editorShorthand = (style: UnderlineStyle): string =>
  editorDeclarations(style).get("text-decoration") ?? "";

const paintedByEditor = (style: UnderlineStyle): PaintedUnderline => {
  const declarations = editorDeclarations(style);
  const tokens = (declarations.get("text-decoration") ?? "").split(/\s+/u);
  const underlines = tokens.includes("underline");
  return {
    underlines,
    decorationStyle: underlines ? tokens.find((token) => token !== "underline") : undefined,
    decorationThickness: declarations.get("text-decoration-thickness"),
  };
};

/** The member a parse rule read out of a value, or `false` where it refused it. */
const parsedUnderlineStyle = (value: string): UnderlineStyle | false => {
  const attrs = textDecorationRule?.getAttrs?.(value);
  if (attrs === false || attrs === null || attrs === undefined) {
    return false;
  }
  const style = String(attrs["style"]);
  if (!isUnderlineStyle(style)) {
    throw new Error(`the parse rule produced a value outside ST_Underline: ${style}`);
  }
  return style;
};

// ---------------------------------------------------------------------------
// The page's DOM
// ---------------------------------------------------------------------------

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

  getContext(): { font: string; measureText: (text: string) => { width: number } } | null {
    if (this.tagName !== "canvas") {
      return null;
    }
    return { font: "", measureText: (text: string) => ({ width: text.length * 7 }) };
  }
}

const fakeDocument = {
  createElement: (tagName: string) => new FakeElement(tagName),
} as unknown as Document;

const paintedByPage = (style: UnderlineStyle): PaintedUnderline => {
  const runs: TextRun[] = [{ kind: "text", text: "text", underline: { style } }];
  const block: ParagraphBlock = { kind: "paragraph", id: "p", runs };
  const line: MeasuredLine = {
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: 4,
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

  const runEl = lineEl.children.find((child) => child.className.includes("layout-run-text"));
  if (!runEl) {
    throw new Error("the painter rendered no text run");
  }
  return {
    underlines: (runEl.style["textDecorationLine"] ?? "").split(/\s+/u).includes("underline"),
    decorationStyle: runEl.style["textDecorationStyle"],
    decorationThickness: runEl.style["textDecorationThickness"],
  };
};

describe("the underline mark's DOM is the painter's DOM", () => {
  test("the member list is the display list's, which is the schema's", () => {
    expect(UNDERLINE_STYLES.toSorted()).toEqual(Object.keys(UNDERLINE_STROKES).toSorted());
  });

  test.each(UNDERLINE_STYLES)("`%s` paints the same in the editor and on the page", (style) => {
    expect(paintedByEditor(style)).toEqual(paintedByPage(style));
  });

  test.each(UNDERLINE_STYLES)("`%s` paints what the one table says", (style) => {
    const { decorationStyle, decorationThickness } = underlineDecorationCss(style);

    expect(paintedByEditor(style)).toEqual({
      underlines: decorationStyle !== undefined,
      decorationStyle,
      decorationThickness,
    });
  });
});

describe("a `text-decoration` value parses back through the same table", () => {
  test("the editor's own output has a rule to parse back through", () => {
    expect(textDecorationRule).toBeDefined();
  });

  test.each(UNDERLINE_STYLES)("`%s` parses back as a member drawn with the same line", (style) => {
    const parsed = parsedUnderlineStyle(editorShorthand(style));
    if (parsed === false) {
      throw new Error(`the rule refused the editor's own output for ${style}`);
    }

    expect(underlineDecorationCss(parsed).decorationStyle).toBe(
      underlineDecorationCss(style).decorationStyle,
    );
  });

  test("each keyword resolves to its canonical author", () => {
    expect(underlineStyleFromCssDecoration("underline solid")).toBe("single");
    expect(underlineStyleFromCssDecoration("underline double")).toBe("double");
    expect(underlineStyleFromCssDecoration("underline dotted")).toBe("dotted");
    expect(underlineStyleFromCssDecoration("underline dashed")).toBe("dash");
    expect(underlineStyleFromCssDecoration("underline wavy")).toBe("wave");
  });

  test("the weight and the word-only line do not survive CSS, which cannot spell them", () => {
    expect(parsedUnderlineStyle(editorShorthand("dottedHeavy"))).toBe("dotted");
    expect(parsedUnderlineStyle(editorShorthand("words"))).toBe(PLAIN_UNDERLINE);
    expect(parsedUnderlineStyle(editorShorthand("wavyDouble"))).toBe("double");
  });

  test("a value naming no keyword folio writes parses as the plain underline", () => {
    expect(underlineStyleFromCssDecoration("underline")).toBe(PLAIN_UNDERLINE);
    expect(underlineStyleFromCssDecoration("underline groove")).toBe(PLAIN_UNDERLINE);
    expect(underlineStyleFromCssDecoration("underline dottedheavy")).toBe(PLAIN_UNDERLINE);
  });

  test("`none` still cancels an inherited underline, and a strike is not an underline", () => {
    expect(parsedUnderlineStyle("none")).toBe("none");
    expect(parsedUnderlineStyle("line-through dotted")).toBe(false);
  });
});
