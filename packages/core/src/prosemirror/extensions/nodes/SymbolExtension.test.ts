import { describe, expect, test } from "bun:test";

import { schema } from "../../schema";

class FakeHTMLElement {
  readonly dataset: Record<string, string>;

  constructor(dataset: Record<string, string>) {
    this.dataset = dataset;
  }
}

const parseSymbolAttrs = (dataset: Record<string, string>) => {
  const getAttrs = schema.nodes.symbol.spec.parseDOM?.at(0)?.getAttrs;
  if (!getAttrs) {
    throw new Error("SymbolExtension must define parseDOM[0].getAttrs");
  }

  const originalElement = globalThis.HTMLElement;
  Object.defineProperty(globalThis, "HTMLElement", {
    configurable: true,
    value: FakeHTMLElement,
  });

  try {
    return getAttrs(new FakeHTMLElement(dataset) as unknown as HTMLElement);
  } finally {
    if (originalElement) {
      Object.defineProperty(globalThis, "HTMLElement", {
        configurable: true,
        value: originalElement,
      });
    } else {
      Reflect.deleteProperty(globalThis, "HTMLElement");
    }
  }
};

describe("SymbolExtension parseDOM", () => {
  test("accepts a non-empty font and four-digit hexadecimal character", () => {
    expect(
      parseSymbolAttrs({
        docxSymbolFont: "Wingdings",
        docxSymbolChar: "aF4B",
      }),
    ).toEqual({ font: "Wingdings", char: "aF4B" });
  });

  test.each([
    { docxSymbolFont: "", docxSymbolChar: "F041" },
    { docxSymbolFont: "   ", docxSymbolChar: "F041" },
    { docxSymbolFont: "Wingdings", docxSymbolChar: "041" },
    { docxSymbolFont: "Wingdings", docxSymbolChar: "F0410" },
    { docxSymbolFont: "Wingdings", docxSymbolChar: "GGGG" },
  ])("rejects malformed symbol attributes", (dataset) => {
    expect(parseSymbolAttrs(dataset)).toBe(false);
  });
});

describe("SymbolExtension display", () => {
  test.each([
    ["Symbol", "F0B7", "\u2022"],
    ["Wingdings", "F06E", "\u25A0"],
    ["Wingdings", "F06F", "\u25CB"],
    ["Wingdings", "F0A7", "\u25A0"],
    ["Wingdings", "F0FC", "\u2713"],
  ])("uses the verified Unicode twin for %s %s", (font, char, display) => {
    const symbol = schema.nodes.symbol.create({ font, char });
    const rendered = symbol.type.spec.toDOM?.(symbol);

    expect(symbol.attrs).toEqual({ font, char });
    expect(rendered).toEqual([
      "span",
      {
        "data-docx-symbol-char": char,
        "data-docx-symbol-font": font,
        style: `font-family: "${font}", sans-serif;`,
      },
      display,
    ]);
  });

  test("keeps unmapped symbol codes unchanged", () => {
    const symbol = schema.nodes.symbol.create({
      font: "Wingdings",
      char: "F06C",
    });

    expect(symbol.type.spec.toDOM?.(symbol)).toEqual([
      "span",
      {
        "data-docx-symbol-char": "F06C",
        "data-docx-symbol-font": "Wingdings",
        style: 'font-family: "Wingdings", sans-serif;',
      },
      "\uF06C",
    ]);
  });

  test.each([
    ["Symbol", "006F"],
    ["Wingdings 2", "F06E"],
    ["Other Font", "F06E"],
  ])("keeps unverified %s code %s unchanged", (font, char) => {
    const symbol = schema.nodes.symbol.create({ font, char });

    expect(symbol.type.spec.toDOM?.(symbol)).toEqual([
      "span",
      {
        "data-docx-symbol-char": char,
        "data-docx-symbol-font": font,
        style: `font-family: "${font}", sans-serif;`,
      },
      String.fromCodePoint(Number.parseInt(char, 16)),
    ]);
  });
});
