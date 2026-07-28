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
