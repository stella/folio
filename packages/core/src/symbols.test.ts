import { describe, expect, it } from "bun:test";

import { ALL_SYMBOLS, SYMBOL_CATEGORIES, filterSymbols } from "./symbols";

describe("filterSymbols", () => {
  it("returns [] for an empty or whitespace query so callers fall back to the active category", () => {
    expect(filterSymbols("")).toEqual([]);
    expect(filterSymbols("   ")).toEqual([]);
  });

  it("matches descriptive names case-insensitively", () => {
    const chars = filterSymbols("COPYright").map((s) => s.char);
    expect(chars).toContain("©");
  });

  it("matches an exact character", () => {
    const result = filterSymbols("π");
    expect(result.map((s) => s.char)).toContain("π");
  });

  it("tags each result with its source category", () => {
    const [first] = filterSymbols("copyright");
    expect(first?.category).toBe("Common");
  });

  it("returns nothing for a query that matches no name or char", () => {
    expect(filterSymbols("definitely-not-a-symbol-name")).toEqual([]);
  });
});

describe("symbol catalog", () => {
  it("flattens every category into ALL_SYMBOLS", () => {
    const total = SYMBOL_CATEGORIES.reduce((sum, category) => sum + category.symbols.length, 0);
    expect(ALL_SYMBOLS).toHaveLength(total);
  });

  it("has no duplicate characters (adapters key their grids by char)", () => {
    const chars = ALL_SYMBOLS.map((s) => s.char);
    expect(new Set(chars).size).toBe(chars.length);
  });
});
