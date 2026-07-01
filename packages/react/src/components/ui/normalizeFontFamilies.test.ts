import { describe, expect, test } from "bun:test";

import type { FontOption } from "./FontPicker";
import { normalizeFontFamilies } from "./normalizeFontFamilies";

describe("normalizeFontFamilies", () => {
  test("returns undefined for undefined so the picker keeps its defaults", () => {
    expect(normalizeFontFamilies(undefined)).toBeUndefined();
  });

  test("returns an empty list for an empty array (empty, enabled dropdown)", () => {
    expect(normalizeFontFamilies([])).toEqual([]);
  });

  test("expands a plain string into an 'other'-category FontOption", () => {
    expect(normalizeFontFamilies(["Roboto"])).toEqual([
      { name: "Roboto", fontFamily: "Roboto", category: "other" },
    ]);
  });

  test("passes a FontOption object through unchanged", () => {
    const option: FontOption = {
      name: "Roboto",
      fontFamily: "Roboto, sans-serif",
      category: "sans-serif",
    };
    const [normalized] = normalizeFontFamilies([option]) ?? [];
    expect(normalized).toBe(option);
  });

  test("normalizes a mixed string / FontOption array in order", () => {
    const option: FontOption = { name: "Georgia", fontFamily: "Georgia, serif", category: "serif" };
    expect(normalizeFontFamilies(["Arial", option])).toEqual([
      { name: "Arial", fontFamily: "Arial", category: "other" },
      option,
    ]);
  });
});
