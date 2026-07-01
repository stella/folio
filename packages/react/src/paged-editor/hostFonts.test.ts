import { describe, expect, test } from "bun:test";

import { toFontFaceInputs } from "./hostFonts";

describe("toFontFaceInputs", () => {
  test("returns [] for undefined and empty input", () => {
    expect(toFontFaceInputs(undefined)).toEqual([]);
    expect(toFontFaceInputs([])).toEqual([]);
  });

  test("wraps src in a CSS url() source and defaults display to swap", () => {
    const [input] = toFontFaceInputs([{ family: "Brand Sans", src: "/fonts/BrandSans.woff2" }]);
    expect(input).toEqual({
      family: "Brand Sans",
      source: 'url("/fonts/BrandSans.woff2")',
      descriptors: { display: "swap" },
    });
  });

  test("stringifies a numeric weight and keeps a keyword weight", () => {
    const inputs = toFontFaceInputs([
      { family: "Brand Sans", src: "/a.woff2", weight: 700 },
      { family: "Brand Sans", src: "/b.woff2", weight: "bold" },
    ]);
    expect(inputs[0]?.descriptors.weight).toBe("700");
    expect(inputs[1]?.descriptors.weight).toBe("bold");
  });

  test("omits the weight descriptor when weight is absent", () => {
    const [input] = toFontFaceInputs([{ family: "Brand Sans", src: "/a.woff2" }]);
    expect(input?.descriptors).not.toHaveProperty("weight");
  });

  test("trims surrounding whitespace on family and src", () => {
    const [input] = toFontFaceInputs([{ family: "  Brand Sans  ", src: "  /a.woff2  " }]);
    expect(input?.family).toBe("Brand Sans");
    expect(input?.source).toBe('url("/a.woff2")');
  });

  test("skips entries with a blank family or src, never throwing", () => {
    const inputs = toFontFaceInputs([
      { family: "", src: "/a.woff2" },
      { family: "  ", src: "/b.woff2" },
      { family: "Brand Sans", src: "" },
      { family: "Brand Sans", src: "   " },
      { family: "Keep", src: "/keep.woff2" },
    ]);
    expect(inputs).toHaveLength(1);
    expect(inputs[0]?.family).toBe("Keep");
  });

  test("escapes quotes in src via JSON.stringify", () => {
    const [input] = toFontFaceInputs([{ family: "Brand Sans", src: '/a".woff2' }]);
    expect(input?.source).toBe('url("/a\\".woff2")');
  });
});
