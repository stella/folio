import { describe, expect, test } from "bun:test";

import { wordFontDefinitions } from "../wordFonts";

describe("wordFontDefinitions", () => {
  test("maps Aptos regular, bold, and italic faces to explicit CSS descriptors", () => {
    const fonts = wordFontDefinitions("/word-fonts");

    expect(fonts).toContainEqual({
      family: "Aptos",
      filePath: "/word-fonts/Aptos.ttf",
      weight: 400,
    });
    expect(fonts).toContainEqual({
      family: "Aptos",
      filePath: "/word-fonts/Aptos-Bold-Italic.ttf",
      weight: 700,
      style: "italic",
    });
    expect(fonts).toContainEqual({
      family: "Aptos Narrow",
      filePath: "/word-fonts/Aptos-Narrow.ttf",
      weight: 400,
    });
  });

  test("uses one unique source file for each declared face", () => {
    const paths = wordFontDefinitions("/word-fonts").map(({ filePath }) => filePath);
    expect(new Set(paths).size).toBe(paths.length);
  });
});
