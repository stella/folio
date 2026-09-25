import { describe, expect, test } from "bun:test";

import { createFontsourceFaces, cssString, type FontsourceFiles } from "./fontsourceFaces";

const encode = (text: string) => new TextEncoder().encode(text);

/** A fake `@fontsource/arimo` with a latin and a latin-ext face; the ranges overlap at U+0100. */
const arimoFiles = (reads: string[]): FontsourceFiles => {
  const files = new Map<string, Uint8Array>([
    ["arimo/files/arimo-latin-400-normal.woff", encode("latin")],
    ["arimo/files/arimo-latin-ext-400-normal.woff", encode("ext")],
    [
      "arimo/unicode.json",
      encode(JSON.stringify({ latin: "U+0000-00FF,U+0100", "latin-ext": "U+0100-024F" })),
    ],
  ]);
  return {
    read: (packageName, relativePath) => {
      reads.push(`${packageName}/${relativePath}`);
      return files.get(`${packageName}/${relativePath}`) ?? null;
    },
  };
};

describe("createFontsourceFaces", () => {
  test("serves every subset of a face, through the authored-name mapping, cached", () => {
    const reads: string[] = [];
    const faces = createFontsourceFaces(arimoFiles(reads));

    const direct = faces.source.load({ family: "Arimo", bold: false, italic: false });
    const mapped = faces.source.load({ family: "Arial", bold: false, italic: false });
    const missing = faces.source.load({ family: "Arimo", bold: true, italic: false });

    expect(direct.map((bytes) => new TextDecoder().decode(bytes))).toEqual(["latin", "ext"]);
    expect(mapped).toEqual(direct);
    expect(missing).toEqual([]);
    expect(reads.filter((read) => read.endsWith("latin-400-normal.woff")).length).toBe(1);
  });

  test("declares each subset under its range minus what a higher-priority subset serves", () => {
    const faces = createFontsourceFaces(arimoFiles([]));

    const css = faces.fontFaceCss({ families: ["Arimo"] });
    const arimoRules = css
      .split("@font-face")
      .filter((rule) => rule.includes('font-family: "Arimo"') && rule.includes("font-weight: 400;"))
      .filter((rule) => rule.includes("font-style: normal"));

    expect(arimoRules.map((rule) => /unicode-range: ([^;]+);/u.exec(rule)?.[1])).toEqual([
      "U+0000-00FF,U+0100",
      "U+0101-024F",
    ]);
  });

  test("quotes a family name so it closes exactly one CSS string", () => {
    expect(cssString('a"b\\c\nd')).toBe('"a\\"b\\\\c\\A d"');
  });
});
