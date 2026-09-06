import { describe, expect, test } from "bun:test";

import { parseSfnt } from "./parse";
import type { SfntFont } from "./parse";
import {
  TEST_FONT_FAMILIES,
  TEST_FONTS_INSTALLED,
  TEST_FONTS_SKIP_REASON,
  readTestFont,
} from "./__tests__/testFonts";
import { decodeWoff } from "./woff";

const LATIN_SMALL_E_ACUTE = 0x00e9;
const CJK_UNIFIED_IDEOGRAPH = 0x4e2d;
const SPACE = 0x20;

const magic = (tag: string): Uint8Array => {
  const bytes = new Uint8Array(64);
  for (let index = 0; index < tag.length; index++) {
    bytes[index] = tag.charCodeAt(index);
  }
  return bytes;
};

const loadFont = async (family: (typeof TEST_FONT_FAMILIES)[number]): Promise<SfntFont> => {
  const sfnt = decodeWoff(await readTestFont({ family }));
  if (sfnt.isErr()) {
    throw sfnt.error;
  }
  const font = parseSfnt(sfnt.value);
  if (font.isErr()) {
    throw font.error;
  }
  return font.value;
};

describe(`parseSfnt over real faces (${TEST_FONTS_SKIP_REASON})`, () => {
  for (const family of TEST_FONT_FAMILIES) {
    test.skipIf(!TEST_FONTS_INSTALLED)(`${family} exposes sane metrics`, async () => {
      const font = await loadFont(family);

      expect(font.unitsPerEm).toBeGreaterThanOrEqual(16);
      expect(font.unitsPerEm).toBeLessThanOrEqual(16384);
      expect(font.ascender).toBeGreaterThan(0);
      expect(font.descender).toBeLessThan(0);
      expect(font.lineGap).toBeGreaterThanOrEqual(0);
      expect(font.capHeight).toBeGreaterThan(0);
      expect(font.xHeight).toBeGreaterThan(0);
      expect(font.xHeight).toBeLessThan(font.capHeight);
      expect(font.italicAngle).toBe(0);
      expect(font.numGlyphs).toBeGreaterThan(100);
      expect(Number.isFinite(font.fsType)).toBe(true);

      const [xMin, yMin, xMax, yMax] = font.bbox;
      expect(xMax).toBeGreaterThan(xMin);
      expect(yMax).toBeGreaterThan(yMin);

      expect(font.postScriptName.length).toBeGreaterThan(0);
      expect(font.postScriptName).not.toContain(" ");
      expect(font.isCff).toBe(false);
    });

    test.skipIf(!TEST_FONTS_INSTALLED)(`${family} maps ASCII and Latin-1`, async () => {
      const font = await loadFont(family);

      for (const character of "ABCXYZabcxyz0123456789") {
        expect(font.glyphIdFor(character.codePointAt(0) ?? 0)).toBeGreaterThan(0);
      }
      expect(font.glyphIdFor(LATIN_SMALL_E_ACUTE)).toBeGreaterThan(0);
      expect(font.glyphIdFor(SPACE)).toBeGreaterThan(0);

      // Unmapped and out-of-range inputs land on .notdef rather than throwing.
      expect(font.glyphIdFor(CJK_UNIFIED_IDEOGRAPH)).toBe(0);
      expect(font.glyphIdFor(-1)).toBe(0);
      expect(font.glyphIdFor(0x110000)).toBe(0);
      expect(font.glyphIdFor(Number.NaN)).toBe(0);
    });

    test.skipIf(!TEST_FONTS_INSTALLED)(`${family} has positive, stable advances`, async () => {
      const font = await loadFont(family);

      for (const character of "AWimx ") {
        const glyphId = font.glyphIdFor(character.codePointAt(0) ?? 0);
        const advance = font.advanceWidthFor(glyphId);
        expect(advance).toBeGreaterThan(0);
        expect(advance).toBeLessThanOrEqual(font.unitsPerEm * 4);
        expect(font.advanceWidthFor(glyphId)).toBe(advance);
      }

      // Past the last longHorMetric the advance clamps instead of failing.
      expect(font.advanceWidthFor(font.numGlyphs * 2)).toBeGreaterThanOrEqual(0);
      expect(font.advanceWidthFor(-5)).toBe(0);
    });

    test.skipIf(!TEST_FONTS_INSTALLED)(`${family} reports ink bounds per glyph`, async () => {
      const font = await loadFont(family);

      const capital = font.glyphBoundsFor(font.glyphIdFor("H".codePointAt(0) ?? 0));
      expect(capital).not.toBeNull();
      expect(capital?.yMax).toBeGreaterThan(0);
      expect(capital?.xMax).toBeGreaterThan(capital?.xMin ?? 0);

      const descender = font.glyphBoundsFor(font.glyphIdFor("g".codePointAt(0) ?? 0));
      expect(descender).not.toBeNull();
      expect(descender?.yMin).toBeLessThan(0);

      // A space has no outline, so it has no ink bounds.
      expect(font.glyphBoundsFor(font.glyphIdFor(SPACE))).toBeNull();
      expect(font.glyphBoundsFor(font.numGlyphs)).toBeNull();
    });
  }

  test.skipIf(!TEST_FONTS_INSTALLED)("bold and italic faces differ from the regular", async () => {
    const regular = await loadFont("arimo");
    const italicBytes = decodeWoff(
      await readTestFont({ family: "arimo", weight: 400, style: "italic" }),
    );
    const boldBytes = decodeWoff(await readTestFont({ family: "arimo", weight: 700 }));
    expect(italicBytes.isErr()).toBe(false);
    expect(boldBytes.isErr()).toBe(false);
    if (italicBytes.isErr() || boldBytes.isErr()) {
      return;
    }

    const italic = parseSfnt(italicBytes.value);
    const bold = parseSfnt(boldBytes.value);
    expect(italic.isErr()).toBe(false);
    expect(bold.isErr()).toBe(false);
    if (italic.isErr() || bold.isErr()) {
      return;
    }

    expect(italic.value.italicAngle).toBeLessThan(0);
    expect(italic.value.postScriptName).not.toBe(regular.postScriptName);
    expect(bold.value.postScriptName).not.toBe(regular.postScriptName);

    const glyphId = bold.value.glyphIdFor("W".codePointAt(0) ?? 0);
    expect(bold.value.advanceWidthFor(glyphId)).toBeGreaterThan(0);
  });

  test.skipIf(!TEST_FONTS_INSTALLED)("rejects truncated sfnt bytes", async () => {
    const sfnt = decodeWoff(await readTestFont({ family: "carlito" }));
    expect(sfnt.isErr()).toBe(false);
    if (sfnt.isErr()) {
      return;
    }

    for (const length of [4, 12, 40, 200, 1024]) {
      const parsed = parseSfnt(sfnt.value.subarray(0, length));
      expect(parsed.isErr()).toBe(true);
    }

    // Truncating the body but not the directory must also fail, not read past
    // the buffer.
    const halved = parseSfnt(sfnt.value.subarray(0, Math.floor(sfnt.value.byteLength / 2)));
    expect(halved.isErr()).toBe(true);
  });
});

describe("parseSfnt container handling", () => {
  test("rejects WOFF magic and names the decoder", () => {
    const parsed = parseSfnt(magic("wOFF"));
    expect(parsed.isErr()).toBe(true);
    if (parsed.isErr()) {
      expect(parsed.error.message).toContain("decodeWoff");
    }
  });

  test("rejects WOFF2 magic", () => {
    const parsed = parseSfnt(magic("wOF2"));
    expect(parsed.isErr()).toBe(true);
    if (parsed.isErr()) {
      expect(parsed.error.message).toContain("WOFF2");
    }
  });

  test("rejects a font collection", () => {
    const parsed = parseSfnt(magic("ttcf"));
    expect(parsed.isErr()).toBe(true);
    if (parsed.isErr()) {
      expect(parsed.error.message).toContain("collection");
    }
  });

  test("rejects an unknown version", () => {
    const parsed = parseSfnt(magic("RIFF"));
    expect(parsed.isErr()).toBe(true);
    if (parsed.isErr()) {
      expect(parsed.error.message).toContain("unrecognized sfnt version");
    }
  });

  test("rejects an empty buffer", () => {
    expect(parseSfnt(new Uint8Array(0)).isErr()).toBe(true);
  });

  test("rejects a directory that overruns the buffer", () => {
    const bytes = new Uint8Array(64);
    const view = new DataView(bytes.buffer);
    view.setUint32(0, 0x00010000);
    view.setUint16(4, 4000);
    const parsed = parseSfnt(bytes);
    expect(parsed.isErr()).toBe(true);
  });

  test("rejects a font whose tables are all missing", () => {
    const bytes = new Uint8Array(12);
    const view = new DataView(bytes.buffer);
    view.setUint32(0, 0x00010000);
    view.setUint16(4, 0);
    const parsed = parseSfnt(bytes);
    expect(parsed.isErr()).toBe(true);
    if (parsed.isErr()) {
      expect(parsed.error.message).toContain("head");
    }
  });
});
