/**
 * Script segmentation drives per-character East-Asian font selection, so it
 * must classify boundary code points correctly and split mixed text into
 * font-homogeneous spans without breaking surrogate pairs.
 */

import { describe, expect, test } from "bun:test";

import {
  eastAsiaHintApplies,
  hasCjk,
  hasEastAsiaSlotText,
  isCjkCodePoint,
  isEastAsiaHintCodePoint,
  scriptClassOf,
  segmentByScript,
} from "./scriptSegments";

describe("isCjkCodePoint", () => {
  test("classifies representative East-Asian code points as CJK", () => {
    expect(isCjkCodePoint("世".codePointAt(0)!)).toBe(true); // CJK ideograph
    expect(isCjkCodePoint("あ".codePointAt(0)!)).toBe(true); // Hiragana
    expect(isCjkCodePoint("カ".codePointAt(0)!)).toBe(true); // Katakana
    expect(isCjkCodePoint("한".codePointAt(0)!)).toBe(true); // Hangul syllable
    expect(isCjkCodePoint("、".codePointAt(0)!)).toBe(true); // CJK punctuation
    expect(isCjkCodePoint("ㄅ".codePointAt(0)!)).toBe(true); // Bopomofo (Traditional Chinese)
    expect(isCjkCodePoint("⼀".codePointAt(0)!)).toBe(true); // Kangxi radical
    expect(isCjkCodePoint("Ａ".codePointAt(0)!)).toBe(true); // fullwidth Latin A
    expect(isCjkCodePoint("𠀀".codePointAt(0)!)).toBe(true); // Ext B (astral)
  });

  test("classifies Latin and common punctuation as non-CJK", () => {
    expect(isCjkCodePoint("A".codePointAt(0)!)).toBe(false);
    expect(isCjkCodePoint("z".codePointAt(0)!)).toBe(false);
    expect(isCjkCodePoint("1".codePointAt(0)!)).toBe(false);
    expect(isCjkCodePoint(" ".codePointAt(0)!)).toBe(false);
    expect(isCjkCodePoint(".".codePointAt(0)!)).toBe(false);
    expect(isCjkCodePoint("é".codePointAt(0)!)).toBe(false);
  });
});

describe("hasCjk", () => {
  test("detects any CJK presence and the all-Latin fast path", () => {
    expect(hasCjk("Hello world")).toBe(false);
    expect(hasCjk("Hello 世界")).toBe(true);
    expect(hasCjk("")).toBe(false);
    expect(hasCjk("𠀀")).toBe(true);
  });
});

describe("segmentByScript", () => {
  test("splits mixed text into maximal same-script segments", () => {
    expect(segmentByScript("Hello世界foo")).toEqual([
      { text: "Hello", script: "western" },
      { text: "世界", script: "eastAsia" },
      { text: "foo", script: "western" },
    ]);
  });

  test("returns one segment for single-class input", () => {
    expect(segmentByScript("plain ascii")).toEqual([{ text: "plain ascii", script: "western" }]);
    expect(segmentByScript("日本語")).toEqual([{ text: "日本語", script: "eastAsia" }]);
  });

  // Word routes Arabic, Hebrew and the Indic scripts through the `w:cs` slot,
  // a third answer to "which font does this character take?" that the previous
  // isCjk boolean could not express.
  test("separates complex-script text from western and East-Asian", () => {
    expect(segmentByScript("abcمكتبdef")).toEqual([
      { text: "abc", script: "western" },
      { text: "مكتب", script: "complex" },
      { text: "def", script: "western" },
    ]);
    expect(segmentByScript("שלום世界")).toEqual([
      { text: "שלום", script: "complex" },
      { text: "世界", script: "eastAsia" },
    ]);
  });

  test.each([
    ["Arabic", "ب"],
    ["Hebrew", "א"],
    ["Devanagari", "क"],
    ["Thai", "ก"],
    ["Arabic presentation form", "ﻻ"],
    // Astral: the BMP ranges stop at U+FEFF, so these need their own intervals
    // and exercise the code-point (not code-unit) iteration.
    ["Arabic Extended-C", String.fromCodePoint(0x10_ec_0)],
    ["Adlam", String.fromCodePoint(0x1e_90_0)],
    ["Hanifi Rohingya", String.fromCodePoint(0x10_d0_0)],
  ])("classifies %s as complex script", (_name, char) => {
    expect(segmentByScript(char)).toEqual([{ text: char, script: "complex" }]);
  });

  test("returns no segments for empty input", () => {
    expect(segmentByScript("")).toEqual([]);
  });

  test("keeps an astral ideograph whole within its CJK segment", () => {
    const segments = segmentByScript("x𠀀y");
    expect(segments).toEqual([
      { text: "x", script: "western" },
      { text: "𠀀", script: "eastAsia" },
      { text: "y", script: "western" },
    ]);
    // The astral glyph must not be split across the surrogate pair.
    expect(segments[1]?.text.length).toBe(2);
  });

  test("reassembles exactly to the original text", () => {
    const input = "ABCあいうDEFがぎぐ。XYZ";
    expect(
      segmentByScript(input)
        .map((s) => s.text)
        .join(""),
    ).toBe(input);
  });
});

describe('w:rFonts w:hint="eastAsia"', () => {
  test("moves the shared symbols and punctuation to the East Asian slot", () => {
    for (const cp of [
      0xa7, // section sign
      0xb0, // degree sign
      0xb7, // middle dot
      0xd7, // multiplication sign
      0xf7, // division sign
      0x2014, // em dash
      0x201c, // left double quotation mark
      0x2026, // horizontal ellipsis
      0x2103, // degree Celsius
      0x2460, // circled digit one
      0x25cb, // white circle
      0x0391, // Greek capital alpha
      0x0416, // Cyrillic capital zhe
      0xe000, // private use
    ]) {
      expect({ cp, hinted: scriptClassOf(cp, true) }).toEqual({ cp, hinted: "eastAsia" });
      expect({ cp, plain: scriptClassOf(cp) }).toEqual({ cp, plain: "western" });
    }
  });

  test("leaves Basic Latin and accented Latin letters on the western slot", () => {
    for (const cp of [
      0x41, // A
      0x7a, // z
      0x2e, // full stop
      0x22, // quotation mark
      0xa0, // no-break space
      0xa9, // copyright sign
      0xc9, // E with acute
      0xe9, // e with acute
      0xfc, // u with diaeresis
      0x0101, // a with macron
      0x1ea1, // a with dot below
    ]) {
      expect({ cp, eastAsia: isEastAsiaHintCodePoint(cp) }).toEqual({ cp, eastAsia: false });
      expect({ cp, hinted: scriptClassOf(cp, true) }).toEqual({ cp, hinted: "western" });
    }
  });

  test("does not move complex-script code points", () => {
    expect(scriptClassOf(0x05d0, true)).toBe("complex"); // Hebrew alef
    expect(scriptClassOf(0x0627, true)).toBe("complex"); // Arabic alef
  });

  test("segments hinted text by the slot each character selects", () => {
    expect(segmentByScript("A\u201cB\u201d\u00b7", true)).toEqual([
      { text: "A", script: "western" },
      { text: "\u201c", script: "eastAsia" },
      { text: "B", script: "western" },
      { text: "\u201d\u00b7", script: "eastAsia" },
    ]);
    expect(segmentByScript("A\u201cB\u201d")).toEqual([
      { text: "A\u201cB\u201d", script: "western" },
    ]);
  });

  test("hasEastAsiaSlotText sees hinted symbols only under the hint", () => {
    expect(hasEastAsiaSlotText("50\u00b0C", true)).toBe(true);
    expect(hasEastAsiaSlotText("50\u00b0C")).toBe(false);
    expect(hasEastAsiaSlotText("plain text", true)).toBe(false);
  });

  test("w:cs and w:rtl keep the hint from taking effect", () => {
    expect(eastAsiaHintApplies({ eastAsiaHint: true })).toBe(true);
    expect(eastAsiaHintApplies({ eastAsiaHint: true, forceComplexScript: true })).toBe(false);
    expect(eastAsiaHintApplies({ eastAsiaHint: true, rtl: true })).toBe(false);
    expect(eastAsiaHintApplies({ eastAsiaHint: false })).toBe(false);
    expect(eastAsiaHintApplies({})).toBe(false);
  });
});
