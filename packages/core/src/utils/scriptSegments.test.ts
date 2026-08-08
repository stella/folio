/**
 * Script segmentation drives per-character East-Asian font selection, so it
 * must classify boundary code points correctly and split mixed text into
 * font-homogeneous spans without breaking surrogate pairs.
 */

import { describe, expect, test } from "bun:test";

import { hasCjk, isCjkCodePoint, segmentByScript } from "./scriptSegments";

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
