import { describe, expect, test } from "bun:test";

import {
  getLineBreakProvider,
  resetLineBreakProvider,
  setLineBreakProvider,
} from "./lineBreakProvider";
import { findGraphemeBreaks, findWordBreaks, isBreakChar } from "./lineBreaks";

describe("findWordBreaks", () => {
  test("breaks after spaces and hyphens in Latin text", () => {
    expect(findWordBreaks("hello world")).toEqual([6]);
    expect(findWordBreaks("well-known")).toEqual([5]);
    expect(findWordBreaks("a  b")).toEqual([2, 3]);
    expect(findWordBreaks("hy\u00ADphen")).toEqual([3]);
  });

  test("adds a break after every CJK code point", () => {
    expect(findWordBreaks("日本語")).toEqual([1, 2, 3]);
    expect(findWordBreaks("A世")).toEqual([2]);
  });

  test("handles astral CJK ideographs as single break units", () => {
    const text = "𠀀文";
    const breaks = findWordBreaks(text);
    expect(breaks).toEqual([2, 3]);
  });

  test("uses dictionary boundaries for scripts without spaces", () => {
    expect(findWordBreaks("ภาษาไทยภาษาไทย", { locale: "th" })).toEqual([4, 7, 11]);
  });

  test("falls back safely for malformed DOCX language tags", () => {
    expect(() => findWordBreaks("hello world", { locale: "not_a_locale" })).not.toThrow();
  });

  test("keeps prohibited CJK punctuation off the next line", () => {
    expect(findWordBreaks("中文。测试", { locale: "zh-CN" })).toEqual([1, 3, 4, 5]);
  });

  test("applies document-specific kinsoku overrides", () => {
    expect(findWordBreaks("中文測", { noLineBreaksBefore: "測" })).toEqual([1, 3]);
  });

  test("applies Word's legacy Ethiopic and Amharic break opportunities", () => {
    const text = "ሀ፡ለ";
    expect(findWordBreaks(text)).toEqual([]);
    expect(findWordBreaks(text, { useLegacyEthiopicAmharicRules: true })).toEqual([2]);
  });
});

describe("findGraphemeBreaks", () => {
  test("never splits an emoji ZWJ sequence or combining character", () => {
    expect(findGraphemeBreaks("👨‍👩‍👧‍👦x")).toEqual([11, 12]);
    expect(findGraphemeBreaks("e\u0301x")).toEqual([2, 3]);
  });
});

describe("line-break provider", () => {
  test("is replaceable and resettable", () => {
    const original = getLineBreakProvider();
    try {
      setLineBreakProvider({
        findBreaks: () => [1],
        findGraphemeBreaks: () => [1],
      });
      expect(findWordBreaks("abc")).toEqual([1]);
    } finally {
      resetLineBreakProvider();
    }
    expect(getLineBreakProvider()).toBe(original);
  });
});

describe("isBreakChar", () => {
  test("treats CJK code points and low surrogates as break chars", () => {
    expect(isBreakChar("世")).toBe(true);
    expect(isBreakChar("A")).toBe(false);
    expect(isBreakChar("\uDC00")).toBe(true);
  });
});
