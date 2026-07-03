import { describe, expect, test } from "bun:test";

import { findWordBreaks, isBreakChar } from "./lineBreaks";

describe("findWordBreaks", () => {
  test("breaks after spaces and hyphens in Latin text", () => {
    expect(findWordBreaks("hello world")).toEqual([6]);
    expect(findWordBreaks("well-known")).toEqual([5]);
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
});

describe("isBreakChar", () => {
  test("treats CJK code points and low surrogates as break chars", () => {
    expect(isBreakChar("世")).toBe(true);
    expect(isBreakChar("A")).toBe(false);
    expect(isBreakChar("\uDC00")).toBe(true);
  });
});
