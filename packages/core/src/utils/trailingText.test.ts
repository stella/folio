import { describe, expect, test } from "bun:test";

import { splitTrailingToken, trailingSpaceStart } from "./trailingText";

// Inputs of this size took seconds with the end-anchored patterns these
// helpers replace; a linear scan finishes in well under a millisecond.
const LARGE = 200_000;
const GENEROUS_BOUND_MS = 1_000;

const elapsedMs = (run: () => unknown): number => {
  const start = performance.now();
  run();
  return performance.now() - start;
};

describe("trailingSpaceStart", () => {
  test.each([
    ["", 0],
    ["abc", 3],
    ["abc ", 3],
    ["a b  ", 3],
    ["   ", 0],
    ["abc\t ", 4],
    ["abc \u00A0", 5],
  ])("finds the trailing spaces of %j at %d", (text, expected) => {
    expect(trailingSpaceStart(text)).toBe(expected);
  });

  test("scans text with long interior space runs in linear time", () => {
    const text = `a${" ".repeat(LARGE)}b`;
    let start = -1;
    expect(elapsedMs(() => (start = trailingSpaceStart(text)))).toBeLessThan(GENEROUS_BOUND_MS);
    expect(start).toBe(text.length);
  });
});

describe("splitTrailingToken", () => {
  test.each([
    ["", undefined],
    ["  \t", undefined],
    ["word", { token: "word", separator: "" }],
    ["a v ", { token: "v", separator: " " }],
    ["one two\u00A0\t", { token: "two", separator: "\u00A0\t" }],
    ["lead  x", { token: "x", separator: "" }],
    ["\u{1F600} ", { token: "\u{1F600}", separator: " " }],
  ])("splits %j", (text, expected) => {
    expect(splitTrailingToken(text)).toEqual(expected);
  });

  test("scans a long final token in linear time", () => {
    const text = `${"x".repeat(LARGE)} y`;
    let result: ReturnType<typeof splitTrailingToken>;
    expect(elapsedMs(() => (result = splitTrailingToken(text)))).toBeLessThan(GENEROUS_BOUND_MS);
    expect(result).toEqual({ token: "y", separator: "" });
  });

  test("scans a long token followed by a long whitespace run in linear time", () => {
    const text = `${"x".repeat(LARGE)}${" ".repeat(LARGE)}`;
    let result: ReturnType<typeof splitTrailingToken>;
    expect(elapsedMs(() => (result = splitTrailingToken(text)))).toBeLessThan(GENEROUS_BOUND_MS);
    expect(result?.token.length).toBe(LARGE);
    expect(result?.separator.length).toBe(LARGE);
  });
});
