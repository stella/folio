import { describe, expect, test } from "bun:test";

import {
  getHorizontalScaleFactor,
  normalizeHorizontalScalePercent,
  parseHorizontalScalePercent,
  roundHorizontalScalePercentForSerialization,
} from "./horizontalScale";

describe("horizontal text scale normalization", () => {
  test.each([
    -50,
    12.5,
    99.99,
    601,
    Number.NaN,
    Number.POSITIVE_INFINITY,
    Number.NEGATIVE_INFINITY,
  ])("defaults malformed scale %p to 100%%", (scale) => {
    expect(normalizeHorizontalScalePercent(scale)).toBeUndefined();
    expect(getHorizontalScaleFactor(scale)).toBe(1);
  });

  test.each([
    [0, 0],
    [50, 0.5],
    [100, 1],
    [150, 1.5],
    [600, 6],
  ])("preserves valid scale %p", (scale, factor) => {
    expect(normalizeHorizontalScalePercent(scale)).toBe(scale);
    expect(getHorizontalScaleFactor(scale)).toBe(factor);
  });

  test.each([
    ["0", 0],
    ["0%", 0],
    ["600", 600],
    ["600%", 600],
    ["000600%", 600],
    [" \t50%\r\n", 50],
    [" +100 ", 100],
  ])("parses strict lexical scale %p", (scale, expected) => {
    expect(parseHorizontalScalePercent(scale)).toBe(expected);
  });

  test.each(["", "   ", "0garbage", "1e2", "600oops", "0x10", "50 %", "+50%", "-1", "601", "601%"])(
    "rejects malformed lexical scale %p",
    (scale) => {
      expect(parseHorizontalScalePercent(scale)).toBeUndefined();
    },
  );

  test("quantizes only in-range finite serialization values", () => {
    expect(roundHorizontalScalePercentForSerialization(99.99999)).toBe(100);
    expect(roundHorizontalScalePercentForSerialization(12.5)).toBe(13);
    expect(roundHorizontalScalePercentForSerialization(-0.1)).toBeUndefined();
    expect(roundHorizontalScalePercentForSerialization(600.1)).toBeUndefined();
    expect(roundHorizontalScalePercentForSerialization(Number.NaN)).toBeUndefined();
  });
});
