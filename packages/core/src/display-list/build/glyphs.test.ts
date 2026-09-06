/**
 * A run's advances have to say what the line was fitted with, and name what
 * went into them: a backend that hands the text to a shaper cannot recover
 * letter spacing, a horizontal scale, justification, kerning or small capitals
 * from the text, and would paint the run at the glyphs' own width instead.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import type {
  FontMetrics,
  FontStyle,
  RunMeasurement,
} from "../../layout-engine/measure/measureTypes";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "../../layout-engine/measure/measureProvider";
import { buildGlyphs } from "./glyphs";

const METRICS: FontMetrics = {
  fontSize: 12,
  ascent: 9,
  descent: 3,
  fontBoxAscent: 9,
  fontBoxDescent: 3,
  lineHeight: 12,
  fontFamily: "Test",
  singleLineRatio: 1,
};

/** Every character is 10 wide. `stringWidth` overrides what the run measures. */
const install = (stringWidth?: (text: string) => number): void => {
  const charWidths = (text: string): number[] =>
    [...text].flatMap((char) => (char.length === 2 ? [10, 0] : [10]));
  setMeasureProvider({
    getFontMetrics: () => METRICS,
    measureTextWidth: (text: string) => stringWidth?.(text) ?? [...text].length * 10,
    measureText: (text: string) => ({
      width: [...text].length * 10,
      height: 12,
      ascent: 9,
      descent: 3,
    }),
    measureRun: (text: string): RunMeasurement => ({
      width: [...text].length * 10,
      charWidths: charWidths(text),
      metrics: METRICS,
    }),
  });
};

const sum = (values: readonly number[]): number =>
  values.reduce((total, value) => total + value, 0);

const PLAIN: FontStyle = { fontSize: 12 };

describe("buildGlyphs", () => {
  let restore: ReturnType<typeof getMeasureProvider>;

  beforeEach(() => {
    restore = getMeasureProvider();
  });

  afterEach(() => {
    setMeasureProvider(restore);
  });

  test("a run with nothing added to it names no adjustments", () => {
    install();

    const glyphs = buildGlyphs({
      text: "ab",
      style: PLAIN,
      allCaps: false,
      spaceDeltaPx: 0,
      collapsed: false,
    });

    expect(glyphs.adjustments).toBeUndefined();
    expect(glyphs.kerning).toBe(false);
    expect(glyphs.smallCaps).toBe(false);
    expect(glyphs.advancesPx).toEqual([10, 10]);
  });

  test("letter spacing, a horizontal scale and justification are named, not only folded in", () => {
    install();

    const glyphs = buildGlyphs({
      text: "a b",
      style: { ...PLAIN, letterSpacing: 2, horizontalScale: 150 },
      allCaps: false,
      spaceDeltaPx: -1.5,
      collapsed: false,
    });

    expect(glyphs.adjustments).toEqual({
      // Reported in painted pixels, as the advances are: the measured spacing
      // is pre-scale and the scale multiplies it.
      letterSpacingPx: 3,
      horizontalScale: 1.5,
      wordSpacingPx: -1.5,
    });
  });

  test("a collapsed run names no spacing, because it is painted at no width", () => {
    install();

    const glyphs = buildGlyphs({
      text: "a b",
      style: { ...PLAIN, letterSpacing: 2 },
      allCaps: false,
      spaceDeltaPx: -1.5,
      collapsed: true,
    });

    expect(glyphs.adjustments).toBeUndefined();
    expect(sum(glyphs.advancesPx)).toBe(0);
  });

  test("the measurement's own settings travel with the advances", () => {
    install();

    const glyphs = buildGlyphs({
      text: "ab",
      style: { ...PLAIN, kerning: true, fontVariant: "small-caps" },
      allCaps: false,
      spaceDeltaPx: 0,
      collapsed: false,
    });

    expect(glyphs.kerning).toBe(true);
    expect(glyphs.smallCaps).toBe(true);
  });

  test("the advances sum to the width the line was broken on", () => {
    // A pair that kerns or ligates makes the string narrower than its
    // characters are apart. Line breaking used the string's width, so a run
    // whose advances summed to anything else would declare an extent no
    // backend paints.
    install(() => 27);

    const glyphs = buildGlyphs({
      text: "abc",
      style: PLAIN,
      allCaps: false,
      spaceDeltaPx: 0,
      collapsed: false,
    });

    expect(sum(glyphs.advancesPx)).toBeCloseTo(27, 10);
    expect(glyphs.widthPx).toBeCloseTo(27, 10);
    // Spread over the run rather than dropped on its last glyph.
    expect(glyphs.advancesPx).toEqual([9, 9, 9]);
  });

  test("justification is added after the shaped width, not scaled by it", () => {
    install(() => 20);

    const glyphs = buildGlyphs({
      text: "a b",
      style: PLAIN,
      allCaps: false,
      spaceDeltaPx: 3,
      collapsed: false,
    });

    // Three characters measured at 30, shaped to 20, so each is 20/30 of its
    // measured width; the space then gains the justification delta whole.
    expect(sum(glyphs.advancesPx)).toBeCloseTo(23, 10);
  });
});
