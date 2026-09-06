/**
 * Differential test between the two measurement backends' arithmetic.
 *
 * Folio now measures through two backends: a canvas in the browser and parsed
 * font metrics headlessly. They may legitimately disagree about a single
 * glyph's advance, because a canvas kerns and forms ligatures and `hmtx` does
 * not. They may not disagree about anything built on top of that: where letter
 * spacing lands, which font a CJK code point takes, how horizontal scale
 * compounds, where `charWidths` puts the zero for a surrogate pair.
 *
 * So this feeds both paths the same per-character advances and asserts they
 * produce the same numbers. A divergence here would be a defect that no unit
 * test of either backend alone can see: pagination is decided on one and the
 * export is painted with the other.
 */

import { describe, expect, test } from "bun:test";

import { composeRunMeasurement, composeTextWidth } from "./advanceComposition";
import { withFakeTextMeasure, uppercaseAwareCharWidth } from "./__tests__/fakeTextMeasure";
import type { FakeCharWidth } from "./__tests__/fakeTextMeasure";
import { buildFontString } from "./measureHelpers";
import { getMeasureProvider } from "./measureProvider";
import type { FontStyle } from "./measureTypes";
import { getFontKerningMode } from "./textMeasurementPolicy";

/**
 * The same advance source the fake canvas uses, exposed as the headless
 * backend's `advanceOf`. Feeding both sides one source is the whole point:
 * what is compared is the arithmetic, not the glyph data.
 */
const advanceOfFrom =
  (charWidth: FakeCharWidth) =>
  (text: string, style: FontStyle): number => {
    const font = buildFontString(style);
    const kerning = getFontKerningMode(style);
    let width = 0;
    for (const char of text) {
      width += charWidth(char, font, kerning);
    }
    return width;
  };

const STYLES = [
  { label: "plain", style: {} },
  { label: "sized", style: { fontFamily: "Arial", fontSize: 14 } },
  { label: "bold italic", style: { bold: true, italic: true } },
  { label: "letter spaced", style: { letterSpacing: 1.25 } },
  { label: "letter spaced and scaled", style: { letterSpacing: 2, horizontalScale: 150 } },
  { label: "scaled down", style: { horizontalScale: 66 } },
  { label: "uppercased", style: { textTransform: "uppercase" } },
  { label: "small caps", style: { fontVariant: "small-caps" } },
  { label: "kerned", style: { kerning: true } },
  { label: "east asian", style: { fontFamily: "Arial", eastAsiaFontFamily: "MS Mincho" } },
  {
    label: "east asian with letter spacing",
    style: { fontFamily: "Arial", eastAsiaFontFamily: "MS Mincho", letterSpacing: 3 },
  },
  {
    label: "complex script",
    style: { fontFamily: "Arial", complexScriptFontFamily: "Arial", complexScriptFontSize: 13 },
  },
  {
    label: "forced complex script",
    style: { fontFamily: "Arial", complexScriptFontFamily: "Arial", forceComplexScript: true },
  },
] as const satisfies readonly { label: string; style: FontStyle }[];

const TEXTS = [
  "",
  "A",
  "Hello world",
  "MiXeD CaSe TeXt",
  "  leading and trailing  ",
  "日本語のテキスト",
  "Latin 日本語 mixed",
  "العربية",
  "emoji \u{1F600} tail",
  "surrogate \u{20BB7} pair",
] as const;

describe("advance composition matches the canvas backend", () => {
  test("measureTextWidth agrees for every style and text", () => {
    withFakeTextMeasure(
      () => {
        const canvas = getMeasureProvider();
        const advanceOf = advanceOfFrom(uppercaseAwareCharWidth);
        for (const { label, style } of STYLES) {
          for (const text of TEXTS) {
            expect({
              label,
              text,
              width: composeTextWidth({ text, style, advanceOf }),
            }).toEqual({ label, text, width: canvas.measureTextWidth(text, style) });
          }
        }
      },
      { charWidth: uppercaseAwareCharWidth },
    );
  });

  test("measureRun agrees on width and on every charWidths entry", () => {
    withFakeTextMeasure(
      () => {
        const canvas = getMeasureProvider();
        const advanceOf = advanceOfFrom(uppercaseAwareCharWidth);
        for (const { label, style } of STYLES) {
          for (const text of TEXTS) {
            const expected = canvas.measureRun(text, style);
            const actual = composeRunMeasurement({
              text,
              style,
              metrics: expected.metrics,
              advanceOf,
            });
            expect({ label, text, width: actual.width, charWidths: actual.charWidths }).toEqual({
              label,
              text,
              width: expected.width,
              charWidths: expected.charWidths,
            });
          }
        }
      },
      { charWidth: uppercaseAwareCharWidth },
    );
  });

  test("charWidths keeps one entry per UTF-16 unit", () => {
    withFakeTextMeasure(
      () => {
        const text = "a\u{1F600}b";
        const advanceOf = advanceOfFrom(uppercaseAwareCharWidth);
        const run = composeRunMeasurement({
          text,
          style: {},
          metrics: getMeasureProvider().getFontMetrics({}),
          advanceOf,
        });
        expect(run.charWidths).toHaveLength(text.length);
        expect(run.charWidths.at(2)).toBe(0);
      },
      { charWidth: uppercaseAwareCharWidth },
    );
  });
});
