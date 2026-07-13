import { describe, expect, test } from "bun:test";

import {
  FONT_KERNING_MODE,
  countCompressibleSpaces,
  getFontKerningMode,
  getRunFontKerningMode,
} from "./textMeasurementPolicy";

describe("text measurement policy", () => {
  test("keeps pair kerning disabled without an authored threshold", () => {
    expect(getRunFontKerningMode({ fontSize: 11 }, 12)).toBe(FONT_KERNING_MODE.disabled);
  });

  test("enables pair kerning at the authored threshold", () => {
    expect(getRunFontKerningMode({ fontSize: 11, kerningMinPt: 11 }, 12)).toBe(
      FONT_KERNING_MODE.enabled,
    );
  });

  test("keeps pair kerning disabled below the authored threshold", () => {
    expect(getRunFontKerningMode({ fontSize: 10, kerningMinPt: 11 }, 12)).toBe(
      FONT_KERNING_MODE.disabled,
    );
  });

  test("uses the fallback size when the run omits its size", () => {
    expect(getRunFontKerningMode({ kerningMinPt: 12 }, 12)).toBe(FONT_KERNING_MODE.enabled);
    expect(getRunFontKerningMode({ kerningMinPt: 12 }, 11)).toBe(FONT_KERNING_MODE.disabled);
  });

  test("maps normalized style to the same canvas and CSS modes", () => {
    expect(getFontKerningMode({ kerning: true })).toBe(FONT_KERNING_MODE.enabled);
    expect(getFontKerningMode({ kerning: false })).toBe(FONT_KERNING_MODE.disabled);
    expect(getFontKerningMode({})).toBe(FONT_KERNING_MODE.disabled);
  });

  test("counts only spaces that justification may compress", () => {
    expect(countCompressibleSpaces("one two  three")).toBe(3);
    expect(countCompressibleSpaces("one\u00a0two\tthree\u2003four")).toBe(0);
  });
});
