/**
 * Parity suite: real @chenglou/pretext (via pretextSegmentFitEngine) vs
 * folio's legacy word-walk, both measuring through the deterministic fake
 * canvas (`withFakeTextMeasure`, fixed 5px/char — linear and kerning-free,
 * so widths agree by construction and any divergence is ALGORITHMIC).
 *
 * Characterized outcomes (probe-verified, frozen):
 * - spaced text, edge trailing-space widths, overlong tokens, CJK: EXACT
 *   line-break AND width parity. (Folio's legacy already hangs trailing
 *   spaces, so the edge-width divergence seen against a non-hanging legacy
 *   walk does not exist here.)
 * - canvas call counts (fixed metrics, 100-word paragraph, width 120):
 *   first measure 199 -> 127; repeat measures are 0 on BOTH paths (folio's
 *   width cache already covers steady state — the honest win is first-pass
 *   and changed-text measurement, NOT steady-state).
 * - overlong 400-char token at width 50: 82 -> 3 canvas calls. Killing the
 *   `findMaxFittingLength` slice probes (unique slice keys that pollute the
 *   width cache; the `workerFontMetrics` prewarm exists to soften exactly
 *   this) is folio's headline gain from the seam.
 */
import { afterEach, describe, expect, test } from "bun:test";

import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "@stll/folio-core/layout-engine/measure/__tests__/fakeTextMeasure";
import { setFolioMeasurementFlags } from "@stll/folio-core/layout-engine/measure/featureFlags";
import { measureParagraph } from "@stll/folio-core/layout-engine/measure/measureParagraph";
import {
  resetSegmentFitEngine,
  setSegmentFitEngine,
} from "@stll/folio-core/layout-engine/measure/segmentFit";
import type { ParagraphBlock } from "@stll/folio-core/layout-engine";

import { clearPreparedCache, pretextSegmentFitEngine } from "./pretextEngine";

const fakeMeasure = { charWidth: fixedCharWidth(5) };

function para(text: string): ParagraphBlock {
  return {
    kind: "paragraph",
    id: "p",
    runs: [{ kind: "text", text, fontFamily: "Stub", fontSize: 12 }],
  };
}

function breakOffsets(m: ReturnType<typeof measureParagraph>): Array<[number, number]> {
  return m.lines.map((l) => [l.fromChar, l.toChar]);
}

function measureLegacy(block: ParagraphBlock, width: number) {
  setFolioMeasurementFlags(undefined);
  resetSegmentFitEngine();
  return measureParagraph(block, width);
}

function measureSegment(block: ParagraphBlock, width: number) {
  setFolioMeasurementFlags({ segmentFitLineBreaking: true });
  setSegmentFitEngine(pretextSegmentFitEngine);
  return measureParagraph(block, width);
}

afterEach(() => {
  setFolioMeasurementFlags(undefined);
  resetSegmentFitEngine();
  clearPreparedCache();
});

describe("pretext vs folio legacy line breaks (deterministic metrics)", () => {
  test("spaced text: exact break and width parity", () => {
    withFakeTextMeasure(() => {
      const legacy = measureLegacy(para("aaaa bbbb cccc dddd"), 50);
      const segment = measureSegment(para("aaaa bbbb cccc dddd"), 50);
      expect(breakOffsets(segment)).toEqual(breakOffsets(legacy));
      expect(breakOffsets(segment)).toEqual([
        [0, 10],
        [10, 19],
      ]);
      expect(segment.lines.map((l) => l.width)).toEqual(legacy.lines.map((l) => l.width));
      expect(segment.totalHeight).toBeCloseTo(legacy.totalHeight, 4);
    }, fakeMeasure);
  });

  test("edge trailing-space width: parity (folio legacy already hangs trailing spaces)", () => {
    withFakeTextMeasure(() => {
      const legacy = measureLegacy(para("aaaa bbbb cccc dddd"), 45);
      const segment = measureSegment(para("aaaa bbbb cccc dddd"), 45);
      expect(breakOffsets(segment)).toEqual(breakOffsets(legacy));
      expect(breakOffsets(segment)).toEqual([
        [0, 10],
        [10, 19],
      ]);
      expect(segment.lines.map((l) => l.width)).toEqual(legacy.lines.map((l) => l.width));
    }, fakeMeasure);
  });

  test("overlong token: exact parity with legacy hard-breaking", () => {
    withFakeTextMeasure(() => {
      const token = "x".repeat(45);
      const legacy = measureLegacy(para(token), 50);
      const segment = measureSegment(para(token), 50);
      expect(breakOffsets(segment)).toEqual(breakOffsets(legacy));
      expect(segment.lines.length).toBe(5);
    }, fakeMeasure);
  });

  test("space-less CJK: exact parity (both sides break per ideograph)", () => {
    withFakeTextMeasure(() => {
      const text = "甲乙丙丁戊己庚辛壬癸".repeat(3);
      const legacy = measureLegacy(para(text), 50);
      const segment = measureSegment(para(text), 50);
      expect(breakOffsets(segment)).toEqual(breakOffsets(legacy));
      expect(segment.lines.length).toBe(3);
    }, fakeMeasure);
  });

  test("offset-safety guard: CR and FF texts are declined and measure legacy-identically", () => {
    withFakeTextMeasure(() => {
      const text = "first part\r\nsecond part with more words";
      expect(pretextSegmentFitEngine.supportsText!(text)).toBe(false);
      const legacy = measureLegacy(para(text), 50);
      const segment = measureSegment(para(text), 50);
      expect(breakOffsets(segment)).toEqual(breakOffsets(legacy));
    }, fakeMeasure);
  });

  test("prepared-cache key: identical concatenations get distinct entries", () => {
    clearPreparedCache();
    const a = pretextSegmentFitEngine.prepare("B hello", "12px A");
    const b = pretextSegmentFitEngine.prepare("hello", "12px A B");
    expect(a).not.toBe(b);
  });
});

describe("canvas call profile (probe-frozen)", () => {
  const text100 = Array.from({ length: 100 }, (_, i) => `w${i}`).join(" ");

  test("first-pass measurement drops (199 -> 127); repeats stay 0 on both paths", () => {
    withFakeTextMeasure((getCount) => {
      const l0 = getCount();
      measureLegacy(para(text100), 120);
      const legacyFirst = getCount() - l0;
      const l1 = getCount();
      measureLegacy(para(text100), 120);
      const legacyRepeat = getCount() - l1;

      clearPreparedCache();
      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      setSegmentFitEngine(pretextSegmentFitEngine);
      const s0 = getCount();
      measureParagraph(para(text100), 120);
      const segmentFirst = getCount() - s0;
      const s1 = getCount();
      measureParagraph(para(text100), 120);
      const segmentRepeat = getCount() - s1;

      expect(legacyRepeat).toBe(0); // folio's width cache already covers repeats
      expect(segmentRepeat).toBe(0);
      expect(segmentFirst).toBeLessThan(legacyFirst * 0.75);
    }, fakeMeasure);
  });

  test("overlong-token slice probes collapse (82 -> 3)", () => {
    withFakeTextMeasure((getCount) => {
      const token = "y".repeat(400);
      const l0 = getCount();
      measureLegacy(para(token), 50);
      const legacyProbes = getCount() - l0;

      clearPreparedCache();
      setFolioMeasurementFlags({ segmentFitLineBreaking: true });
      setSegmentFitEngine(pretextSegmentFitEngine);
      const s0 = getCount();
      measureParagraph(para(`${token}z`), 50); // distinct text: no width-cache reuse
      const segmentProbes = getCount() - s0;

      expect(legacyProbes).toBeGreaterThan(50);
      expect(segmentProbes).toBeLessThan(10);
    }, fakeMeasure);
  });
});
