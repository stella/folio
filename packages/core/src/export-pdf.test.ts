/**
 * The export's handling of the one piece of process-wide state it touches.
 *
 * `MeasureProvider` is ambient: an export installs its own fonts and must give
 * the caller's provider back, on every exit path and without two overlapping
 * exports being able to see each other's fonts. Neither property is visible in
 * an export's output, so nothing else in the suite would notice them breaking.
 */

import { describe, expect, test } from "bun:test";

import { exportDocxToPdf } from "./export-pdf";
import type { HeadlessFontSource } from "./fonts/headlessMeasure";
import { getMeasureProvider, setMeasureProvider } from "./layout-engine/measure/measureProvider";
import type { MeasureProvider } from "./layout-engine/measure/measureProvider";

const FIXTURE = new URL("../../../tests/visual/fixtures/sample.docx", import.meta.url);

const NO_FONTS: HeadlessFontSource = { load: () => [] };

/** A provider that is recognisable by identity and never actually measures. */
const markerProvider = (): MeasureProvider => ({
  getFontMetrics: () => ({
    fontSize: 11,
    ascent: 1,
    descent: 1,
    fontBoxAscent: 1,
    fontBoxDescent: 1,
    lineHeight: 1,
    fontFamily: "marker",
    singleLineRatio: 1,
  }),
  measureTextWidth: () => 0,
  measureText: () => ({ width: 0, height: 0, ascent: 1, descent: 1 }),
  measureRun: (text) => ({
    width: 0,
    charWidths: [...text].map(() => 0),
    metrics: {
      fontSize: 11,
      ascent: 1,
      descent: 1,
      fontBoxAscent: 1,
      fontBoxDescent: 1,
      lineHeight: 1,
      fontFamily: "marker",
      singleLineRatio: 1,
    },
  }),
});

describe("exportDocxToPdf and the ambient measurement provider", () => {
  test("gives the caller's provider back after a successful export", async () => {
    const caller = markerProvider();
    setMeasureProvider(caller);

    await exportDocxToPdf(await Bun.file(FIXTURE).arrayBuffer(), {
      fonts: NO_FONTS,
      timestamp: "2026-01-01T00:00:00.000Z",
    });

    expect(getMeasureProvider()).toBe(caller);
  });

  test("gives it back when the input cannot be parsed", async () => {
    const caller = markerProvider();
    setMeasureProvider(caller);

    const result = await exportDocxToPdf(new Uint8Array([1, 2, 3]).buffer, {
      fonts: NO_FONTS,
      timestamp: "2026-01-01T00:00:00.000Z",
    });

    expect(result.isErr()).toBe(true);
    expect(getMeasureProvider()).toBe(caller);
  });

  test("two overlapping exports do not observe each other's provider", async () => {
    const caller = markerProvider();
    setMeasureProvider(caller);
    const bytes = await Bun.file(FIXTURE).arrayBuffer();

    // Started together and awaited together: without serialisation the second
    // export replaces the first's provider mid-flight, and the first restores
    // while the second is still measuring.
    const [first, second] = await Promise.all([
      exportDocxToPdf(bytes, { fonts: NO_FONTS, timestamp: "2026-01-01T00:00:00.000Z" }),
      exportDocxToPdf(bytes, { fonts: NO_FONTS, timestamp: "2026-01-01T00:00:00.000Z" }),
    ]);

    expect(first.isErr()).toBe(false);
    expect(second.isErr()).toBe(false);
    // Both ran against the same empty source, so identical input must give
    // identical bytes; a provider swapped mid-flight would not.
    if (!first.isErr() && !second.isErr()) {
      expect(Buffer.from(second.value.bytes).equals(Buffer.from(first.value.bytes))).toBe(true);
    }
    expect(getMeasureProvider()).toBe(caller);
  });
});
