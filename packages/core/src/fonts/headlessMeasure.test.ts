/**
 * The headless measurement backend over real font binaries.
 *
 * The bundled `@fontsource` faces are a dependency of `@stll/folio-react`, not
 * of this package, so these tests resolve them from disk and skip when they are
 * absent. That keeps the assertions honest (they run against the very faces
 * folio substitutes for Word's) without adding a dependency here for a test.
 */

import { describe, expect, test } from "bun:test";

import { assertNoSubstitutions, createHeadlessMeasureProvider } from "./headlessMeasure";
import type { HeadlessFontSource } from "./headlessMeasure";
import { ptToPx } from "../layout-engine/measure/measureHelpers";

/** The family folio substitutes for Arial, and the file that carries it. */
const ARIMO_REGULAR = new URL(
  "../../../react/node_modules/@fontsource/arimo/files/arimo-latin-400-normal.woff",
  import.meta.url,
);

const loadFace = async (url: URL): Promise<Uint8Array | null> => {
  const file = Bun.file(url);
  return (await file.exists()) ? new Uint8Array(await file.arrayBuffer()) : null;
};

const arimo = await loadFace(ARIMO_REGULAR);

/** Serves the one real face for upright requests and nothing for italic. */
const sourceFrom = (bytes: Uint8Array): HeadlessFontSource => ({
  load: (request) => (request.italic ? null : bytes),
});

describe.skipIf(arimo === null)("headless measurement over a real face", () => {
  // SAFETY: the suite is skipped when the face is absent, so the null branch
  // is unreachable inside it.
  const face = arimo ?? new Uint8Array();

  test("advances are positive and scale linearly with size", () => {
    const { provider } = createHeadlessMeasureProvider(sourceFrom(face));

    const at11 = provider.measureTextWidth("Hamburgefonstiv", {
      fontFamily: "Arial",
      fontSize: 11,
    });
    const at22 = provider.measureTextWidth("Hamburgefonstiv", {
      fontFamily: "Arial",
      fontSize: 22,
    });

    expect(at11).toBeGreaterThan(0);
    expect(at22).toBeCloseTo(at11 * 2, 6);
  });

  test("ascent and descent come from the ink bounds the canvas backend reads", () => {
    const { provider } = createHeadlessMeasureProvider(sourceFrom(face));

    const metrics = provider.getFontMetrics({ fontFamily: "Arial", fontSize: 12 });
    const fontSizePx = ptToPx(12);

    // `H` sits above the baseline and `g` below it, so both must be positive
    // and inside the em: a sign error here moves every baseline on the page.
    expect(metrics.ascent).toBeGreaterThan(0);
    expect(metrics.descent).toBeGreaterThan(0);
    expect(metrics.ascent).toBeLessThan(fontSizePx);
    expect(metrics.descent).toBeLessThan(fontSizePx);
    expect(metrics.lineHeight).toBeGreaterThanOrEqual(fontSizePx);
  });

  test("charWidths sum to the run width and letter spacing lands between code points", () => {
    const { provider } = createHeadlessMeasureProvider(sourceFrom(face));
    const spaced = { fontFamily: "Arial", fontSize: 11, letterSpacing: 2 };

    const run = provider.measureRun("abcd", spaced);
    const plain = provider.measureRun("abcd", { fontFamily: "Arial", fontSize: 11 });

    expect(run.charWidths).toHaveLength(4);
    expect(run.charWidths.reduce((total, width) => total + width, 0)).toBeCloseTo(run.width, 6);
    // Three gaps between four code points, none after the last.
    expect(run.width - plain.width).toBeCloseTo(6, 6);
  });

  test("a face the source cannot supply is reported, never silently substituted", () => {
    const headless = createHeadlessMeasureProvider(sourceFrom(face));

    headless.provider.measureTextWidth("italic", { fontFamily: "Arial", italic: true });

    const substitutions = headless.substitutions();
    expect(substitutions).toHaveLength(1);
    expect(substitutions.at(0)?.requested.italic).toBe(true);
    expect(assertNoSubstitutions(headless).isErr()).toBe(true);
  });

  test("a document that resolves every face reports no substitution", () => {
    const headless = createHeadlessMeasureProvider(sourceFrom(face));

    headless.provider.measureTextWidth("plain", { fontFamily: "Arial" });

    expect(assertNoSubstitutions(headless).isErr()).toBe(false);
  });
});
