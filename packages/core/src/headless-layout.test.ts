/**
 * The headless pipeline over real packages.
 *
 * Measurement is supplied by a fixed-width provider rather than the canvas
 * backend: the point of these assertions is that pagination runs with no DOM
 * at all, which a canvas-backed provider would quietly hide.
 */

import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { buildDisplayList } from "./display-list/build/buildDisplayList";
import type { DisplayPrimitive } from "./display-list/types";
import { installHeadlessMeasureProvider } from "./fonts/headlessMeasure";
import { layoutDocxHeadless } from "./headless-layout";
import {
  getMeasureProvider,
  resetMeasureProvider,
  setMeasureProvider,
} from "./layout-engine/measure/measureProvider";
import { getPageTextFromLayout } from "./paged-layout/pageText";
import { ptToPx } from "./layout-engine/measure/measureHelpers";

const FIXTURE = new URL("../../../tests/visual/fixtures/sample.docx", import.meta.url);

/**
 * The provider is process-wide state installed once by the bun preload, so
 * every test here restores what it found. Leaving a provider behind would fail
 * whichever file bun happens to run next, in a way that reads as that file's
 * bug rather than this one's.
 */
let installedProvider = getMeasureProvider();

beforeEach(() => {
  installedProvider = getMeasureProvider();
});

afterEach(() => {
  setMeasureProvider(installedProvider);
});

/** Every glyph is half an em wide; enough to paginate, cheap to reason about. */
const FIXED_ADVANCE_RATIO = 0.5;

const installFixedWidthProvider = (): void => {
  const metricsOf = (fontSize: number | undefined) => {
    const px = ptToPx(fontSize ?? 11);
    return {
      fontSize: fontSize ?? 11,
      ascent: px * 0.8,
      descent: px * 0.2,
      fontBoxAscent: px * 0.9,
      fontBoxDescent: px * 0.25,
      lineHeight: px,
      fontFamily: "fixed",
      singleLineRatio: 1.15,
    };
  };
  const widthOf = (text: string, fontSize: number | undefined) =>
    [...text].length * ptToPx(fontSize ?? 11) * FIXED_ADVANCE_RATIO;

  setMeasureProvider({
    getFontMetrics: (style) => metricsOf(style.fontSize),
    measureTextWidth: (text, style) => widthOf(text, style.fontSize),
    measureText: (text, style) => {
      const metrics = metricsOf(style.fontSize);
      return {
        width: widthOf(text, style.fontSize),
        height: metrics.ascent + metrics.descent,
        ascent: metrics.ascent,
        descent: metrics.descent,
      };
    },
    measureRun: (text, style) => {
      const per = ptToPx(style.fontSize ?? 11) * FIXED_ADVANCE_RATIO;
      const charWidths = [...text].flatMap((char) => (char.length === 2 ? [per, 0] : [per]));
      return {
        width: charWidths.reduce((a, b) => a + b, 0),
        charWidths,
        metrics: metricsOf(style.fontSize),
      };
    },
  });
};

const ARIMO_REGULAR = new URL(
  "../../react/node_modules/@fontsource/arimo/files/arimo-latin-400-normal.woff",
  import.meta.url,
);

const loadFace = async (url: URL): Promise<Uint8Array | null> => {
  const file = Bun.file(url);
  return (await file.exists()) ? new Uint8Array(await file.arrayBuffer()) : null;
};

const arimo = await loadFace(ARIMO_REGULAR);

describe.skipIf(arimo === null)("layoutDocxHeadless with parsed font metrics", () => {
  test("paginates a real package with no canvas anywhere in the chain", async () => {
    // SAFETY: the suite is skipped when the face is absent.
    const face = arimo ?? new Uint8Array();
    const headless = installHeadlessMeasureProvider({ load: () => [face] });
    const bytes = await Bun.file(FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes);

    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    expect(result.value.layout.pages.length).toBeGreaterThan(0);
    // Every face resolved, so the pagination is the authored one rather than
    // one measured against a stand-in.
    expect(headless.substitutions()).toHaveLength(0);
  });
});

describe("layoutDocxHeadless", () => {
  test("paginates a real package with no DOM", async () => {
    installFixedWidthProvider();
    const bytes = await Bun.file(FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes);

    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    const { layout, blockLookup, proseDoc } = result.value;
    expect(layout.pages.length).toBeGreaterThan(0);
    expect(layout.pages.at(0)?.size.w).toBeGreaterThan(0);
    expect(blockLookup.size).toBeGreaterThan(0);

    // Every fragment must resolve to a block and a measure, or a painter has
    // nothing to paint it from.
    for (const page of layout.pages) {
      for (const fragment of page.fragments) {
        expect(blockLookup.has(String(fragment.blockId))).toBe(true);
      }
    }

    expect(getPageTextFromLayout(layout, proseDoc, 1)?.length ?? 0).toBeGreaterThan(0);
  });

  test("refuses to lay out when no measurement backend is installed", async () => {
    resetMeasureProvider();
    const bytes = await Bun.file(FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes);

    expect(result.isErr()).toBe(true);
    if (result.isErr()) {
      expect(result.error.message).toContain("MeasureProvider");
    }
  });

  test("reports the stories it does not paginate rather than dropping them", async () => {
    installFixedWidthProvider();
    const bytes = await Bun.file(FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes);

    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    for (const gap of result.value.unsupported) {
      expect(gap.detail.length).toBeGreaterThan(0);
    }
    // Headers, footers and footnotes are laid out here now; only the endnote
    // story is still collected rather than placed.
    expect(result.value.unsupported.every((gap) => gap.story === "endnote")).toBe(true);
  });
});

const STORY_FIXTURE = new URL(
  "../../../tests/visual/fixtures/docx-editor-demo.docx",
  import.meta.url,
);

const EMBEDDED_FONT_FIXTURE = new URL(
  "../../../tests/visual/fixtures/performance-mixed-script-embedded-font.docx",
  import.meta.url,
);

const glyphTextOf = (primitives: readonly DisplayPrimitive[]): string[] =>
  primitives.flatMap((primitive) => {
    switch (primitive.kind) {
      case "glyphRun":
        return [primitive.text];
      case "clipGroup":
      case "rotateGroup":
      case "opacityGroup":
        return glyphTextOf(primitive.children);
      default:
        return [];
    }
  });

describe("layoutDocxHeadless furniture", () => {
  test("converts the package's header and footer parts without an editing view", async () => {
    installFixedWidthProvider();
    const bytes = await Bun.file(STORY_FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes, { now: new Date("2026-01-01T00:00:00Z") });

    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    const { furniture, layout } = result.value;
    expect(furniture.headerContentByRId?.size ?? 0).toBeGreaterThan(0);
    expect(furniture.footerContentByRId?.size ?? 0).toBeGreaterThan(0);
    for (const content of furniture.headerContentByRId ?? []) {
      expect(content[1].blocks.length).toBeGreaterThan(0);
      expect(content[1].height).toBeGreaterThan(0);
    }
    // The stories are selected per page from the layout's own section refs.
    expect(layout.pages.every((page) => page.headerFooterRefs !== undefined)).toBe(true);
  });

  test("the display list paints the header on every page and the footer's own page number", async () => {
    installFixedWidthProvider();
    const bytes = await Bun.file(STORY_FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes, { now: new Date("2026-01-01T00:00:00Z") });
    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    const { layout, blockLookup, furniture, documentFeatures, embeddedFonts } = result.value;
    const list = buildDisplayList({
      layout,
      blockLookup,
      documentFeatures,
      embeddedFonts,
      ...furniture,
    });

    expect(list.pages.length).toBeGreaterThan(1);
    for (const [index, page] of list.pages.entries()) {
      const texts = glyphTextOf(page.primitives);
      expect(texts.some((text) => text.includes("Project Charter"))).toBe(true);
      // The footer's `PAGE` field resolves against the page it paints on.
      expect(texts).toContain(String(index + 1));
    }
    // Nothing was withheld, so nothing is named.
    expect(list.unsupported).toEqual([]);
  });

  test("an embedded face travels with its bytes", async () => {
    installFixedWidthProvider();
    const bytes = await Bun.file(EMBEDDED_FONT_FIXTURE).arrayBuffer();

    const result = await layoutDocxHeadless(bytes);
    expect(result.isErr()).toBe(false);
    if (result.isErr()) {
      return;
    }
    expect(result.value.embeddedFonts.length).toBeGreaterThan(0);

    const list = buildDisplayList({
      layout: result.value.layout,
      blockLookup: result.value.blockLookup,
      documentFeatures: result.value.documentFeatures,
      embeddedFonts: result.value.embeddedFonts,
      ...result.value.furniture,
    });
    const embedded = list.fonts.flatMap((face) =>
      face.embedded === undefined ? [] : [face.embedded],
    );
    expect(embedded.length).toBeGreaterThan(0);
    expect(embedded.at(0)?.bytes.byteLength ?? 0).toBeGreaterThan(0);
    // One face, one binary: the table interns by id.
    expect(new Set(embedded.map((face) => face.id)).size).toBe(embedded.length);
  });
});
