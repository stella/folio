/**
 * The two page renderers must paint the same pages.
 *
 * One is the painter that walks `Layout` directly; the other builds the paint
 * IR from that same layout and paints from it. They are two code paths to one
 * result, which is exactly the arrangement that drifts silently: a change to
 * one is not a compile error in the other, and neither has a test that can see
 * the difference.
 *
 * So this paints every fixture twice and compares what a reader would notice:
 * the page count, each page's box, and the text on each page in order. It does
 * not compare pixel geometry, because the two renderers place glyphs by
 * different mechanisms on purpose (the IR renderer honours the advances the
 * layout engine measured; the existing painter lets inline layout advance
 * them). The raster harness measures that difference; this test guards against
 * a page, a paragraph or a word going missing.
 */

import { afterEach, describe, expect, test } from "bun:test";

import { createDisplayListPagePainter } from "./displayListPagePainter";
import { renderPage } from "../../layout-painter/renderPage";
import { fakeDocument } from "../../layout-painter/__tests__/fakeDom";
import type { RenderContext } from "../../layout-painter/renderUtils";
import { layoutDocxHeadless } from "../../headless-layout";
import {
  getMeasureProvider,
  setMeasureProvider,
} from "../../layout-engine/measure/measureProvider";
import { ptToPx } from "../../layout-engine/measure/measureHelpers";
import type { Page } from "../../layout-engine/types";

/**
 * Real documents, and the synthetic pages for the constructs a real document
 * has only some of: a page's furniture is painted by the same two renderers,
 * and can go missing from one of them just as a paragraph can.
 */
const FIXTURES = [
  { name: "sample.docx", from: "visual" },
  { name: "docx-editor-demo.docx", from: "visual" },
  { name: "podily-bps.docx", from: "visual" },
  { name: "step3-footnotes.docx", from: "corpus" },
  { name: "step3-watermark-text.docx", from: "corpus" },
  { name: "step3-watermark-picture.docx", from: "corpus" },
  { name: "step3-page-borders.docx", from: "corpus" },
  { name: "step3-header-footer-fields.docx", from: "corpus" },
] as const;

const fixtureUrl = ({ name, from }: (typeof FIXTURES)[number]) =>
  from === "visual"
    ? new URL(`../../../../../tests/visual/fixtures/${name}`, import.meta.url)
    : new URL(`../../docx/__tests__/__fixtures__/corpus/${name}`, import.meta.url);

const FIXED_ADVANCE_RATIO = 0.5;

/**
 * A provider with no canvas and no font files, so both renderers are compared
 * on one set of measurements and neither can win by measuring differently.
 */
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
        width: charWidths.reduce((total, width) => total + width, 0),
        charWidths,
        metrics: metricsOf(style.fontSize),
      };
    },
  });
};

/** Whitespace and case are rendering concerns; a missing word is not. */
const comparable = (text: string): string => text.replace(/\s+/gu, "").toLocaleUpperCase();

const textOf = (element: Element): string => element.textContent ?? "";

const boxOf = (element: HTMLElement) => ({
  width: element.style.width,
  height: element.style.height,
});

const contextFor = (page: Page, totalPages: number): RenderContext => ({
  pageNumber: page.logicalNumber,
  totalPages,
  section: "body",
});

describe("the two page renderers paint the same pages", () => {
  const installed = getMeasureProvider();

  afterEach(() => {
    setMeasureProvider(installed);
  });

  for (const fixture of FIXTURES) {
    test(`${fixture.name} paints the same page count, boxes and text either way`, async () => {
      installFixedWidthProvider();
      const bytes = await Bun.file(fixtureUrl(fixture)).arrayBuffer();

      const laidOut = await layoutDocxHeadless(bytes);
      expect(laidOut.isErr()).toBe(false);
      if (laidOut.isErr()) {
        return;
      }
      const { layout, blockLookup, furniture, embeddedFonts } = laidOut.value;

      const legacyOptions = { blockLookup, document: fakeDocument };
      const painter = createDisplayListPagePainter({
        layout,
        blockLookup,
        doc: fakeDocument,
        embeddedFonts,
        ...furniture,
      });

      // The IR renderer builds a page when it paints it, so the list holds
      // exactly the pages painted below: it is checked after the loop.

      for (const [index, page] of layout.pages.entries()) {
        const context = contextFor(page, layout.pages.length);
        const legacy = renderPage(page, context, legacyOptions);
        const fromList = painter.paintPage({ page });

        expect({ page: index + 1, painted: fromList !== null }).toEqual({
          page: index + 1,
          painted: true,
        });
        if (fromList === null) {
          continue;
        }

        expect({ page: index + 1, box: boxOf(fromList) }).toEqual({
          page: index + 1,
          box: boxOf(legacy),
        });

        // Every character the existing painter puts on this page must appear
        // on the same page of the other renderer, in the same order. The IR
        // renderer legitimately paints more (a list marker it draws as glyphs
        // rather than as a CSS marker), so this is containment, not equality.
        const expected = comparable(textOf(legacy));
        const actual = comparable(textOf(fromList));
        let cursor = 0;
        let uncoveredAt = -1;
        for (let position = 0; position < expected.length; position += 1) {
          const found = actual.indexOf(expected[position] ?? "", cursor);
          if (found === -1) {
            uncoveredAt = position;
            break;
          }
          cursor = found + 1;
        }
        expect({
          page: index + 1,
          uncoveredAt,
          missingFrom: uncoveredAt === -1 ? "" : expected.slice(uncoveredAt, uncoveredAt + 40),
        }).toEqual({ page: index + 1, uncoveredAt: -1, missingFrom: "" });
      }

      // Every page the layout has was asked for above, so the list now holds
      // one for each: a mismatch means the list and the layout disagree about
      // which pages exist.
      expect(painter.list().pages).toHaveLength(layout.pages.length);
    });
  }
});
