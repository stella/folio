/**
 * Property — the display list owns an underline's weight and its lines.
 *
 * `ST_Underline` states eight members that differ from a plain rule only in
 * weight (`thick` and the seven `*Heavy`), one that is two waves
 * (`wavyDouble`), and one that skips the gaps between words (`words`). The
 * builder strokes every member at the plain thickness of a single rule before
 * this, so all eight painted as `single` on the page and in the PDF while the
 * editor drew them heavier: the two renderers disagreed about the same run.
 *
 * The expectation is computed from the tables rather than restated, so a new
 * member is covered the moment it is added to them.
 */

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../../layout-engine/index";
import {
  fixedCharWidth,
  withFakeTextMeasure,
} from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type {
  FlowBlock,
  LayoutOptions,
  PageMargins,
  ParagraphBlock,
} from "../../layout-engine/types";
import type { BlockLookup } from "../../layout-painter/index";
import { UNDERLINE_STYLE_VALUES } from "../../types/documentEnumValues";
import type { UnderlineStyle } from "../../types/document";
import { DOUBLE_STROKE_GAP_FACTOR, strokeCrossExtentPx } from "../primitives";
import type { DisplayGlyphRun, DisplayLine, DisplayPrimitive } from "../types";
import { buildDisplayList } from "./buildDisplayList";
import { UNDERLINE_STROKE_COUNTS, UNDERLINE_STROKE_PATTERNS, UNDERLINE_WEIGHTS } from "./strokes";
import { underlineCenterYPx, underlineThicknessPx } from "./textDecorations";

const CHAR_WIDTH_PX = 5;
const fakeMeasure = { charWidth: fixedCharWidth(CHAR_WIDTH_PX) };

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

/** Two two-letter words with one space between them, at a known advance each. */
const TEXT = "ab cd";
const WORD_WIDTH_PX = 2 * CHAR_WIDTH_PX;

const underlined = (style: UnderlineStyle): ParagraphBlock => ({
  kind: "paragraph",
  id: `u-${style}`,
  runs: [{ kind: "text", text: TEXT, underline: { style } }],
});

const pagePrimitives = (blocks: FlowBlock[]): readonly DisplayPrimitive[] => {
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const blockLookup: BlockLookup = new Map();
  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure) {
      blockLookup.set(String(block.id), { block, measure });
    }
  }
  const list = buildDisplayList({
    layout: layoutDocument(blocks, measures, LAYOUT_OPTIONS),
    blockLookup,
  });
  // Index 0 is always the page background rect; the body starts after it.
  return list.pages.at(0)?.primitives.slice(1) ?? [];
};

type Painted = { readonly lines: readonly DisplayLine[]; readonly run: DisplayGlyphRun };

const paint = (style: UnderlineStyle): Painted => {
  const primitives = pagePrimitives([underlined(style)]);
  const run = primitives.find(
    (primitive): primitive is DisplayGlyphRun => primitive.kind === "glyphRun",
  );
  expect(run, `no glyph run painted for ${style}`).toBeDefined();
  if (run === undefined) {
    throw new Error("unreachable");
  }
  return {
    lines: primitives.filter((primitive): primitive is DisplayLine => primitive.kind === "line"),
    run,
  };
};

describe("an underline's weight comes from the total table", () => {
  test("every member strokes at the plain thickness times its weight", () => {
    withFakeTextMeasure(() => {
      for (const style of UNDERLINE_STYLE_VALUES) {
        const { lines, run } = paint(style);
        const pattern = UNDERLINE_STROKE_PATTERNS[style];
        if (pattern === "none") {
          expect(lines, `${style} cancels an underline and paints nothing`).toHaveLength(0);
          continue;
        }

        const expectedPx = underlineThicknessPx(run.fontSizePx) * UNDERLINE_WEIGHTS[style];
        expect(lines.length, `${style} paints no underline`).toBeGreaterThan(0);
        for (const line of lines) {
          expect(line.stroke.pattern, `${style} pattern`).toBe(pattern);
          expect(line.stroke.thicknessPx, `${style} thickness`).toBeCloseTo(expectedPx, 6);
        }
      }
    }, fakeMeasure);
  });

  test("a heavy member is twice the plain member it is heavy for", () => {
    withFakeTextMeasure(() => {
      const plain = paint("single").lines.at(0)?.stroke.thicknessPx;
      const thick = paint("thick").lines.at(0)?.stroke.thicknessPx;
      expect(plain).toBeGreaterThan(0);
      expect(thick).toBe((plain ?? 0) * 2);
      expect(paint("dottedHeavy").lines.at(0)?.stroke.thicknessPx).toBe(thick);
      expect(paint("dotted").lines.at(0)?.stroke.thicknessPx).toBe(plain);
    }, fakeMeasure);
  });

  test("every member paints the number of strokes per word its tables state", () => {
    withFakeTextMeasure(() => {
      for (const style of UNDERLINE_STYLE_VALUES) {
        if (UNDERLINE_STROKE_PATTERNS[style] === "none") {
          continue;
        }
        const words = style === "words" ? 2 : 1;
        expect(paint(style).lines, `${style} stroke count`).toHaveLength(
          words * UNDERLINE_STROKE_COUNTS[style],
        );
      }
    }, fakeMeasure);
  });
});

describe("the members CSS cannot draw", () => {
  test("`words` underlines each word and skips the space between them", () => {
    withFakeTextMeasure(() => {
      const { lines, run } = paint("words");

      expect(lines.map((line) => [line.x1Px - run.xPx, line.x2Px - run.xPx])).toEqual([
        [0, WORD_WIDTH_PX],
        [WORD_WIDTH_PX + CHAR_WIDTH_PX, TEXT.length * CHAR_WIDTH_PX],
      ]);
      // Skipping is the only difference: both spans sit where a `single`
      // underline would, at the same weight.
      const single = paint("single").lines.at(0);
      for (const line of lines) {
        expect(line.y1Px).toBe(single?.y1Px ?? 0);
        expect(line.stroke).toEqual(single?.stroke ?? line.stroke);
      }
    }, fakeMeasure);
  });

  test("`wavyDouble` paints two waves, centred on the single underline", () => {
    withFakeTextMeasure(() => {
      const { lines, run } = paint("wavyDouble");
      const [upper, lower] = lines;
      expect(upper).toBeDefined();
      expect(lower).toBeDefined();
      if (upper === undefined || lower === undefined) {
        return;
      }

      expect(upper.stroke.pattern).toBe("wavy");
      expect(lower.stroke.pattern).toBe("wavy");
      const centerYPx = underlineCenterYPx(run.baselineYPx, run.fontSizePx);
      expect((upper.y1Px + lower.y1Px) / 2).toBeCloseTo(centerYPx, 6);
      // The waves clear each other by the gap `double` leaves between its rules.
      const separationPx =
        strokeCrossExtentPx(upper.stroke) + DOUBLE_STROKE_GAP_FACTOR * upper.stroke.thicknessPx;
      expect(lower.y1Px - upper.y1Px).toBeCloseTo(separationPx, 6);
    }, fakeMeasure);
  });
});
