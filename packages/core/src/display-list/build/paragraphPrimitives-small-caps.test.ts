// A mixed-case `w:smallCaps` run paints as one `glyphRun` per same-size
// stretch (see smallCapsCasing.ts and `smallCapsGlyphRunSegments`): no
// backend can draw two glyph sizes from one `fontSizePx`, so the split
// happens once, in the producer, ahead of both the DOM and PDF backends.

import { describe, expect, test } from "bun:test";

import { layoutDocument } from "../../layout-engine/index";
import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type {
  FlowBlock,
  Layout,
  LayoutOptions,
  PageMargins,
  ParagraphBlock,
} from "../../layout-engine/types";
import type { BlockLookup } from "../../layout-painter/index";
import type { DisplayGlyphRun, DisplayList, DisplayPrimitive } from "../types";
import { buildDisplayList } from "./buildDisplayList";

const PAGE_SIZE = { w: 816, h: 1056 };
const MARGINS: PageMargins = { top: 96, right: 96, bottom: 96, left: 96 };
const CONTENT_WIDTH = PAGE_SIZE.w - MARGINS.left - MARGINS.right;
const LAYOUT_OPTIONS: LayoutOptions = { pageSize: PAGE_SIZE, margins: MARGINS };

const buildFrom = (blocks: FlowBlock[]): DisplayList => {
  const measures = measureBlocks(blocks, CONTENT_WIDTH);
  const blockLookup: BlockLookup = new Map();
  for (const [index, block] of blocks.entries()) {
    const measure = measures[index];
    if (measure) {
      blockLookup.set(String(block.id), { block, measure });
    }
  }
  const layout: Layout = layoutDocument(blocks, measures, LAYOUT_OPTIONS);
  return buildDisplayList({ layout, blockLookup });
};

const pagePrimitives = (blocks: FlowBlock[]): readonly DisplayPrimitive[] => {
  const list = buildFrom(blocks);
  const page = list.pages.at(0);
  expect(page).toBeDefined();
  // Index 0 is always the page background rect; the body starts after it.
  return page?.primitives.slice(1) ?? [];
};

const glyphRuns = (primitives: readonly DisplayPrimitive[]): DisplayGlyphRun[] =>
  primitives.filter((primitive): primitive is DisplayGlyphRun => primitive.kind === "glyphRun");

const para = (text: string, smallCaps: boolean, allCaps?: boolean): ParagraphBlock => ({
  kind: "paragraph",
  id: "p",
  runs: [
    {
      kind: "text",
      text,
      fontSize: 12,
      ...(smallCaps ? { smallCaps: true } : {}),
      ...(allCaps ? { allCaps: true } : {}),
    },
  ],
});

describe("buildDisplayList: small-caps glyph run splitting", () => {
  test("splits mixed-case text into full-size and shrunken-capital glyph runs", () => {
    withFakeTextMeasure(() => {
      const runs = glyphRuns(pagePrimitives([para("Alert", true)]));

      expect(runs.map((run) => run.text)).toEqual(["A", "LERT"]);
      // 12pt -> 16px at 96dpi; the synthesized segment draws at 0.8 of that.
      expect(runs.map((run) => run.fontSizePx)).toEqual([16, 12.8]);
    });
  });

  test("keeps a single glyph run when nothing in the text is lowercase", () => {
    withFakeTextMeasure(() => {
      const runs = glyphRuns(pagePrimitives([para("HEADING", true)]));

      expect(runs).toHaveLength(1);
      expect(runs[0]?.text).toBe("HEADING");
      expect(runs[0]?.fontSizePx).toBe(16);
    });
  });

  test("w:caps wins over w:smallCaps: one full-size glyph run, text unchanged", () => {
    withFakeTextMeasure(() => {
      const runs = glyphRuns(pagePrimitives([para("Alert", true, true)]));

      expect(runs).toHaveLength(1);
      expect(runs[0]?.fontSizePx).toBe(16);
    });
  });

  test("the split segments paint back to back with no gap or overlap", () => {
    withFakeTextMeasure(() => {
      const runs = glyphRuns(pagePrimitives([para("Alert", true)]));
      expect(runs).toHaveLength(2);
      const first = runs[0]!;
      const second = runs[1]!;
      const firstWidth = first.advancesPx.reduce((sum, advance) => sum + advance, 0);
      expect(second.xPx).toBeCloseTo(first.xPx + firstWidth, 6);
    });
  });

  test("a run with no smallCaps formatting paints as one glyph run, unaffected", () => {
    withFakeTextMeasure(() => {
      const runs = glyphRuns(pagePrimitives([para("Alert", false)]));

      expect(runs).toHaveLength(1);
      expect(runs[0]?.text).toBe("Alert");
    });
  });
});
