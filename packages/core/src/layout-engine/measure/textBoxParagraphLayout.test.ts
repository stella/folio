import { describe, expect, test } from "bun:test";

import { layoutTextBoxContent } from "./textBoxParagraphLayout";
import type { ParagraphBlock, ParagraphMeasure, TableBlock, TableMeasure } from "../types";

const paragraph = (
  id: string,
  text: string,
  spacing?: { before?: number; after?: number },
): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: text === "" ? [] : [{ kind: "text", text }],
  attrs: {
    styleId: "authored-style",
    ...(spacing ? { spacing } : {}),
  },
});

const measure = (lineHeight: number): ParagraphMeasure => ({
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 0,
      ascent: lineHeight * 0.8,
      descent: lineHeight * 0.2,
      lineHeight,
    },
  ],
  totalHeight: lineHeight,
});

const table = (id: string): TableBlock => ({
  kind: "table",
  id,
  rows: [],
});

const tableMeasure = (height: number): TableMeasure => ({
  kind: "table",
  rows: [],
  columnWidths: [],
  totalWidth: 100,
  totalHeight: height,
});

describe("text box content flow", () => {
  test("rejects mismatched block and measure counts", () => {
    expect(() => layoutTextBoxContent([paragraph("only", "value")], [])).toThrow(
      "block and measure counts must match",
    );
  });

  test("rejects mismatched block and measure kinds", () => {
    expect(() => layoutTextBoxContent([table("table")], [measure(20)])).toThrow(
      "block and measure kinds must match",
    );
  });

  test("preserves authored spacing around empty paragraphs", () => {
    const blocks = [
      paragraph("first", "value"),
      paragraph("empty", "", { after: 30 }),
      paragraph("last", "value"),
    ];
    const layout = layoutTextBoxContent(blocks, [measure(20), measure(20), measure(20)]);

    expect(layout.placements).toEqual([
      { leadingSpacing: 0, contentHeight: 20 },
      { leadingSpacing: 0, contentHeight: 20 },
      { leadingSpacing: 30, contentHeight: 20 },
    ]);
    expect(layout.totalHeight).toBe(90);
  });

  test("collapses adjacent trailing and leading spacing", () => {
    const blocks = [
      paragraph("first", "value", { after: 18 }),
      paragraph("second", "value", { before: 12, after: 7 }),
    ];
    const layout = layoutTextBoxContent(blocks, [measure(20), measure(20)]);

    expect(layout.placements).toEqual([
      { leadingSpacing: 0, contentHeight: 20 },
      { leadingSpacing: 18, contentHeight: 20 },
    ]);
    expect(layout.totalHeight).toBe(65);
  });

  test("places tables in source order between paragraph spacing", () => {
    const blocks = [
      paragraph("first", "value", { after: 12 }),
      table("table"),
      paragraph("last", "value", { before: 8, after: 4 }),
    ];
    const layout = layoutTextBoxContent(blocks, [measure(20), tableMeasure(40), measure(20)]);

    expect(layout.placements).toEqual([
      { leadingSpacing: 0, contentHeight: 20 },
      { leadingSpacing: 12, contentHeight: 40 },
      { leadingSpacing: 8, contentHeight: 20 },
    ]);
    expect(layout.totalHeight).toBe(104);
  });
});
