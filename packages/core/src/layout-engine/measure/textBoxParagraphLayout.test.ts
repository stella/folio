import { describe, expect, test } from "bun:test";

import { layoutTextBoxParagraphs } from "./textBoxParagraphLayout";
import type { ParagraphBlock, ParagraphMeasure } from "../types";

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

describe("text box paragraph flow", () => {
  test("rejects mismatched paragraph and measure counts", () => {
    expect(() => layoutTextBoxParagraphs([paragraph("only", "value")], [])).toThrow(
      "block and measure counts must match",
    );
  });

  test("preserves authored spacing around empty paragraphs", () => {
    const blocks = [
      paragraph("first", "value"),
      paragraph("empty", "", { after: 30 }),
      paragraph("last", "value"),
    ];
    const layout = layoutTextBoxParagraphs(blocks, [measure(20), measure(20), measure(20)]);

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
    const layout = layoutTextBoxParagraphs(blocks, [measure(20), measure(20)]);

    expect(layout.placements).toEqual([
      { leadingSpacing: 0, contentHeight: 20 },
      { leadingSpacing: 18, contentHeight: 20 },
    ]);
    expect(layout.totalHeight).toBe(65);
  });
});
