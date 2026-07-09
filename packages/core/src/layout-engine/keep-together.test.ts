import { describe, expect, test } from "bun:test";

import { calculateChainHeight } from "./keep-together";
import type { FlowBlock, Measure, ParagraphBlock, ParagraphMeasure } from "./types";

const paragraph = (id: string, keepNext = false): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: id }],
  attrs: { keepNext },
});

const paragraphMeasure = (...lineHeights: number[]): ParagraphMeasure => ({
  kind: "paragraph",
  lines: lineHeights.map((lineHeight) => ({
    fromRun: 0,
    fromChar: 0,
    toRun: 0,
    toChar: 1,
    width: 10,
    ascent: lineHeight * 0.8,
    descent: lineHeight * 0.2,
    lineHeight,
  })),
  totalHeight: lineHeights.reduce((total, lineHeight) => total + lineHeight, 0),
});

describe("calculateChainHeight", () => {
  test("reserves only the first line of a splittable successor", () => {
    const blocks: FlowBlock[] = [
      paragraph("first", true),
      paragraph("splittable", true),
      paragraph("anchor"),
    ];
    const measures: Measure[] = [
      paragraphMeasure(12, 12),
      paragraphMeasure(14, 14, 14, 14),
      paragraphMeasure(16),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 0,
          endIndex: 1,
          memberIndices: [0, 1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(38);
  });
});
