import { describe, expect, test } from "bun:test";

import { calculateChainHeight, computeKeepNextChains } from "./keep-together";
import type { FlowBlock, Measure, ParagraphBlock, ParagraphMeasure } from "./types";

const paragraph = (id: string, keepNext = false): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: id }],
  attrs: { keepNext },
});

const emptyParagraph = (id: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [],
});

const table = (id: string): FlowBlock => ({
  kind: "table",
  id,
  rows: [],
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

  test("reserves every consecutive single-line member through the anchor", () => {
    const blocks: FlowBlock[] = [
      paragraph("empty heading", true),
      paragraph("heading", true),
      paragraph("anchor"),
    ];
    const measures: Measure[] = [paragraphMeasure(12), paragraphMeasure(14), paragraphMeasure(16)];

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
    ).toBe(42);
  });

  test("reserves two anchor lines after a trailing table separator", () => {
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), paragraph("body")];
    const measures: Measure[] = [
      { kind: "table", rows: [], columnWidths: [], totalWidth: 0, totalHeight: 100 },
      paragraphMeasure(12),
      paragraphMeasure(14, 14, 14),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 1,
          endIndex: 1,
          memberIndices: [1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(40);
  });

  test("reserves one anchor line when widow control is disabled", () => {
    const anchor = paragraph("body");
    anchor.attrs = { widowControl: false };
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), anchor];
    const measures: Measure[] = [
      { kind: "table", rows: [], columnWidths: [], totalWidth: 0, totalHeight: 100 },
      paragraphMeasure(12),
      paragraphMeasure(14, 14, 14),
    ];

    expect(
      calculateChainHeight(
        {
          startIndex: 1,
          endIndex: 1,
          memberIndices: [1],
          anchorIndex: 2,
        },
        blocks,
        measures,
      ),
    ).toBe(26);
  });
});

describe("computeKeepNextChains", () => {
  test("carries keepNext through an empty separator to the next content paragraph", () => {
    const blocks: FlowBlock[] = [
      paragraph("heading", true),
      emptyParagraph("separator"),
      paragraph("body"),
    ];

    expect(computeKeepNextChains(blocks).get(0)).toEqual({
      startIndex: 0,
      endIndex: 1,
      memberIndices: [0, 1],
      anchorIndex: 2,
    });
  });

  test("carries a trailing table separator to the next content paragraph", () => {
    const blocks: FlowBlock[] = [table("table"), emptyParagraph("separator"), paragraph("body")];

    expect(computeKeepNextChains(blocks).get(1)).toEqual({
      startIndex: 1,
      endIndex: 1,
      memberIndices: [1],
      anchorIndex: 2,
    });
  });

  test("does not implicitly keep an ordinary empty paragraph with the next paragraph", () => {
    const blocks: FlowBlock[] = [
      paragraph("body before"),
      emptyParagraph("separator"),
      paragraph("body after"),
    ];

    expect(computeKeepNextChains(blocks)).toEqual(new Map());
  });
});
