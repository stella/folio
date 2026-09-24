import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type {
  FlowBlock,
  LayoutOptions,
  Measure,
  ParagraphBlock,
  ParagraphMeasure,
  TextBoxBlock,
  TextBoxMeasure,
} from "./types";

// ECMA-376 §20.4.2.6: `wp:effectExtent` is added to each edge of an inline
// object, so the line reserves it beyond the object's own extent and the object
// itself sits `l`/`t` inside the reserved space.

const layoutOptions: LayoutOptions = {
  pageSize: { w: 600, h: 1000 },
  margins: { top: 50, right: 50, bottom: 50, left: 50 },
  pageGap: 0,
};

const paragraph = (id: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: "After", pmStart: 1, pmEnd: 6 }],
  pmStart: 0,
  pmEnd: 7,
});

const paragraphMeasure = (lineHeight: number): ParagraphMeasure => ({
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 5,
      width: 40,
      ascent: lineHeight * 0.8,
      descent: lineHeight * 0.2,
      lineHeight,
    },
  ],
  totalHeight: lineHeight,
});

describe("inline text box effect extent", () => {
  test("the box sits inside the space its effect extent reserves", () => {
    const box: TextBoxBlock = {
      kind: "textBox",
      id: "box",
      width: 300,
      height: 100,
      effectExtent: { left: 4, top: 4, right: 13, bottom: 12 },
      content: [],
    };
    const boxMeasure: TextBoxMeasure = {
      kind: "textBox",
      width: 300,
      height: 100,
      innerMeasures: [],
    };
    const blocks: FlowBlock[] = [box, paragraph("after")];
    const measures: Measure[] = [boxMeasure, paragraphMeasure(20)];

    const fragments = layoutDocument(blocks, measures, layoutOptions).pages[0]?.fragments ?? [];
    const boxFragment = fragments.find((fragment) => fragment.kind === "textBox");
    const after = fragments.find((fragment) => fragment.kind === "paragraph");

    expect(boxFragment?.x).toBe(54);
    expect(boxFragment?.y).toBe(54);
    expect(boxFragment?.height).toBe(100);
    // The next block follows the reserved bottom edge, not the box's own.
    expect(after?.y).toBe(50 + 4 + 100 + 12);
  });
});
