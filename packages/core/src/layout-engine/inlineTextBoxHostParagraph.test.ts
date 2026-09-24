import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type {
  FlowBlock,
  LayoutOptions,
  Measure,
  ParagraphAttrs,
  ParagraphBlock,
  ParagraphMeasure,
  TextBoxBlock,
  TextBoxMeasure,
} from "./types";

// ECMA-376 §17.3.3: an inline drawing is run content of its paragraph, so a
// paragraph holding only an inline text box still spaces, indents and aligns
// the line the box sits on.

const layoutOptions: LayoutOptions = {
  pageSize: { w: 600, h: 1000 },
  margins: { top: 50, right: 50, bottom: 50, left: 50 },
  pageGap: 0,
};

const paragraph = (id: string, attrs?: ParagraphAttrs): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: "Text", pmStart: 1, pmEnd: 5 }],
  ...(attrs ? { attrs } : {}),
  pmStart: 0,
  pmEnd: 6,
});

const paragraphMeasure = (lineHeight: number): ParagraphMeasure => ({
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 4,
      width: 40,
      ascent: lineHeight * 0.8,
      descent: lineHeight * 0.2,
      lineHeight,
    },
  ],
  totalHeight: lineHeight,
});

const hostedBox = (hostParagraph?: ParagraphAttrs): TextBoxBlock => ({
  kind: "textBox",
  id: "box",
  width: 200,
  height: 100,
  effectExtent: { left: 4, top: 4, right: 6, bottom: 6 },
  ...(hostParagraph ? { hostParagraph } : {}),
  content: [],
});

const boxMeasure: TextBoxMeasure = {
  kind: "textBox",
  width: 200,
  height: 100,
  innerMeasures: [],
};

const layoutBoxBetween = (hostParagraph?: ParagraphAttrs) => {
  const blocks: FlowBlock[] = [
    paragraph("before", { spacing: { after: 8 } }),
    hostedBox(hostParagraph),
    paragraph("after"),
  ];
  const measures: Measure[] = [paragraphMeasure(20), boxMeasure, paragraphMeasure(20)];
  const fragments = layoutDocument(blocks, measures, layoutOptions).pages[0]?.fragments ?? [];
  return {
    box: fragments.find((fragment) => fragment.kind === "textBox"),
    after: fragments.find(
      (fragment) => fragment.kind === "paragraph" && fragment.blockId === "after",
    ),
  };
};

describe("inline text box host paragraph", () => {
  test("the host paragraph's spacing surrounds the box and collapses with its neighbours'", () => {
    const { box, after } = layoutBoxBetween({ spacing: { before: 12, after: 15 } });

    // The larger of the previous paragraph's 8 after and the host's 12 before.
    expect(box?.y).toBe(50 + 20 + 12 + 4);
    // The next paragraph starts after the reserved bottom edge and the host's after.
    expect(after?.y).toBe(50 + 20 + 12 + 4 + 100 + 6 + 15);
  });

  test("the host paragraph's indentation and alignment place the box on its line", () => {
    const centered = layoutBoxBetween({ alignment: "center", indent: { left: 20, right: 40 } });
    // Column 500 wide; 440 remain inside the indents for a 210-wide occupied box.
    expect(centered.box?.x).toBe(50 + 20 + (440 - 210) / 2 + 4);

    const right = layoutBoxBetween({ alignment: "right", indent: { right: 40 } });
    expect(right.box?.x).toBe(50 + (460 - 210) + 4);

    const indented = layoutBoxBetween({ indent: { left: 30, firstLine: 12 } });
    expect(indented.box?.x).toBe(50 + 42 + 4);
  });

  test("a box without a host paragraph keeps the column start and no spacing of its own", () => {
    const { box, after } = layoutBoxBetween();

    expect(box?.x).toBe(54);
    expect(box?.y).toBe(50 + 20 + 8 + 4);
    expect(after?.y).toBe(50 + 20 + 8 + 4 + 100 + 6);
  });
});
