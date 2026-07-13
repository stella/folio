import { describe, expect, test } from "bun:test";

import { layoutDocument } from "./index";
import type { LayoutOptions, ParagraphBlock, ParagraphMeasure } from "./types";

const layoutOptions: LayoutOptions = {
  pageSize: { w: 600, h: 1200 },
  margins: { top: 0, right: 0, bottom: 0, left: 0 },
  pageGap: 20,
};

const paragraphMeasure: ParagraphMeasure = {
  kind: "paragraph",
  lines: [
    {
      fromRun: 0,
      fromChar: 0,
      toRun: 0,
      toChar: 0,
      width: 100,
      ascent: 14,
      descent: 4,
      lineHeight: 18,
    },
  ],
  totalHeight: 18,
};

type NumberedParagraphOptions = {
  id: string;
  numId: number;
  automaticSpacing?: { before?: boolean; after?: boolean };
};

const numberedParagraph = ({
  id,
  numId,
  automaticSpacing,
}: NumberedParagraphOptions): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [{ kind: "text", text: id }],
  attrs: {
    spacing: { before: 18, after: 18 },
    ...(automaticSpacing ? { automaticSpacing } : {}),
    numPr: { numId, ilvl: 0 },
  },
});

describe("automatic numbered-list spacing", () => {
  test("suppresses automatic spacing inside one sequence", () => {
    const first = numberedParagraph({
      id: "first",
      numId: 4,
      automaticSpacing: { after: true },
    });
    const second = numberedParagraph({
      id: "second",
      numId: 4,
      automaticSpacing: { before: true },
    });

    layoutDocument([first, second], [paragraphMeasure, paragraphMeasure], layoutOptions);

    expect(first.attrs?.spacing).toEqual({ before: 18, after: 0 });
    expect(second.attrs?.spacing).toEqual({ before: 0, after: 18 });
  });

  test("keeps automatic spacing at a sequence boundary", () => {
    const first = numberedParagraph({
      id: "first",
      numId: 4,
      automaticSpacing: { after: true },
    });
    const second = numberedParagraph({
      id: "second",
      numId: 5,
      automaticSpacing: { before: true },
    });

    layoutDocument([first, second], [paragraphMeasure, paragraphMeasure], layoutOptions);

    expect(first.attrs?.spacing).toEqual({ before: 18, after: 18 });
    expect(second.attrs?.spacing).toEqual({ before: 18, after: 18 });
  });

  test("keeps explicit spacing inside one sequence", () => {
    const first = numberedParagraph({ id: "first", numId: 4 });
    const second = numberedParagraph({ id: "second", numId: 4 });

    layoutDocument([first, second], [paragraphMeasure, paragraphMeasure], layoutOptions);

    expect(first.attrs?.spacing).toEqual({ before: 18, after: 18 });
    expect(second.attrs?.spacing).toEqual({ before: 18, after: 18 });
  });
});
