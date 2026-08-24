import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { clearAllCaches } from "../../layout-engine/measure";
import {
  installCanvasMeasureProvider,
  resetCanvasContext,
} from "../../layout-engine/measure/measureContainer";
import type {
  Layout,
  Measure,
  ParagraphBlock,
  ParagraphFragment,
  ParagraphMeasure,
  TableBlock,
  TableMeasure,
} from "../../layout-engine/types";
import {
  clickToPositionInParagraph,
  clickToPositionInTableCell,
  getPositionRect,
} from "./clickToPosition";
import { hitTestTableCell } from "./hitTest";
import { getCaretPosition, selectionToRects } from "./selectionRects";

const originalDocument = globalThis.document;

beforeEach(() => {
  installCanvasMeasureProvider();
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: {
      createElement(tagName: string) {
        if (tagName !== "canvas") {
          return {};
        }
        return {
          getContext() {
            return {
              font: "",
              measureText(text: string) {
                return { width: text.length * 7 };
              },
            };
          },
        };
      },
    },
  });
  clearAllCaches();
  resetCanvasContext();
});

afterEach(() => {
  Object.defineProperty(globalThis, "document", {
    configurable: true,
    value: originalDocument,
  });
  clearAllCaches();
  resetCanvasContext();
});

const makeFixture = (alignment: "left" | "right", bidi?: boolean) => {
  const block: ParagraphBlock = {
    kind: "paragraph",
    id: "footer-page-number",
    pmStart: 1,
    pmEnd: 7,
    runs: [{ kind: "text", text: "abcd", pmStart: 1, pmEnd: 5 }],
    attrs: {
      alignment,
      ...(bidi === undefined ? {} : { bidi }),
      indent: { left: 10, right: 30 },
    },
  };
  const measure: ParagraphMeasure = {
    kind: "paragraph",
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 4,
        width: 28,
        ascent: 10,
        descent: 4,
        lineHeight: 14,
      },
    ],
    totalHeight: 14,
  };
  const fragment: ParagraphFragment = {
    kind: "paragraph",
    blockId: block.id,
    x: 0,
    y: 0,
    width: 200,
    height: 14,
    fromLine: 0,
    toLine: 1,
  };
  const layout: Layout = {
    pageGap: 0,
    pages: [
      {
        number: 1,
        size: { w: 200, h: 200 },
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        fragments: [fragment],
      },
    ],
  };
  const measures: Measure[] = [measure];

  return { block, fragment, layout, measure, measures };
};

const makeTableFixture = () => {
  const paragraph: ParagraphBlock = {
    kind: "paragraph",
    id: "table-paragraph",
    pmStart: 1,
    pmEnd: 5,
    runs: [{ kind: "text", text: "abcd", pmStart: 1, pmEnd: 5 }],
    attrs: {
      alignment: "left",
      bidi: true,
      indent: { left: 10, right: 30 },
    },
  };
  const paragraphMeasure: ParagraphMeasure = {
    kind: "paragraph",
    lines: [
      {
        fromRun: 0,
        fromChar: 0,
        toRun: 0,
        toChar: 4,
        width: 28,
        ascent: 10,
        descent: 4,
        lineHeight: 14,
      },
    ],
    totalHeight: 14,
  };
  const block: TableBlock = {
    kind: "table",
    id: "table",
    rows: [
      {
        id: "row",
        cells: [
          {
            id: "cell",
            padding: { top: 0, right: 15, bottom: 0, left: 5 },
            blocks: [paragraph],
          },
        ],
      },
    ],
    columnWidths: [200],
  };
  const measure: TableMeasure = {
    kind: "table",
    rows: [
      {
        cells: [
          {
            blocks: [paragraphMeasure],
            width: 200,
            height: 14,
          },
        ],
        height: 14,
      },
    ],
    columnWidths: [200],
    totalWidth: 200,
    totalHeight: 14,
  };
  const layout: Layout = {
    pageGap: 0,
    pages: [
      {
        number: 1,
        size: { w: 240, h: 200 },
        margins: { top: 0, right: 0, bottom: 0, left: 0 },
        fragments: [
          {
            kind: "table",
            blockId: block.id,
            x: 20,
            y: 0,
            width: 200,
            height: 14,
            fromRow: 0,
            toRow: 1,
          },
        ],
      },
    ],
  };

  return { block, layout, measure };
};

describe("bidi paragraph interaction geometry", () => {
  test("keeps click, caret, and selection on the mirrored physical left", () => {
    const { block, fragment, layout, measure, measures } = makeFixture("right", true);
    const click = clickToPositionInParagraph({
      fragment,
      block,
      measure,
      pageIndex: 0,
      localX: 38,
      localY: 7,
    });
    const directCaret = getPositionRect(block, measure, 1, 0, 0, fragment.width, 0);
    const caret = getCaretPosition(layout, [block], measures, 1);
    const selection = selectionToRects(layout, [block], measures, 1, 5);

    expect(click?.charOffset).toBe(1);
    expect(directCaret?.x).toBe(30);
    expect(caret?.x).toBe(30);
    expect(selection[0]?.x).toBe(30);
  });

  test("mirrors bidi left to physical right with asymmetric indents", () => {
    const { block, fragment, layout, measure, measures } = makeFixture("left", true);
    const expectedStart = 30 + (fragment.width - 30 - 10 - 28);

    expect(getPositionRect(block, measure, 1, 0, 0, fragment.width, 0)?.x).toBe(expectedStart);
    expect(getCaretPosition(layout, [block], measures, 1)?.x).toBe(expectedStart);
    expect(selectionToRects(layout, [block], measures, 1, 5)[0]?.x).toBe(expectedStart);
  });

  test("leaves explicit bidi=false physical geometry unchanged", () => {
    const { block, fragment, measure } = makeFixture("right", false);
    const expectedStart = 10 + (fragment.width - 10 - 30 - 28);

    expect(getPositionRect(block, measure, 1, 0, 0, fragment.width, 0)?.x).toBe(expectedStart);
  });

  test("leaves absent bidi physical geometry unchanged", () => {
    const { block, fragment, measure } = makeFixture("right");
    const expectedStart = 10 + (fragment.width - 10 - 30 - 28);

    expect(getPositionRect(block, measure, 1, 0, 0, fragment.width, 0)?.x).toBe(expectedStart);
  });

  test("keeps table-cell click and selection on the painted physical side", () => {
    const { block, layout, measure } = makeTableFixture();
    const page = layout.pages[0];
    if (!page) {
      throw new Error("expected table page");
    }
    const hit = hitTestTableCell({ pageIndex: 0, page, pageY: 7 }, [block], [measure], {
      x: 175,
      y: 7,
    });
    if (!hit) {
      throw new Error("expected table cell hit");
    }

    // Content width is 200 - 5 - 15 = 180. Bidi mirrors the authored
    // indents to physical left=30/right=10 and authored left alignment to
    // physical right, placing the 28px line at x=20+5+30+(140-28)=167.
    expect(hit.cellContentWidth).toBe(180);
    expect(hit.cellLocalX).toBe(150);
    expect(clickToPositionInTableCell(hit)).toBe(2);
    expect(selectionToRects(layout, [block], [measure], 1, 5)[0]?.x).toBe(167);
  });
});
