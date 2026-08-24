import { afterEach, beforeEach, describe, expect, test } from "bun:test";

import { clearAllCaches } from "../layout-engine/measure";
import {
  installCanvasMeasureProvider,
  resetCanvasContext,
} from "../layout-engine/measure/measureContainer";
import type {
  Layout,
  ParagraphBlock,
  ParagraphMeasure,
  TableBlock,
  TableFragment,
  TableMeasure,
} from "../layout-engine/types";
import { hitTestTableCell } from "../layout-bridge/engine/hitTest";
import { selectionToRects } from "../layout-bridge/engine/selectionRects";
import { renderTableFragment, TABLE_CLASS_NAMES } from "../layout-painter/renderTable";
import type { RenderContext } from "../layout-painter/renderUtils";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];

  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }

  querySelectorAll(): FakeElement[] {
    return [];
  }
}

const fakeDocument = {
  createElement(): FakeElement {
    return new FakeElement();
  },
  createTextNode(): FakeElement {
    return new FakeElement();
  },
} as unknown as Document;

const renderContext: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: "body",
};

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

const emptyParagraph = (id: string): ParagraphBlock => ({
  kind: "paragraph",
  id,
  runs: [],
});

const emptyParagraphMeasure = (): ParagraphMeasure => ({
  kind: "paragraph",
  lines: [],
  totalHeight: 0,
});

const findCells = (element: FakeElement): FakeElement[] => {
  const cells = element.className.split(" ").includes(TABLE_CLASS_NAMES.cell) ? [element] : [];
  for (const child of element.children) {
    cells.push(...findCells(child));
  }
  return cells;
};

const renderContinuation = (bidi: boolean): FakeElement[] => {
  const block: TableBlock = {
    kind: "table",
    id: "continuation-table",
    ...(bidi ? { bidi: true } : {}),
    columnWidths: [100, 80, 120],
    rows: [
      {
        id: "source-row",
        cells: [
          {
            id: "spanning-cell",
            rowSpan: 2,
            blocks: [emptyParagraph("span")],
          },
          {
            id: "source-body",
            colSpan: 2,
            blocks: [emptyParagraph("source-body-p")],
          },
        ],
      },
      {
        id: "continued-row",
        cells: [
          {
            id: "continued-body",
            colSpan: 2,
            blocks: [emptyParagraph("continued-body-p")],
          },
        ],
      },
    ],
  };
  const measure: TableMeasure = {
    kind: "table",
    columnWidths: [100, 80, 120],
    totalWidth: 300,
    totalHeight: 40,
    rows: [
      {
        height: 20,
        cells: [
          { width: 100, height: 40, blocks: [emptyParagraphMeasure()] },
          { width: 200, height: 20, blocks: [emptyParagraphMeasure()] },
        ],
      },
      {
        height: 20,
        cells: [{ width: 200, height: 20, blocks: [emptyParagraphMeasure()] }],
      },
    ],
  };
  const fragment: TableFragment = {
    kind: "table",
    blockId: block.id,
    x: 0,
    y: 0,
    width: 300,
    height: 20,
    fromRow: 1,
    toRow: 2,
    continuesFromPrev: true,
  };

  const rendered = renderTableFragment(fragment, block, measure, renderContext, {
    document: fakeDocument,
  }) as unknown as FakeElement;
  return findCells(rendered);
};

describe("canonical table-cell placement", () => {
  test("places LTR and RTL continuations beneath a prior-fragment row span", () => {
    for (const [bidi, expectedLeft] of [
      [false, "100px"],
      [true, "0px"],
    ] as const) {
      const cells = renderContinuation(bidi);

      expect(cells).toHaveLength(1);
      expect(cells[0]?.dataset["columnIndex"]).toBe("1");
      expect(cells[0]?.style["left"]).toBe(expectedLeft);
      expect(cells[0]?.style["width"]).toBe("200px");
    }
  });

  test("honors gridBefore and clamps an oversized colSpan to the remaining grid", () => {
    const block: TableBlock = {
      kind: "table",
      id: "grid-before-table",
      columnWidths: [40, 60, 100],
      rows: [
        {
          id: "grid-before-row",
          gridBefore: 2,
          cells: [
            {
              id: "grid-before-cell",
              colSpan: 5,
              blocks: [emptyParagraph("grid-before-p")],
            },
          ],
        },
      ],
    };
    const measure: TableMeasure = {
      kind: "table",
      columnWidths: [40, 60, 100],
      totalWidth: 200,
      totalHeight: 20,
      rows: [
        {
          height: 20,
          cells: [{ width: 100, height: 20, blocks: [emptyParagraphMeasure()] }],
        },
      ],
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: block.id,
      x: 0,
      y: 0,
      width: 200,
      height: 20,
      fromRow: 0,
      toRow: 1,
    };

    const rendered = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;
    const cell = findCells(rendered)[0];

    expect(cell?.dataset["columnIndex"]).toBe("2");
    expect(cell?.style["left"]).toBe("100px");
    expect(cell?.style["width"]).toBe("100px");
  });

  test("uses the same asymmetric RTL cell boxes for hit testing and selection", () => {
    const firstParagraph: ParagraphBlock = {
      kind: "paragraph",
      id: "first-p",
      pmStart: 1,
      pmEnd: 2,
      runs: [{ kind: "text", text: "a", pmStart: 1, pmEnd: 2 }],
    };
    const secondParagraph: ParagraphBlock = {
      kind: "paragraph",
      id: "second-p",
      pmStart: 3,
      pmEnd: 4,
      runs: [{ kind: "text", text: "b", pmStart: 3, pmEnd: 4 }],
    };
    const lineMeasure = (): ParagraphMeasure => ({
      kind: "paragraph",
      lines: [
        {
          fromRun: 0,
          fromChar: 0,
          toRun: 0,
          toChar: 1,
          width: 7,
          ascent: 10,
          descent: 4,
          lineHeight: 14,
        },
      ],
      totalHeight: 14,
    });
    const block: TableBlock = {
      kind: "table",
      id: "interaction-table",
      bidi: true,
      columnWidths: [100, 200],
      rows: [
        {
          id: "interaction-row",
          cells: [
            { id: "first-cell", blocks: [firstParagraph] },
            { id: "second-cell", blocks: [secondParagraph] },
          ],
        },
      ],
    };
    const measure: TableMeasure = {
      kind: "table",
      columnWidths: [100, 200],
      totalWidth: 300,
      totalHeight: 14,
      rows: [
        {
          height: 14,
          cells: [
            { width: 100, height: 14, blocks: [lineMeasure()] },
            { width: 200, height: 14, blocks: [lineMeasure()] },
          ],
        },
      ],
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: block.id,
      x: 10,
      y: 0,
      width: 300,
      height: 14,
      fromRow: 0,
      toRow: 1,
    };
    const layout: Layout = {
      pageGap: 0,
      pages: [
        {
          number: 1,
          size: { w: 320, h: 100 },
          margins: { top: 0, right: 0, bottom: 0, left: 0 },
          fragments: [fragment],
        },
      ],
    };
    const page = layout.pages[0];
    if (!page) {
      throw new Error("expected table page");
    }

    const leftHit = hitTestTableCell({ pageIndex: 0, page, pageY: 7 }, [block], [measure], {
      x: 50,
      y: 7,
    });
    const rightHit = hitTestTableCell({ pageIndex: 0, page, pageY: 7 }, [block], [measure], {
      x: 250,
      y: 7,
    });

    expect(leftHit?.colIndex).toBe(1);
    expect(leftHit?.cellLocalX).toBe(33);
    expect(rightHit?.colIndex).toBe(0);
    expect(rightHit?.cellLocalX).toBe(33);
    expect(selectionToRects(layout, [block], [measure], 3, 4)[0]?.x).toBe(17);
  });
});
