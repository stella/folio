import { describe, expect, test } from "bun:test";

import type { TableBlock, TableFragment, TableMeasure } from "../layout-engine/types";
import { renderTableFragment, TABLE_CLASS_NAMES } from "./renderTable";
import type { RenderContext } from "./renderUtils";

class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  private ownText = "";
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
  }

  get textContent(): string {
    return this.ownText + this.children.map((child) => child.textContent).join("");
  }

  set textContent(value: string) {
    this.ownText = value;
    this.children = [];
  }

  append(...children: FakeElement[]): void {
    this.children.push(...children);
  }

  appendChild(child: FakeElement): FakeElement {
    this.children.push(child);
    return child;
  }

  getContext(): null {
    return null;
  }

  querySelectorAll(): FakeElement[] {
    return [];
  }
}

const fakeDocument = {
  createElement(tagName: string): FakeElement {
    return new FakeElement(tagName);
  },
  createTextNode(text: string): FakeElement {
    const node = new FakeElement("#text");
    node.textContent = text;
    return node;
  },
} as unknown as Document;

const renderContext: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: "body",
};

function findRows(element: FakeElement): FakeElement[] {
  const matches: FakeElement[] = [];
  if (element.className.split(" ").includes(TABLE_CLASS_NAMES.row)) {
    matches.push(element);
  }
  for (const child of element.children) {
    matches.push(...findRows(child));
  }
  return matches;
}

function findByClass(element: FakeElement, className: string): FakeElement[] {
  const matches: FakeElement[] = [];
  if (element.className.split(" ").includes(className)) {
    matches.push(element);
  }
  for (const child of element.children) {
    matches.push(...findByClass(child, className));
  }
  return matches;
}

function buildHeaderContinuation(): {
  fragment: TableFragment;
  block: TableBlock;
  measure: TableMeasure;
} {
  const block: TableBlock = {
    kind: "table",
    id: "tbl",
    rows: [
      {
        id: "header",
        isHeader: true,
        cells: [
          {
            id: "header-cell",
            blocks: [
              {
                kind: "paragraph",
                id: "header-p",
                runs: [{ kind: "text", text: "Header" }],
              },
            ],
          },
        ],
      },
      {
        id: "body",
        cells: [
          {
            id: "body-cell",
            blocks: [
              {
                kind: "paragraph",
                id: "body-p",
                runs: [{ kind: "text", text: "Body" }],
              },
            ],
          },
        ],
      },
    ],
    columnWidths: [100],
  };
  const measure: TableMeasure = {
    kind: "table",
    rows: [
      {
        cells: [
          {
            blocks: [
              {
                kind: "paragraph",
                lines: [],
                totalHeight: 20,
              },
            ],
            width: 100,
            height: 20,
          },
        ],
        height: 20,
      },
      {
        cells: [
          {
            blocks: [
              {
                kind: "paragraph",
                lines: [],
                totalHeight: 100,
              },
            ],
            width: 100,
            height: 100,
          },
        ],
        height: 100,
      },
    ],
    columnWidths: [100],
    totalWidth: 100,
    totalHeight: 120,
  };
  const fragment: TableFragment = {
    kind: "table",
    blockId: "tbl",
    x: 0,
    y: 0,
    width: 100,
    height: 80,
    fromRow: 1,
    toRow: 2,
    continuesFromPrev: true,
    headerRowCount: 1,
    topClip: 40,
    bottomClip: 100,
  };
  return { fragment, block, measure };
}

describe("renderTableFragment clipped header continuations", () => {
  test("keeps repeated headers pinned while clipping the body row", () => {
    const { fragment, block, measure } = buildHeaderContinuation();

    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;

    const rows = findRows(tableEl);
    const headerRow = rows.find((row) => row.dataset["repeatedHeader"] === "true");
    const bodyRow = rows.find((row) => row.dataset["rowIndex"] === "1");
    const clipElement = tableEl.children.find(
      (child) =>
        child.style["top"] === "20px" &&
        child.style["height"] === "60px" &&
        child.style["overflow"] === "hidden",
    );

    expect(headerRow?.style["top"]).toBe("0px");
    expect(bodyRow?.style["top"]).toBe("-40px");
    expect(clipElement).toBeDefined();
  });
});

describe("renderTableFragment cell paragraph spacing", () => {
  test("reserves measured paragraph spacing inside table cells", () => {
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              blocks: [
                {
                  kind: "paragraph",
                  id: "para",
                  attrs: { spacing: { before: 6, after: 12 } },
                  runs: [{ kind: "text", text: "Cell text" }],
                },
              ],
            },
          ],
        },
      ],
      columnWidths: [100],
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        {
          cells: [
            {
              blocks: [
                {
                  kind: "paragraph",
                  lines: [],
                  totalHeight: 30,
                },
              ],
              width: 100,
              height: 30,
            },
          ],
          height: 30,
        },
      ],
      columnWidths: [100],
      totalWidth: 100,
      totalHeight: 30,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 100,
      height: 30,
      fromRow: 0,
      toRow: 1,
    };

    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;

    const paragraph = findByClass(tableEl, "layout-paragraph").at(0);

    expect(paragraph?.style["height"]).toBe("30px");
    expect(paragraph?.style["paddingTop"]).toBe("6px");
    expect(paragraph?.style["boxSizing"]).toBe("border-box");
  });
});

describe("renderTableFragment interior border ownership", () => {
  test("retains top and left edges when adjacent cells leave them unclaimed", () => {
    const border = { width: 1, style: "solid", color: "#000000" };
    const paragraph = (id: string) => ({
      kind: "paragraph" as const,
      id,
      runs: [{ kind: "text" as const, text: id }],
    });
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "top",
          cells: [
            {
              id: "top-left",
              borders: { right: border, bottom: border },
              blocks: [paragraph("top-left-p")],
            },
            { id: "top-right", blocks: [paragraph("top-right-p")] },
          ],
        },
        {
          id: "bottom",
          cells: [
            {
              id: "bottom-left",
              borders: { top: border },
              blocks: [paragraph("bottom-left-p")],
            },
            {
              id: "bottom-right",
              borders: { top: border, left: border },
              blocks: [paragraph("bottom-right-p")],
            },
          ],
        },
      ],
      columnWidths: [100, 100],
    };
    const measuredCell = () => ({
      blocks: [{ kind: "paragraph" as const, lines: [], totalHeight: 20 }],
      width: 100,
      height: 20,
    });
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        { cells: [measuredCell(), measuredCell()], height: 20 },
        { cells: [measuredCell(), measuredCell()], height: 20 },
      ],
      columnWidths: [100, 100],
      totalWidth: 200,
      totalHeight: 40,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 200,
      height: 40,
      fromRow: 0,
      toRow: 2,
    };

    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;
    const rows = findRows(tableEl);
    const cellAt = (rowIndex: string, columnIndex: string) =>
      rows
        .find((row) => row.dataset["rowIndex"] === rowIndex)
        ?.children.find((candidate) => candidate.dataset["columnIndex"] === columnIndex);

    expect(cellAt("1", "0")?.style["borderTop"]).toBeUndefined();
    expect(cellAt("1", "1")?.style["borderTop"]).toBe("1px solid #000000");
    expect(cellAt("0", "1")?.style["borderLeft"]).toBeUndefined();
    expect(cellAt("1", "1")?.style["borderLeft"]).toBe("1px solid #000000");
  });
});

describe("renderTableFragment floating cell content", () => {
  test("paints anchored images and text boxes outside the cell clip", () => {
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              padding: { top: 2, right: 3, bottom: 2, left: 4 },
              blocks: [
                {
                  kind: "paragraph",
                  id: "anchor",
                  runs: [
                    {
                      kind: "image",
                      src: "floating.png",
                      width: 80,
                      height: 60,
                      displayMode: "float",
                      wrapType: "inFront",
                      position: {
                        horizontal: { relativeTo: "column", posOffset: 95_250 },
                        vertical: { relativeTo: "paragraph", posOffset: 47_625 },
                      },
                    },
                  ],
                },
                {
                  kind: "textBox",
                  id: "inline-box",
                  width: 70,
                  height: 30,
                  content: [],
                  displayMode: "inline",
                },
                {
                  kind: "textBox",
                  id: "box",
                  width: 90,
                  height: 50,
                  content: [],
                  displayMode: "float",
                  wrapType: "inFront",
                  position: {
                    horizontal: { relativeTo: "column", posOffset: 190_500 },
                    vertical: { relativeTo: "paragraph", posOffset: 95_250 },
                  },
                },
              ],
            },
          ],
        },
      ],
      columnWidths: [100],
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        {
          cells: [
            {
              blocks: [
                { kind: "paragraph", lines: [], totalHeight: 20 },
                { kind: "textBox", width: 70, height: 30, innerMeasures: [] },
                { kind: "textBox", width: 90, height: 50, innerMeasures: [] },
              ],
              width: 100,
              height: 24,
            },
          ],
          height: 24,
        },
      ],
      columnWidths: [100],
      totalWidth: 100,
      totalHeight: 24,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 100,
      height: 24,
      fromRow: 0,
      toRow: 1,
    };

    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;
    const cell = findByClass(tableEl, TABLE_CLASS_NAMES.cell).at(0);
    const content = findByClass(tableEl, TABLE_CLASS_NAMES.cellContent).at(0);
    const imageLayer = findByClass(tableEl, "layout-cell-floating-images-layer").at(0);
    const textBoxLayer = findByClass(tableEl, "layout-cell-floating-text-boxes-layer").at(0);
    const textBox = findByClass(tableEl, "layout-textbox").at(1);

    expect(tableEl.style["overflow"]).toBe("visible");
    expect(cell?.style["overflow"]).toBe("visible");
    expect(content?.style["overflow"]).toBe("hidden");
    expect(imageLayer?.style["left"]).toBe("4px");
    expect(imageLayer?.style["top"]).toBe("2px");
    expect(imageLayer?.style["overflow"]).toBe("visible");
    expect(textBoxLayer?.style["left"]).toBe("4px");
    expect(textBoxLayer?.style["top"]).toBe("2px");
    expect(textBox?.style["left"]).toBe("20px");
    expect(Number.parseFloat(textBox?.style["top"] ?? "0")).toBeGreaterThan(50);
  });
});
