import { describe, expect, test } from "bun:test";

import { withFakeTextMeasure } from "../layout-engine/measure/__tests__/fakeTextMeasure";
import type { ImageRun, TableBlock, TableFragment, TableMeasure } from "../layout-engine/types";
import type { PageGeometry } from "./anchoredImagePosition";
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

const pageGeometry: PageGeometry = {
  pageWidth: 816,
  pageHeight: 1056,
  marginLeft: 96,
  marginTop: 96,
  marginRight: 96,
  marginBottom: 96,
  contentWidth: 624,
  contentHeight: 864,
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

describe("renderTableFragment split-row cell content", () => {
  test("clips ordinary cell content to each intersecting row slice", () => {
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              padding: { top: 0, right: 0, bottom: 0, left: 0 },
              blocks: [
                {
                  kind: "paragraph",
                  id: "content",
                  runs: [{ kind: "text", text: "Tall cell content" }],
                },
                {
                  kind: "paragraph",
                  id: "floating-anchor",
                  runs: [
                    {
                      kind: "image",
                      src: "floating.png",
                      width: 20,
                      height: 20,
                      displayMode: "float",
                      wrapType: "inFront",
                    },
                  ],
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
                { kind: "paragraph", lines: [], totalHeight: 100 },
                { kind: "paragraph", lines: [], totalHeight: 0 },
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
      totalHeight: 100,
    };
    const fragments: TableFragment[] = [
      {
        kind: "table",
        blockId: "tbl",
        x: 0,
        y: 0,
        width: 100,
        height: 30,
        fromRow: 0,
        toRow: 1,
        bottomClip: 30,
        continuesOnNext: true,
      },
      {
        kind: "table",
        blockId: "tbl",
        x: 0,
        y: 0,
        width: 100,
        height: 40,
        fromRow: 0,
        toRow: 1,
        topClip: 30,
        bottomClip: 70,
        continuesFromPrev: true,
        continuesOnNext: true,
      },
      {
        kind: "table",
        blockId: "tbl",
        x: 0,
        y: 0,
        width: 100,
        height: 30,
        fromRow: 0,
        toRow: 1,
        topClip: 70,
        continuesFromPrev: true,
      },
    ];

    let clipPaths: (string | undefined)[] = [];
    withFakeTextMeasure(() => {
      clipPaths = fragments.map((fragment) => {
        const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
          document: fakeDocument,
        }) as unknown as FakeElement;
        const content = findByClass(tableEl, TABLE_CLASS_NAMES.cellContent).at(0);
        const floatingLayer = findByClass(tableEl, "layout-cell-floating-images-layer").at(0);

        expect(tableEl.style["overflow"]).toBe("visible");
        expect(content?.style["height"]).toBe("100px");
        expect(content?.style["overflow"]).toBe("hidden");
        expect(floatingLayer?.style["clipPath"]).toBeUndefined();

        return content?.style["clipPath"];
      });
    });

    expect(clipPaths).toEqual([
      "inset(0px 0 70px 0)",
      "inset(30px 0 30px 0)",
      "inset(70px 0 0px 0)",
    ]);
  });

  test("applies a bottom slice only to the final row of a multi-row fragment", () => {
    const paragraph = (id: string) => ({
      kind: "paragraph" as const,
      id,
      runs: [
        { kind: "text" as const, text: "Cell content" },
        {
          kind: "image" as const,
          src: "floating.png",
          width: 20,
          height: 20,
          displayMode: "float" as const,
          wrapType: "inFront" as const,
        },
      ],
    });
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row-1",
          cells: [
            {
              id: "cell-1",
              padding: { top: 0, right: 0, bottom: 0, left: 0 },
              blocks: [paragraph("paragraph-1")],
            },
          ],
        },
        {
          id: "row-2",
          cells: [
            {
              id: "cell-2",
              padding: { top: 0, right: 0, bottom: 0, left: 0 },
              blocks: [paragraph("paragraph-2")],
            },
          ],
        },
      ],
      columnWidths: [100],
    };
    const measuredRow = {
      cells: [
        {
          blocks: [{ kind: "paragraph" as const, lines: [], totalHeight: 100 }],
          width: 100,
          height: 100,
        },
      ],
      height: 100,
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [measuredRow, measuredRow],
      columnWidths: [100],
      totalWidth: 100,
      totalHeight: 200,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 100,
      height: 130,
      fromRow: 0,
      toRow: 2,
      bottomClip: 30,
      continuesOnNext: true,
    };

    withFakeTextMeasure(() => {
      const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
        document: fakeDocument,
      }) as unknown as FakeElement;
      const contents = findByClass(tableEl, TABLE_CLASS_NAMES.cellContent);

      expect(contents.at(0)?.style["clipPath"]).toBeUndefined();
      expect(contents.at(1)?.style["clipPath"]).toBe("inset(0px 0 70px 0)");
    });
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

describe("renderTableFragment bottom-to-top cell text", () => {
  test("rotates and centers content within the cell height", () => {
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              textDirection: "btLr",
              padding: { top: 2, right: 3, bottom: 4, left: 5 },
              blocks: [
                {
                  kind: "paragraph",
                  id: "para",
                  runs: [{ kind: "text", text: "Rotated cell" }],
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
              blocks: [{ kind: "paragraph", lines: [], totalHeight: 20 }],
              width: 100,
              height: 80,
            },
          ],
          height: 80,
        },
      ],
      columnWidths: [100],
      totalWidth: 100,
      totalHeight: 80,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 100,
      height: 80,
      fromRow: 0,
      toRow: 1,
    };

    withFakeTextMeasure(() => {
      const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
        document: fakeDocument,
      }) as unknown as FakeElement;
      const content = findByClass(tableEl, TABLE_CLASS_NAMES.cellContent).at(0);

      expect(content?.style["position"]).toBe("absolute");
      expect(content?.style["left"]).toBe("50%");
      expect(content?.style["top"]).toBe("50%");
      expect(content?.style["width"]).toBe("74px");
      expect(content?.style["transform"]).toBe("translate(-50%, -50%) rotate(-90deg)");
    });
  });
});

describe("renderTableFragment interior border ownership", () => {
  test("paints a zero-width styled edge as one visible hairline", () => {
    const hairline = { width: 0, style: "solid", color: "#000000" };
    const paragraph = {
      kind: "paragraph" as const,
      id: "p",
      runs: [{ kind: "text" as const, text: "content" }],
    };
    const block: TableBlock = {
      kind: "table",
      id: "tbl",
      rows: [
        {
          id: "row",
          cells: [{ id: "cell", borders: { top: hairline }, blocks: [paragraph] }],
        },
      ],
      columnWidths: [100],
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        {
          cells: [
            { blocks: [{ kind: "paragraph", lines: [], totalHeight: 20 }], width: 100, height: 20 },
          ],
          height: 20,
        },
      ],
      columnWidths: [100],
      totalWidth: 100,
      totalHeight: 20,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "tbl",
      x: 0,
      y: 0,
      width: 100,
      height: 20,
      fromRow: 0,
      toRow: 1,
    };

    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;
    const cell = findRows(tableEl).at(0)?.children.at(0);

    expect(cell?.style["borderTop"]).toBe("1px solid #000000");
  });

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

describe("renderTableFragment cell floating images", () => {
  const renderFloatingImage = (
    image: ImageRun,
    options: { pageGeometry?: PageGeometry } = {},
  ): FakeElement => {
    const block: TableBlock = {
      kind: "table",
      id: "table",
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              padding: { top: 10, right: 10, bottom: 10, left: 10 },
              blocks: [{ kind: "paragraph", id: "anchor", runs: [image] }],
            },
          ],
        },
      ],
      columnWidths: [220],
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        {
          cells: [
            {
              blocks: [{ kind: "paragraph", lines: [], totalHeight: 40 }],
              width: 220,
              height: 80,
            },
          ],
          height: 80,
        },
      ],
      columnWidths: [220],
      totalWidth: 220,
      totalHeight: 80,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "table",
      x: 196,
      y: 196,
      width: 220,
      height: 80,
      fromRow: 0,
      toRow: 1,
    };

    let rendered = new FakeElement("div");
    withFakeTextMeasure(() => {
      rendered = renderTableFragment(fragment, block, measure, renderContext, {
        document: fakeDocument,
        ...options,
      }) as unknown as FakeElement;
    });
    return rendered;
  };

  test("resolves negative margin-relative offsets against page geometry", () => {
    const table = renderFloatingImage(
      {
        kind: "image",
        src: "floating.png",
        width: 80,
        height: 30,
        displayMode: "float",
        wrapType: "inFront",
        position: {
          horizontal: { relativeTo: "margin", posOffset: -457_200 },
          vertical: { relativeTo: "paragraph", posOffset: 0 },
        },
      },
      { pageGeometry },
    );

    const image = findByClass(table, "layout-cell-floating-image").at(0);
    // The cell content starts 110px into the content area. The authored
    // margin-relative offset is -48px, so the cell-local result is -158px.
    expect(image?.style["left"]).toBe("-158px");
    expect(image?.style["top"]).toBe("0px");
  });

  test("applies crop and opacity attributes in the cell floating layer", () => {
    const table = renderFloatingImage({
      kind: "image",
      src: "cropped.png",
      width: 120,
      height: 40,
      displayMode: "float",
      wrapType: "inFront",
      cropLeft: 0.2,
      cropRight: 0.1,
      opacity: 0.6,
    });

    const container = findByClass(table, "layout-cell-floating-image").at(0);
    const image = container?.children.at(0);
    expect(container?.style["width"]).toBe("120px");
    expect(container?.style["height"]).toBe("40px");
    expect(container?.style["overflow"]).toBe("hidden");
    expect(Number.parseFloat(image?.style["width"] ?? "")).toBeCloseTo((1 / 0.7) * 100, 6);
    expect(image?.style["opacity"]).toBe("0.6");
    expect(image?.style["objectFit"]).toBe("fill");
  });

  test("does not clip an uncropped image that only has opacity", () => {
    const table = renderFloatingImage({
      kind: "image",
      src: "transparent.png",
      width: 120,
      height: 40,
      displayMode: "float",
      wrapType: "inFront",
      opacity: 0.6,
      transform: "rotate(10deg)",
    });

    const container = findByClass(table, "layout-cell-floating-image").at(0);
    const image = container?.children.at(0);
    expect(container?.style["overflow"]).toBeUndefined();
    expect(container?.style["width"]).toBeUndefined();
    expect(image?.style["opacity"]).toBe("0.6");
  });
});
