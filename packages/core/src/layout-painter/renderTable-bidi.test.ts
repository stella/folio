import { describe, expect, test } from "bun:test";

import type { TableBlock, TableFragment, TableMeasure } from "../layout-engine/types";
import { renderNestedTable, renderTableFragment, TABLE_CLASS_NAMES } from "./renderTable";
import type { RenderContext } from "./renderUtils";

// Minimal DOM stand-in (mirrors renderTableFragment.test.ts) so the painter can
// run under bun without a real document.
class FakeElement {
  className = "";
  dataset: Record<string, string> = {};
  style: Record<string, string> = {};
  children: FakeElement[] = [];
  readonly tagName: string;

  constructor(tagName: string) {
    this.tagName = tagName;
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
    node.dataset["text"] = text;
    return node;
  },
} as unknown as Document;

const renderContext: RenderContext = {
  pageNumber: 1,
  totalPages: 1,
  section: "body",
};

const COLUMN_WIDTHS = [100, 150];
const TABLE_WIDTH = 250;

const solidBorder = { width: 1, style: "solid", color: "var(--doc-border)" };
const allBorders = {
  top: solidBorder,
  right: solidBorder,
  bottom: solidBorder,
  left: solidBorder,
};

function buildTwoColumnTable(bidi: boolean): {
  fragment: TableFragment;
  block: TableBlock;
  measure: TableMeasure;
} {
  const block: TableBlock = {
    kind: "table",
    id: "tbl",
    rows: [
      {
        id: "row",
        cells: [
          {
            id: "cell-0",
            borders: allBorders,
            blocks: [
              {
                kind: "paragraph",
                id: "p0",
                runs: [{ kind: "text", text: "A" }],
              },
            ],
          },
          {
            id: "cell-1",
            borders: allBorders,
            blocks: [
              {
                kind: "paragraph",
                id: "p1",
                runs: [{ kind: "text", text: "B" }],
              },
            ],
          },
        ],
      },
    ],
    columnWidths: COLUMN_WIDTHS,
  };
  if (bidi) {
    block.bidi = true;
  }

  const measure: TableMeasure = {
    kind: "table",
    rows: [
      {
        cells: [
          {
            blocks: [{ kind: "paragraph", lines: [], totalHeight: 20 }],
            width: 100,
            height: 20,
          },
          {
            blocks: [{ kind: "paragraph", lines: [], totalHeight: 20 }],
            width: 150,
            height: 20,
          },
        ],
        height: 20,
      },
    ],
    columnWidths: COLUMN_WIDTHS,
    totalWidth: TABLE_WIDTH,
    totalHeight: 20,
  };

  const fragment: TableFragment = {
    kind: "table",
    blockId: "tbl",
    x: 0,
    y: 0,
    width: TABLE_WIDTH,
    height: 20,
    fromRow: 0,
    toRow: 1,
  };

  return { fragment, block, measure };
}

function findCells(element: FakeElement): FakeElement[] {
  const matches: FakeElement[] = [];
  if (element.className.split(" ").includes(TABLE_CLASS_NAMES.cell)) {
    matches.push(element);
  }
  for (const child of element.children) {
    matches.push(...findCells(child));
  }
  return matches;
}

function cellByColumn(cells: FakeElement[], columnIndex: string): FakeElement {
  const cell = cells.find((c) => c.dataset["columnIndex"] === columnIndex);
  if (!cell) {
    throw new Error(`Expected a cell for column ${columnIndex}`);
  }
  return cell;
}

function render(bidi: boolean): FakeElement[] {
  const { fragment, block, measure } = buildTwoColumnTable(bidi);
  const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
    document: fakeDocument,
  }) as unknown as FakeElement;
  return findCells(tableEl);
}

describe("renderTableFragment RTL column order (w:bidiVisual)", () => {
  test("uses the measured zero vertical margin when OOXML omits cell padding", () => {
    const cells = render(true);

    expect(cellByColumn(cells, "0").style["padding"]).toBe("0px 7px 0px 7px");
    expect(cellByColumn(cells, "1").style["padding"]).toBe("0px 7px 0px 7px");
  });

  test("LTR paints logical column 0 on the left", () => {
    const cells = render(false);
    expect(cellByColumn(cells, "0").style["left"]).toBe("0px");
    expect(cellByColumn(cells, "1").style["left"]).toBe("100px");
    // Outer left border lives on the leftmost cell (logical column 0).
    expect(cellByColumn(cells, "0").style["borderLeft"]).toBe("1px solid var(--doc-border)");
    expect(cellByColumn(cells, "1").style["borderLeft"]).toBeUndefined();
  });

  test("bidiVisual mirrors columns: logical column 0 paints on the right", () => {
    const cells = render(true);
    // tableWidth(250) - x - cellWidth: col0 -> 250-0-100=150, col1 -> 250-100-150=0
    expect(cellByColumn(cells, "0").style["left"]).toBe("150px");
    expect(cellByColumn(cells, "1").style["left"]).toBe("0px");
    // Logical indices are preserved for selection/editing.
    expect(cellByColumn(cells, "0").dataset["columnIndex"]).toBe("0");
    expect(cellByColumn(cells, "1").dataset["columnIndex"]).toBe("1");
    // The visual-leftmost cell (logical last column) now draws the left border.
    expect(cellByColumn(cells, "1").style["borderLeft"]).toBe("1px solid var(--doc-border)");
    expect(cellByColumn(cells, "0").style["borderLeft"]).toBeUndefined();
  });

  test("bidiVisual mirrors the internal column resize handle", () => {
    const { fragment, block, measure } = buildTwoColumnTable(true);
    const tableEl = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;

    const handle = tableEl.children.find(
      (child) =>
        child.className === TABLE_CLASS_NAMES.resizeHandle && child.dataset["columnIndex"] === "0",
    );
    // Boundary after logical column 0 is at x=100; mirrored to 250-100=150, less the 3px grab offset.
    expect(handle?.style["left"]).toBe("147px");
    // Tagged bidi so the resize path inverts the drag delta.
    expect(handle?.dataset["bidi"]).toBe("true");
  });

  test("bidiVisual edge handle resizes logical column 0 (the visual right edge)", () => {
    const renderTable = (bidi: boolean): FakeElement => {
      const { fragment, block, measure } = buildTwoColumnTable(bidi);
      return renderTableFragment(fragment, block, measure, renderContext, {
        document: fakeDocument,
      }) as unknown as FakeElement;
    };
    const ltr = renderTable(false);
    const rtl = renderTable(true);

    const edgeOf = (el: FakeElement): FakeElement | undefined =>
      el.children.find((child) => child.className === TABLE_CLASS_NAMES.tableEdgeHandleRight);

    // LTR: the right edge resizes the last logical column.
    expect(edgeOf(ltr)?.dataset["columnIndex"]).toBe("1");
    expect(edgeOf(ltr)?.dataset["bidi"]).toBeUndefined();
    // RTL: the visual right edge is logical column 0; internal handle is bidi-tagged.
    expect(edgeOf(rtl)?.dataset["columnIndex"]).toBe("0");
    const internal = rtl.children.find(
      (child) => child.className === TABLE_CLASS_NAMES.resizeHandle,
    );
    expect(internal?.dataset["bidi"]).toBe("true");
  });
});

describe("renderNestedTable RTL placement", () => {
  test.each([
    {
      justification: "left" as const,
      marginLeft: "auto",
      marginRight: "10px",
    },
    {
      justification: "right" as const,
      marginLeft: "0px",
      marginRight: undefined,
    },
  ])(
    "mirrors explicit $justification justification onto the physical edge",
    ({ justification, marginLeft, marginRight }) => {
      const { block, measure } = buildTwoColumnTable(true);
      block.indent = 10;
      block.justification = justification;

      const tableEl = renderNestedTable(
        block,
        measure,
        renderContext,
        fakeDocument,
      ) as unknown as FakeElement;

      expect(tableEl.style["marginLeft"]).toBe(marginLeft);
      expect(tableEl.style["marginRight"]).toBe(marginRight);
    },
  );

  test("anchors an authored indent from the right leading edge", () => {
    const { block, measure } = buildTwoColumnTable(true);
    block.indent = 10;

    const tableEl = renderNestedTable(
      block,
      measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    expect(tableEl.style["marginLeft"]).toBe("auto");
    expect(tableEl.style["marginRight"]).toBe("10px");
  });

  test("aligns absent and authored-zero indentation to the same table edge", () => {
    const absent = buildTwoColumnTable(true);
    const absentEl = renderNestedTable(
      absent.block,
      absent.measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    const zero = buildTwoColumnTable(true);
    zero.block.indent = 0;
    const zeroEl = renderNestedTable(
      zero.block,
      zero.measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    expect(absentEl.style["marginRight"]).toBe("0px");
    expect(zeroEl.style["marginRight"]).toBe("0px");
  });

  test("keeps non-default leading cell padding independent of an authored indent", () => {
    const { block, measure } = buildTwoColumnTable(true);
    block.indent = 10;
    block.rows[0]!.cells[0]!.padding = { top: 0, right: 12, bottom: 0, left: 7 };

    const tableEl = renderNestedTable(
      block,
      measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    expect(tableEl.style["marginRight"]).toBe("10px");
  });
});

describe("renderNestedTable LTR placement", () => {
  test("aligns absent and authored-zero indentation to the same table edge", () => {
    const absent = buildTwoColumnTable(false);
    const absentEl = renderNestedTable(
      absent.block,
      absent.measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    const zero = buildTwoColumnTable(false);
    zero.block.indent = 0;
    const zeroEl = renderNestedTable(
      zero.block,
      zero.measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    expect(absentEl.style["marginLeft"]).toBe("0px");
    expect(zeroEl.style["marginLeft"]).toBe("0px");
  });

  test("keeps non-default leading cell padding independent of an authored indent", () => {
    const { block, measure } = buildTwoColumnTable(false);
    block.indent = 10;
    block.rows[0]!.cells[0]!.padding = { top: 0, right: 7, bottom: 0, left: 12 };

    const tableEl = renderNestedTable(
      block,
      measure,
      renderContext,
      fakeDocument,
    ) as unknown as FakeElement;

    expect(tableEl.style["marginLeft"]).toBe("10px");
  });
});

describe("renderNestedTable row justification", () => {
  test.each([
    { bidi: false, firstLeft: "0px", secondLeft: "140px" },
    { bidi: true, firstLeft: "0px", secondLeft: "-140px" },
  ])(
    "positions overridden rows within the same $bidi bidi table",
    ({ bidi, firstLeft, secondLeft }) => {
      const { block, measure } = buildTwoColumnTable(bidi);
      block.indent = 10;
      block.rows[0]!.justification = "left";
      const secondRow = structuredClone(block.rows[0]!);
      secondRow.id = "row-2";
      secondRow.justification = "right";
      block.rows.push(secondRow);
      measure.rows.push(structuredClone(measure.rows[0]!));
      measure.totalHeight *= 2;

      const tableEl = renderNestedTable(
        block,
        measure,
        renderContext,
        fakeDocument,
        400,
      ) as unknown as FakeElement;
      const rows = tableEl.children.filter((child) =>
        child.className.split(" ").includes(TABLE_CLASS_NAMES.row),
      );

      expect(rows[0]?.style["left"]).toBe(firstLeft);
      expect(rows[1]?.style["left"]).toBe(secondLeft);
    },
  );
});
