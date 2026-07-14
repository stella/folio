import { describe, expect, test } from "bun:test";

import type { TableBlock, TableFragment, TableMeasure } from "../layout-engine/types";
import { renderTableFragment } from "./renderTable";
import type { RenderContext } from "./renderUtils";

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

  getContext(): null {
    return null;
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

const findByClass = (element: FakeElement, className: string): FakeElement[] => {
  const matches = element.className.split(" ").includes(className) ? [element] : [];
  for (const child of element.children) {
    matches.push(...findByClass(child, className));
  }
  return matches;
};

describe("table cell diagonal border painting", () => {
  test("paints both directions across the measured cell box", () => {
    const block: TableBlock = {
      kind: "table",
      id: "table",
      columnWidths: [100],
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              blocks: [{ kind: "paragraph", id: "paragraph", runs: [] }],
              borders: {
                topLeftToBottomRight: { color: "#123456", style: "dashed", width: 2 },
                topRightToBottomLeft: { color: "#654321", style: "double", width: 3 },
              },
            },
          ],
        },
      ],
    };
    const measure: TableMeasure = {
      kind: "table",
      rows: [
        {
          cells: [
            {
              blocks: [{ kind: "paragraph", lines: [], totalHeight: 40 }],
              height: 40,
              width: 100,
            },
          ],
          height: 40,
        },
      ],
      columnWidths: [100],
      totalHeight: 40,
      totalWidth: 100,
    };
    const fragment: TableFragment = {
      kind: "table",
      blockId: "table",
      x: 0,
      y: 0,
      width: 100,
      height: 40,
      fromRow: 0,
      toRow: 1,
    };

    const rendered = renderTableFragment(fragment, block, measure, renderContext, {
      document: fakeDocument,
    }) as unknown as FakeElement;
    const diagonals = findByClass(rendered, "layout-table-cell-diagonal-border");

    expect(diagonals).toHaveLength(2);
    expect(diagonals.at(0)?.dataset["direction"]).toBe("top-left-to-bottom-right");
    expect(diagonals.at(0)?.style["width"]).toBe(`${Math.hypot(100, 40)}px`);
    expect(diagonals.at(0)?.style["backgroundImage"]).toContain("repeating-linear-gradient");
    expect(diagonals.at(1)?.dataset["direction"]).toBe("top-right-to-bottom-left");
    expect(diagonals.at(1)?.style["top"]).toBe("38.5px");
    expect(diagonals.at(1)?.style["backgroundImage"]).toContain("linear-gradient");
  });
});
