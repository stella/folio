import { expect, test } from "bun:test";

import { convertBorderSpecToLayout } from "../../layout-bridge/convert/flowBorders";
import { layoutDocument } from "../../layout-engine/index";
import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type { TableBlock } from "../../layout-engine/types";
import { fakeDocument } from "../../layout-painter/__tests__/fakeDom";
import {
  renderTableFragment,
  TABLE_CLASS_NAMES,
} from "../../layout-painter/renderTable";
import { buildDisplayList } from "./buildDisplayList";

test("an embossed table-cell top edge paints three toned strokes", () => {
  withFakeTextMeasure(() => {
    const border = convertBorderSpecToLayout({
      style: "threeDEmboss",
      size: 24,
      color: { rgb: "000000" },
    });
    if (!border) throw new Error("Expected a visible table border");
    const table: TableBlock = {
      kind: "table",
      id: "embossed-table",
      columnWidths: [120],
      rows: [
        {
          id: "row",
          cells: [
            {
              id: "cell",
              borders: { top: border },
              padding: { top: 0, right: 0, bottom: 0, left: 0 },
              blocks: [
                {
                  kind: "paragraph",
                  id: "text",
                  runs: [{ kind: "text", text: "Cell" }],
                },
              ],
            },
          ],
        },
      ],
    };
    const pageSize = { w: 480, h: 640 };
    const margins = { top: 48, right: 48, bottom: 48, left: 48 };
    const measures = measureBlocks(
      [table],
      pageSize.w - margins.left - margins.right,
    );
    const measure = measures.at(0);
    if (measure?.kind !== "table") throw new Error("Expected a table measure");
    const layout = layoutDocument([table], measures, { pageSize, margins });
    const blockLookup = new Map([
      [String(table.id), { block: table, measure }],
    ]);
    const fragment = layout.pages.at(0)?.fragments.at(0);
    const page = buildDisplayList({ layout, blockLookup }).pages.at(0);
    if (fragment?.kind !== "table" || !page) {
      throw new Error("Expected a painted table");
    }

    const strokes = page.primitives
      .filter((primitive) => primitive.kind === "line")
      .toSorted((a, b) => a.y1Px - b.y1Px);
    const bands = strokes.map(({ y1Px, y2Px, stroke }) => {
      expect(y1Px).toBe(y2Px);
      return {
        centerPx: y1Px - fragment.y,
        thicknessPx: stroke.thicknessPx,
        color: stroke.color,
      };
    });
    // The authored three-point edge has a 1.5pt shoulder on each side.
    expect(bands).toEqual([
      { centerPx: 1, thicknessPx: 2, color: { r: 192, g: 192, b: 192, a: 1 } },
      { centerPx: 4, thicknessPx: 4, color: { r: 96, g: 96, b: 96, a: 1 } },
      { centerPx: 7, thicknessPx: 2, color: { r: 0, g: 0, b: 0, a: 1 } },
    ]);

    const row = measure.rows.at(0);
    const paragraph = row?.cells.at(0)?.blocks.at(0);
    if (!row || paragraph?.kind !== "paragraph")
      throw new Error("Expected a measured cell paragraph");
    expect(row.height).toBeCloseTo(paragraph.totalHeight + 8);
    const painted = renderTableFragment(
      fragment,
      table,
      measure,
      { pageNumber: 1, totalPages: 1, section: "body" },
      { document: fakeDocument },
    );
    const cell = painted.querySelector<HTMLElement>(
      `.${TABLE_CLASS_NAMES.cell}`,
    );
    expect(cell?.style.paddingTop).toBe("8px");
    expect(cell?.style.backgroundSize).toBe("100% 8px");
    expect(cell?.style.backgroundImage).toBe(
      "linear-gradient(to bottom, rgb(192, 192, 192) 0px 2px, rgb(96, 96, 96) 2px 6px, rgb(0, 0, 0) 6px 8px)",
    );
  });
});
