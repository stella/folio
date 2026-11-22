import { expect, test } from "bun:test";

import { convertCellBorderSpecToLayout } from "../../layout-bridge/convert/flowBorders";
import { layoutDocument } from "../../layout-engine/index";
import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import {
  tableCellBorderBands,
  tableCellBorderWidth,
} from "../../layout-engine/measure/tableBorderBands";
import type { TableBlock } from "../../layout-engine/types";
import { fakeDocument } from "../../layout-painter/__tests__/fakeDom";
import { computePageFingerprint } from "../../layout-painter/renderPage";
import {
  renderTableFragment,
  TABLE_CLASS_NAMES,
} from "../../layout-painter/renderTable";
import { buildDisplayList } from "./buildDisplayList";

const cases = [
  {
    style: "thinThickLargeGap",
    bands: [
      { offsetPx: 0, widthPx: 2 },
      { offsetPx: 6, widthPx: 1 },
    ],
  },
  {
    style: "thickThinLargeGap",
    bands: [
      { offsetPx: 0, widthPx: 1 },
      { offsetPx: 5, widthPx: 2 },
    ],
  },
] as const;

test("large-gap compound table borders reverse the two stroke widths", () => {
  withFakeTextMeasure(() => {
    const pageSize = { w: 480, h: 640 };
    const margins = { top: 48, right: 48, bottom: 48, left: 48 };
    const contentWidth = pageSize.w - margins.left - margins.right;
    const results = cases.map(({ style }) => {
      const top = convertCellBorderSpecToLayout({ style, size: 24 });
      if (!top) throw new Error("Expected a visible top border");
      const table: TableBlock = {
        kind: "table",
        id: "compound-table",
        columnWidths: [120],
        rows: [
          {
            id: "row",
            cells: [
              {
                id: "cell",
                borders: { top },
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
      const measures = measureBlocks([table], contentWidth);
      const measure = measures.at(0);
      if (measure?.kind !== "table")
        throw new Error("Expected a table measure");
      const layout = layoutDocument([table], measures, { pageSize, margins });
      const blockLookup = new Map([
        [String(table.id), { block: table, measure }],
      ]);
      const layoutPage = layout.pages.at(0);
      const fragment = layoutPage?.fragments.at(0);
      const page = buildDisplayList({ layout, blockLookup }).pages.at(0);
      if (!layoutPage || fragment?.kind !== "table" || !page)
        throw new Error("Expected a painted table");

      const row = measure.rows.at(0);
      const paragraph = row?.cells.at(0)?.blocks.at(0);
      if (!row || paragraph?.kind !== "paragraph")
        throw new Error("Expected a measured cell paragraph");
      expect(row.height).toBeCloseTo(paragraph.totalHeight + 7);

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
      expect(cell?.style.paddingTop).toBe("7px");
      expect(cell?.style.backgroundSize).toBe("100% 7px");
      return {
        lines: page.primitives,
        fingerprint: computePageFingerprint(layoutPage, blockLookup),
        gradient: cell?.style.backgroundImage,
      };
    });
    const paintedBands = results.map(
      ({ lines: primitives, gradient }, index) => {
        const firstWidth = cases.at(index)?.bands[0].widthPx;
        if (firstWidth === undefined)
          throw new Error("Expected a border fixture");
        expect(gradient).toContain(
          `transparent ${firstWidth}px ${firstWidth + 4}px`,
        );
        const lines = primitives
          .filter((primitive) => primitive.kind === "line")
          .toSorted((a, b) => a.y1Px - b.y1Px);
        const first = lines.at(0);
        if (!first) throw new Error("Expected at least one border line");
        const startPx = first.y1Px - first.stroke.thicknessPx / 2;
        return lines.map(({ y1Px, y2Px, stroke }) => {
          expect(y1Px).toBe(y2Px);
          return {
            offsetPx: y1Px - stroke.thicknessPx / 2 - startPx,
            widthPx: stroke.thicknessPx,
          };
        });
      },
    );
    // Each style keeps the same four-pixel gap and seven-pixel total band.
    expect(paintedBands).toEqual(cases.map(({ bands }) => bands));
    expect(new Set(results.map(({ fingerprint }) => fingerprint)).size).toBe(
      cases.length,
    );
  });
});

test.each([
  ["thinThickSmallGap", 2 / 3, 1],
  ["thickThinSmallGap", 1, 2 / 3],
] as const)(
  "retains the small-gap band and authored width for %s",
  (style, first, last) => {
    const compound = convertCellBorderSpecToLayout({ style, size: 4 });
    const single = convertCellBorderSpecToLayout({ style: "single", size: 12 });
    if (!compound || !single) throw new Error("Expected two visible borders");

    expect(compound.width).toBeCloseTo(2 / 3);
    expect(single.width).toBe(2);
    expect(tableCellBorderWidth(compound)).toBeCloseTo(8 / 3);
    const bands = tableCellBorderBands(compound)?.bands;
    expect(bands).toHaveLength(2);
    expect(bands?.at(0)?.offsetPx).toBe(0);
    expect(bands?.at(0)?.widthPx).toBeCloseTo(first);
    expect(bands?.at(1)?.offsetPx).toBeCloseTo(first + 1);
    expect(bands?.at(1)?.widthPx).toBeCloseTo(last);
  },
);
