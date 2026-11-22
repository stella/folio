import { expect, test } from "bun:test";

import { convertCellBorderSpecToLayout } from "../../layout-bridge/convert/flowBorders";
import { layoutDocument } from "../../layout-engine/index";
import { withFakeTextMeasure } from "../../layout-engine/measure/__tests__/fakeTextMeasure";
import { measureBlocks } from "../../layout-engine/measure/measureBlocks";
import type { TableBlock } from "../../layout-engine/types";
import { fakeDocument } from "../../layout-painter/__tests__/fakeDom";
import {
  renderTableFragment,
  TABLE_CLASS_NAMES,
} from "../../layout-painter/renderTable";
import { paintPage } from "../../pdf/paint";
import { renderDisplayListToDom } from "../dom/renderDisplayListToDom";
import { strokeCrossExtentPx } from "../primitives";
import { buildDisplayList } from "./buildDisplayList";

test("a hatched table edge reserves a fixed four-pixel band across authored widths", () => {
  withFakeTextMeasure(() => {
    const pageSize = { w: 480, h: 640 };
    const margins = { top: 48, right: 48, bottom: 48, left: 48 };
    const contentWidth = pageSize.w - margins.left - margins.right;
    const cases = [
      { size: 4, side: "top" },
      { size: 48, side: "top" },
      { size: 48, side: "right" },
    ] as const;
    const bandWidths = cases.map(({ size, side }) => {
      const edge = convertCellBorderSpecToLayout({
        style: "dashDotStroked",
        size,
      });
      if (!edge) throw new Error("Expected a visible table border");
      const table: TableBlock = {
        kind: "table",
        id: "hatched-table",
        columnWidths: [120],
        rows: [
          {
            id: "row",
            cells: [
              {
                id: "cell",
                borders: { [side]: edge },
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
      const row = measure.rows.at(0);
      const paragraph = row?.cells.at(0)?.blocks.at(0);
      if (!row || paragraph?.kind !== "paragraph")
        throw new Error("Expected a measured cell paragraph");
      const layout = layoutDocument([table], measures, { pageSize, margins });
      const blockLookup = new Map([
        [String(table.id), { block: table, measure }],
      ]);
      const list = buildDisplayList({ layout, blockLookup });
      const page = list.pages.at(0);
      const fragment = layout.pages.at(0)?.fragments.at(0);
      if (!page || fragment?.kind !== "table")
        throw new Error("Expected a painted table");
      const lines = page.primitives.filter(({ kind }) => kind === "line");
      const line = lines.at(0);
      if (!line) throw new Error("Expected a table edge");

      expect(line.hatch).toEqual({
        pitchPx: 0.96,
        lineWidthPx: 0.32,
        slots: [true, true, true, false, true, false, true, true],
      });

      const legacy = renderTableFragment(
        fragment,
        table,
        measure,
        { pageNumber: 1, totalPages: 1, section: "body" },
        { document: fakeDocument },
      );
      const cell = legacy.querySelector<HTMLElement>(
        `.${TABLE_CLASS_NAMES.cell}`,
      );
      expect(cell?.style.backgroundImage).toContain(
        "repeating-linear-gradient(135deg, #000000 0px 0.32px",
      );
      expect(cell?.style.backgroundSize).toBe(
        side === "top" ? "100% 4px" : "4px 100%",
      );
      const dom = renderDisplayListToDom(list, { doc: fakeDocument }).at(0);
      const gradients = Array.from(dom?.querySelectorAll("div") ?? [])
        .map(({ style }) => style.background)
        .filter((background) =>
          background?.startsWith("repeating-linear-gradient(135deg,"),
        );
      expect(gradients).toEqual([
        cell?.style.backgroundImage.replaceAll("#000000", "rgb(0, 0, 0)"),
      ]);

      const pdf = paintPage({
        page: { ...page, primitives: lines },
        fonts: new Map(),
      })
        .flatMap((part) => (part.kind === "literal" ? [part.text] : []))
        .join("");
      expect(pdf).toContain("W n\n");
      expect(pdf).toContain("0.32 w\n");
      const readPoints = (operator: string) =>
        Array.from(
          pdf.matchAll(
            new RegExp(`^(-?[0-9.]+) (-?[0-9.]+) ${operator}$`, "gmu"),
          ),
          ([, x, y]) => ({
            x: Number(x),
            y: Number(y),
          }),
        );
      const starts = readPoints("m");
      const ends = readPoints("l");
      expect(starts.length).toBeGreaterThanOrEqual(6);
      expect(ends.length).toBe(starts.length);
      for (const [index, start] of starts.entries()) {
        const end = ends.at(index);
        if (!end) throw new Error("Expected a diagonal PDF mark");
        expect(end.x).toBeGreaterThan(start.x);
        expect(end.x - start.x).toBeCloseTo(start.y - end.y, 2);
      }
      const first = starts.at(0);
      if (!first) throw new Error("Expected a hatch period");
      for (const [index, slot] of [0, 1, 2, 4, 6, 7].entries()) {
        const point = starts.at(index);
        if (!point) throw new Error("Expected an ink slot");
        expect(point.x - first.x + point.y - first.y).toBeCloseTo(
          slot * 0.96,
          2,
        );
      }
      return {
        reservedPx: row.height - paragraph.totalHeight,
        strokes: lines.map(({ stroke }) => strokeCrossExtentPx(stroke)),
      };
    });

    expect(bandWidths).toEqual([
      { reservedPx: 4, strokes: [4] },
      { reservedPx: 4, strokes: [4] },
      { reservedPx: 0, strokes: [4] },
    ]);
  });
});
