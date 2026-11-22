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
import { buildDisplayList } from "./buildDisplayList";

test("a dot-dash table edge paints its authored thickness and four-part period", () => {
  withFakeTextMeasure(() => {
    const top = convertCellBorderSpecToLayout({ style: "dotDash", size: 12 });
    if (!top) throw new Error("Expected a visible table border");
    const table: TableBlock = {
      kind: "table",
      id: "dot-dash-table",
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
    const list = buildDisplayList({ layout, blockLookup });
    const displayPage = list.pages.at(0);
    const fragment = layout.pages.at(0)?.fragments.at(0);
    if (!displayPage || fragment?.kind !== "table")
      throw new Error("Expected a painted table");
    const lines = displayPage.primitives.filter(({ kind }) => kind === "line");
    expect(lines.map(({ stroke }) => stroke.thicknessPx)).toEqual([2]);

    const page = renderDisplayListToDom(list, { doc: fakeDocument }).at(0);
    const periods = Array.from(page?.querySelectorAll("div") ?? [])
      .map(({ style }) => style.background)
      .filter((background) =>
        background?.startsWith("repeating-linear-gradient("),
      );
    expect(periods).toHaveLength(1);

    // At 96dpi the dot-dash cycle is 14px ink, 6px clear, 6px ink, 6px clear.
    const stops = Array.from(
      (periods.at(0) ?? "").matchAll(
        /(transparent|rgb\([^)]+\)) ([\d.]+)px(?: ([\d.]+)px)?/gu,
      ),
    ).flatMap(([, color, from, to]) =>
      (to === undefined ? [from] : [from, to]).map((stop) => ({
        painted: color !== "transparent",
        at: Number(stop),
      })),
    );
    expect(stops).toEqual([
      { painted: true, at: 0 },
      { painted: true, at: 14 },
      { painted: false, at: 14 },
      { painted: false, at: 20 },
      { painted: true, at: 20 },
      { painted: true, at: 26 },
      { painted: false, at: 26 },
      { painted: false, at: 32 },
    ]);

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
    expect(cell?.style.paddingTop).toBe("2px");
    expect(cell?.style.backgroundImage).toContain(
      "#000000 0px 14px, transparent 14px 20px, #000000 20px 26px, transparent 26px 32px",
    );

    const pdfContent = paintPage({
      page: { ...displayPage, primitives: lines },
      fonts: new Map(),
    });
    expect(
      pdfContent
        .flatMap((part) => (part.kind === "literal" ? [part.text] : []))
        .join(""),
    ).toContain("[14 6 6 6] 0 d");
  });
});
