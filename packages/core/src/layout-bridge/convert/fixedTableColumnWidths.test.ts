import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../../test/property-testing";
import { parseTable } from "../../docx/tableParser";
import { parseXmlDocument } from "../../docx/xmlParser";
import { fromProseDoc } from "../../prosemirror/conversion/fromProseDoc";
import { toProseDoc } from "../../prosemirror/conversion/toProseDoc";
import type { Document } from "../../types/document";
import { toFlowBlocks } from "./toFlowBlocks";

type TableFixtureOptions = {
  grid: number[];
  rows: string[][];
  properties?: string;
  rowProperties?: string;
  layout?: "fixed" | "autofit";
};

const absoluteCell = (width: number) => `<w:tcW w:type="dxa" w:w="${width}"/>`;

const project = ({
  grid,
  rows,
  properties = "",
  rowProperties = "",
  layout = "fixed",
}: TableFixtureOptions) => {
  const rowXml = rows
    .map((cells) => {
      const cellXml = cells.map((cell) => `<w:tc><w:tcPr>${cell}</w:tcPr><w:p/></w:tc>`).join("");
      return `<w:tr><w:trPr>${rowProperties}</w:trPr>${cellXml}</w:tr>`;
    })
    .join("");
  const root = parseXmlDocument(
    `<w:tbl xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
      <w:tblPr><w:tblLayout w:type="${layout}"/>${properties}</w:tblPr>
      <w:tblGrid>${grid.map((width) => `<w:gridCol w:w="${width}"/>`).join("")}</w:tblGrid>
      ${rowXml}
    </w:tbl>`,
  );
  if (!root) {
    throw new TypeError("Expected table XML");
  }
  const table = parseTable(root, null, null, null, null, null);
  if (!table) {
    throw new TypeError("Expected parsed table");
  }
  const document = { package: { document: { content: [table] } } } satisfies Document;
  const pmDoc = toProseDoc(document);
  const block = toFlowBlocks(pmDoc).at(0);
  if (block?.kind !== "table") {
    throw new TypeError("Expected table flow block");
  }
  return { widths: block.columnWidths?.map((width) => width * 15), pmDoc, table };
};

describe("fixed table absolute cell preferences", () => {
  test(
    "resolves each column from its largest stated cell width independently of the authored grid",
    () => {
      fc.assert(
        fc.property(
          fc.array(
            fc.tuple(fc.integer({ min: 60, max: 5000 }), fc.integer({ min: 60, max: 5000 })),
            { minLength: 1, maxLength: 5 },
          ),
          (columns) => {
            const grid = columns.map(() => 3000);
            const first = columns.map(([width]) => absoluteCell(width));
            const second = columns.map(([, width]) => absoluteCell(width));
            const result = project({ grid, rows: [first, second] });
            for (const [column, width] of (result.widths ?? []).entries()) {
              const pair = columns.at(column);
              expect(pair).toBeDefined();
              if (pair) {
                expect(width).toBeCloseTo(Math.max(...pair), 8);
              }
            }
            expect(result.widths).toHaveLength(columns.length);
            const repeated = toFlowBlocks(result.pmDoc).at(0);
            expect(repeated?.kind).toBe("table");
            if (repeated?.kind === "table") {
              expect(repeated.columnWidths?.map((width) => width * 15)).toEqual(result.widths);
            }
            const saved = fromProseDoc(result.pmDoc).package.document.content.at(0);
            expect(saved?.type).toBe("table");
            if (saved?.type === "table") {
              expect(saved.columnWidths).toEqual(result.table.columnWidths);
            }
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );

  test.each([
    "",
    '<w:tcW w:type="auto" w:w="600"/>',
    '<w:tcW w:type="nil" w:w="600"/>',
    '<w:tcW w:type="pct" w:w="600"/>',
    absoluteCell(0),
    absoluteCell(-60),
  ])(
    "retains the grid when a cell does not supply a single-column absolute preference: %s",
    (cell) => {
      expect(project({ grid: [3000, 4500], rows: [[cell, absoluteCell(900)]] }).widths).toEqual([
        3000, 4500,
      ]);
    },
  );

  test("preserves the grid for horizontal and vertical spans", () => {
    const horizontal = project({
      grid: [3000, 4500],
      rows: [[`${absoluteCell(600)}<w:gridSpan w:val="2"/>`]],
    });
    expect(horizontal.widths).toEqual([3000, 4500]);
    const vertical = project({
      grid: [3000],
      rows: [
        [`${absoluteCell(600)}<w:vMerge w:val="restart"/>`],
        [`${absoluteCell(900)}<w:vMerge/>`],
      ],
    });
    expect(vertical.widths).toEqual([3000]);
  });

  test(
    "retains the grid for vertical merge declarations regardless of editor row spans",
    () => {
      fc.assert(
        fc.property(
          fc.integer({ min: 60, max: 5000 }),
          fc.array(fc.integer({ min: 60, max: 5000 }), { minLength: 1, maxLength: 5 }),
          fc.constantFrom("restart", "continue", "chain"),
          fc.boolean(),
          (gridWidth, preferences, merge, companionColumn) => {
            const grid = companionColumn ? [gridWidth, 3000] : [gridWidth];
            const rows = preferences.map((width, index) => {
              const restart = merge === "restart" || (merge === "chain" && index === 0);
              const mergeXml = restart ? '<w:vMerge w:val="restart"/>' : "<w:vMerge/>";
              const cell = `${absoluteCell(width)}${mergeXml}`;
              return companionColumn ? [cell, absoluteCell(900)] : [cell];
            });
            const result = project({ grid, rows });
            expect(result.widths).toHaveLength(grid.length);
            for (const [column, width] of (result.widths ?? []).entries()) {
              const authored = grid.at(column);
              expect(authored).toBeDefined();
              if (authored !== undefined) {
                expect(width).toBeCloseTo(authored, 8);
              }
            }
            const repeated = toFlowBlocks(result.pmDoc).at(0);
            expect(repeated?.kind).toBe("table");
            if (repeated?.kind === "table") {
              expect(repeated.columnWidths?.map((width) => width * 15)).toEqual(result.widths);
            }
          },
        ),
        propertyConfig(),
      );
    },
    propertyTestTimeout(10_000),
  );

  test("keeps the authored grid under autofit", () => {
    expect(
      project({
        grid: [3000, 4500],
        rows: [[absoluteCell(600), absoluteCell(900)]],
        layout: "autofit",
      }).widths,
    ).toEqual([3000, 4500]);
  });

  test.each(["dxa", "pct"])("retains the grid when w:tblW supplies a %s constraint", (type) => {
    expect(
      project({
        grid: [3000, 4500],
        rows: [[absoluteCell(600), absoluteCell(900)]],
        properties: `<w:tblW w:type="${type}" w:w="5000"/>`,
      }).widths,
    ).toEqual([3000, 4500]);
  });

  test.each(["auto", "nil"])("ignores the numeric residue of a %s table width", (type) => {
    expect(
      project({
        grid: [3000, 4500],
        rows: [[absoluteCell(600), absoluteCell(900)]],
        properties: `<w:tblW w:type="${type}" w:w="5000"/>`,
      }).widths,
    ).toEqual([600, 900]);
  });

  test("preserves columns omitted by a row", () => {
    expect(
      project({
        grid: [3000, 4500],
        rows: [[absoluteCell(900)]],
        rowProperties: '<w:gridBefore w:val="1"/>',
      }).widths,
    ).toEqual([3000, 900]);
    expect(
      project({
        grid: [3000, 4500],
        rows: [[absoluteCell(600)]],
        rowProperties: '<w:gridAfter w:val="1"/>',
      }).widths,
    ).toEqual([600, 4500]);
  });

  test.each([
    "<w:hidden/>",
    '<w:ins w:id="1" w:author="Reviewer"/>',
    '<w:del w:id="1" w:author="Reviewer"/>',
  ])("retains the grid when row visibility can change: %s", (rowProperties) => {
    expect(
      project({ grid: [3000, 4500], rows: [[absoluteCell(600), absoluteCell(900)]], rowProperties })
        .widths,
    ).toEqual([3000, 4500]);
  });
});
