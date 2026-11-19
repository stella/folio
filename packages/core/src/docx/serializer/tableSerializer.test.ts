import { describe, expect, test } from "bun:test";

import type { Table } from "../../types/document";
import { serializeTable } from "./tableSerializer";

const serializeParagraph = (): string => "<w:p/>";

const cell = (gridSpan?: number) =>
  ({
    type: "tableCell",
    ...(gridSpan === undefined ? {} : { formatting: { gridSpan } }),
    content: [{ type: "paragraph", content: [] }],
  }) as const satisfies Table["rows"][number]["cells"][number];

// `w:tbl` is `w:tblPr, w:tblGrid, (rows)*`: both are required and both precede
// every row. A table the model carries no properties or measured widths for —
// the shape a tracked table insertion produces — used to open with its first
// `w:tr`, which the content model has no place for.
describe("serializeTable required children", () => {
  test("opens with the properties and grid even when the model carries neither", () => {
    const table: Table = {
      type: "table",
      rows: [
        { type: "tableRow", cells: [cell(), cell()] },
        { type: "tableRow", cells: [cell(), cell()] },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(
      xml.startsWith(
        "<w:tbl><w:tblPr/><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>",
      ),
    ).toBe(true);
  });

  test("declares a grid column for every column the widest row spans", () => {
    const table: Table = {
      type: "table",
      rows: [
        { type: "tableRow", cells: [cell(3)] },
        { type: "tableRow", cells: [cell(), cell(), cell()] },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml.match(/<w:gridCol\b/gu)).toHaveLength(3);
  });

  test("keeps authored column widths when the model has them", () => {
    const table: Table = {
      type: "table",
      columnWidths: [2000, 3000],
      rows: [{ type: "tableRow", cells: [cell(), cell()] }],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain(
      '<w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>',
    );
  });
});

// A grid captured from the source is written back verbatim only while it
// still declares the grid the rows cover: an edit that adds a column without
// measuring it leaves the widths unchanged (none) but not the grid.
describe("serializeTable grid source", () => {
  const W =
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"';
  const twoColumns = `<w:tblGrid ${W}><w:gridCol/><w:gridCol/></w:tblGrid>`;

  test("keeps the captured grid while it still describes the table", () => {
    const table: Table = {
      type: "table",
      formatting: { gridSourceXml: twoColumns },
      rows: [{ type: "tableRow", cells: [cell(), cell()] }],
    };

    expect(serializeTable(table, serializeParagraph)).toContain(twoColumns);
  });

  test("rebuilds it once the rows cover another column and a grid change", () => {
    const table: Table = {
      type: "table",
      formatting: {
        gridSourceXml: twoColumns,
        gridChange: { id: 4, columnWidths: [undefined, undefined] },
      },
      rows: [{ type: "tableRow", cells: [cell(), cell(), cell()] }],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml).toContain(
      '<w:tblGrid><w:gridCol/><w:gridCol/><w:gridCol/><w:tblGridChange w:id="4"><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid></w:tblGridChange></w:tblGrid>',
    );
  });

  test("declares the columns a row skips before and after its cells", () => {
    const table: Table = {
      type: "table",
      rows: [
        {
          type: "tableRow",
          formatting: { gridBefore: 1, gridAfter: 1 },
          cells: [cell()],
        },
        {
          type: "tableRow",
          formatting: { gridBefore: 2 },
          cells: [cell()],
        },
      ],
    };

    const xml = serializeTable(table, serializeParagraph);

    expect(xml.match(/<w:gridCol\b/gu)).toHaveLength(3);
  });
});
