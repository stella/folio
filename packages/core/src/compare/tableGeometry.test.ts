/**
 * A comparison that adds, removes or edits a table has to carry the table, not
 * a grid of strings.
 *
 * The round-trip properties the rest of the suite pins are about words: every
 * block present, in the right container, with the right text. A table is more
 * than its words — `w:tblPr`, the `w:tblGrid` widths, `w:trPr`, and per-cell
 * `w:tcPr` with its spans, merges, shading, borders and margins — and none of
 * that appears in a block projection. So these cases compare the parsed table
 * model on both sides of the round trip, and read the serialized XML where the
 * point is that a specific element was written.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import JSZip from "jszip";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import { FolioDocxReviewer } from "../ai-edits/headless";
import { projectTableGeometry } from "../internal/compare/table-geometry-program";
import { buildBodySequenceDocx, type BodyItem, type TableRow } from "./__fixtures__/body-sequence";
import { compareDocx } from "./compare";

const OPTIONS = { author: "compare", timestamp: "2024-03-01T00:00:00.000Z" } as const;

const documentXml = async (buffer: ArrayBuffer): Promise<string> => {
  const entry = (await JSZip.loadAsync(buffer)).file("word/document.xml");
  if (!entry) {
    throw new Error("The compared package has no main document part.");
  }
  return await entry.async("string");
};

type RoundTrip = {
  /** The redlined package's main document part. */
  xml: string;
  /** Table models on each side of the round trip, one line per table. */
  accepted: readonly string[];
  target: readonly string[];
  rejected: readonly string[];
  base: readonly string[];
};

/**
 * Compare, then read the table model back out of the package on both sides.
 *
 * Read from the SERIALIZED result rather than from the in-memory document, so
 * a property that survives the redline and is lost on the way to `w:tcPr` is
 * caught here rather than passing.
 */
const roundTrip = async (base: ArrayBuffer, target: ArrayBuffer): Promise<RoundTrip> => {
  const result = await compareDocx(base, target, { ...OPTIONS, mode: "bestEffort" });
  if (result.isErr()) {
    throw result.error;
  }
  const compared = await FolioDocxReviewer.fromBuffer(result.value.buffer);
  return {
    xml: await documentXml(result.value.buffer),
    accepted: projectTableGeometry(compared.storyTables({ view: "final" })),
    target: projectTableGeometry((await FolioDocxReviewer.fromBuffer(target)).storyTables()),
    rejected: projectTableGeometry(compared.storyTables({ view: "original" })),
    base: projectTableGeometry((await FolioDocxReviewer.fromBuffer(base)).storyTables()),
  };
};

const INTRO = { kind: "paragraph", text: "This agreement is made between the parties." } as const;
/**
 * A body-level paragraph after the table, so both sides of a comparison end
 * with the same one. A container may not end with a table, so a sequence that
 * does gets a blank paragraph — and a base that has one where the target does
 * not turns "the table went away" into "the table and a paragraph went away".
 */
const OUTRO = { kind: "paragraph", text: "Signed by the parties above." } as const;

/**
 * A table using every property class the comparison claims to carry: a table
 * style and look, explicit widths and a grid, borders and cell margins, a
 * repeating header row, a horizontal span, a vertical merge, cell shading and
 * alignment, and a nested table.
 */
const RICH_TABLE = {
  kind: "table",
  columnWidths: [3200, 2400, 2400],
  properties: {
    styleId: "TableGrid",
    width: { value: 8000, type: "dxa" },
    justification: "center",
    indent: 120,
    borderSize: 8,
    layout: "fixed",
    cellMargin: 80,
    look: "04A0",
  },
  rows: [
    {
      header: true,
      height: 480,
      justification: "center",
      cells: [
        { content: "Obligation", gridSpan: 2, width: 5600, shadingFill: "D9D9D9" },
        { content: "Owner", width: 2400, verticalAlign: "center" },
      ],
    },
    {
      cells: [
        {
          content: "Deliver the goods to the named place.",
          width: 3200,
          verticalMerge: "restart",
          borderSize: 4,
        },
        { content: "Before the agreed date.", width: 2400, margin: 40 },
        { content: "Supplier", width: 2400, shadingFill: "FFF2CC" },
      ],
    },
    {
      cells: [
        { content: "", width: 3200, verticalMerge: "continue" },
        {
          content: [
            { kind: "paragraph", text: "The schedule below lists the dates." },
            {
              kind: "table",
              columnWidths: [1200, 1200],
              properties: { width: { value: 2400, type: "dxa" }, borderSize: 4 },
              rows: [
                [
                  { content: "First", width: 1200 },
                  { content: "Second", width: 1200 },
                ],
              ],
            },
          ],
          width: 2400,
        },
        { content: "Buyer", width: 2400, verticalAlign: "bottom" },
      ],
    },
  ],
} as const satisfies BodyItem;

describe("a table the comparison adds", () => {
  test("carries its grid, table, row and cell properties into the redline", async () => {
    const base = await buildBodySequenceDocx([INTRO, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, RICH_TABLE, OUTRO]);

    const {
      xml,
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    // The grid, not a column count: an inserted table used to arrive with
    // `w:gridCol` elements carrying no width at all.
    expect(xml).toContain('<w:gridCol w:w="3200"/>');
    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(xml).toContain('<w:tblW w:w="8000" w:type="dxa"/>');
    expect(xml).toContain('<w:tblLook w:val="04A0"');
    expect(xml).toContain("<w:tblHeader/>");
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('w:fill="D9D9D9"');
    expect(xml).toContain('<w:vAlign w:val="center"/>');
    expect(xml).toContain('<w:tcW w:w="3200" w:type="dxa"/>');
    // Every row of an inserted table is an insertion; there is no whole-table
    // insertion element in the format.
    expect(xml).toContain("<w:ins ");

    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(expectedBase);
  });

  test("keeps a nested table nested instead of flattening it", async () => {
    const base = await buildBodySequenceDocx([INTRO, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, RICH_TABLE, OUTRO]);

    const { accepted, target: expected } = await roundTrip(base, target);

    // Two tables: the outer one and the one nested in a cell.
    expect(expected).toHaveLength(2);
    expect(accepted).toEqual(expected);
  });
});

describe("a table the comparison removes", () => {
  test("keeps its geometry under the deletion marks, so a reject restores it", async () => {
    const base = await buildBodySequenceDocx([INTRO, RICH_TABLE, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, OUTRO]);

    const { xml, rejected, base: expectedBase } = await roundTrip(base, target);

    // A removed table keeps every row, cell and property it had, under
    // deletion marks: the row carries `w:trPr/w:del` and every run in it
    // carries `w:del`, so a consumer that reads only one of the two still
    // resolves the deletion.
    expect(xml).toContain("<w:del ");
    expect(xml).toContain("<w:delText");
    expect(xml).toContain('<w:gridCol w:w="3200"/>');
    expect(xml).toContain('<w:gridSpan w:val="2"/>');
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain('<w:tblLook w:val="04A0"');

    expect(rejected).toEqual(expectedBase);
  });
});

const TWO_ROW_TABLE = {
  kind: "table",
  columnWidths: [2400, 2400],
  properties: { width: { value: 4800, type: "dxa" }, borderSize: 4 },
  rows: [
    {
      header: true,
      cells: [
        { content: "Clause", width: 2400, shadingFill: "D9D9D9" },
        { content: "Owner", width: 2400, shadingFill: "D9D9D9" },
      ],
    },
    {
      cells: [
        { content: "Delivery", width: 2400 },
        { content: "Supplier", width: 2400 },
      ],
    },
  ],
} as const satisfies BodyItem;

const withRow = (row: TableRow): BodyItem => ({
  ...TWO_ROW_TABLE,
  rows: [TWO_ROW_TABLE.rows[0], row, TWO_ROW_TABLE.rows[1]],
});

const INSERTED_ROW = {
  height: 620,
  cells: [
    { content: "Payment", width: 2400, shadingFill: "FFF2CC", verticalAlign: "bottom" },
    { content: "Buyer", width: 2400, margin: 60 },
  ],
} as const satisfies TableRow;

describe("a row the comparison adds or removes mid-table", () => {
  test("an inserted row carries its own row and cell properties", async () => {
    const base = await buildBodySequenceDocx([INTRO, TWO_ROW_TABLE, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, withRow(INSERTED_ROW), OUTRO]);

    const {
      xml,
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    expect(xml).toContain('<w:trHeight w:val="620"');
    expect(xml).toContain('w:fill="FFF2CC"');
    expect(xml).toContain('<w:vAlign w:val="bottom"/>');
    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(expectedBase);
  });

  test("a deleted row keeps its properties, so a reject restores it", async () => {
    const base = await buildBodySequenceDocx([INTRO, withRow(INSERTED_ROW), OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, TWO_ROW_TABLE, OUTRO]);

    const {
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    expect(rejected).toEqual(expectedBase);
    expect(accepted).toEqual(expected);
  });
});

describe("a table whose properties changed and whose words did not", () => {
  test("a changed cell property is written as w:tcPrChange", async () => {
    const shaded = {
      ...TWO_ROW_TABLE,
      rows: [
        TWO_ROW_TABLE.rows[0],
        {
          cells: [
            { content: "Delivery", width: 2400, shadingFill: "C6E0B4" },
            { content: "Supplier", width: 2400 },
          ],
        },
      ],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, TWO_ROW_TABLE, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, shaded, OUTRO]);

    const {
      xml,
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    expect(xml).toContain("<w:tcPrChange ");
    expect(xml).toContain('w:fill="C6E0B4"');
    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(expectedBase);
  });

  test("a changed row property is written as w:trPrChange", async () => {
    const taller = {
      ...TWO_ROW_TABLE,
      rows: [{ ...TWO_ROW_TABLE.rows[0], height: 900 }, TWO_ROW_TABLE.rows[1]],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, TWO_ROW_TABLE, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, taller, OUTRO]);

    const {
      xml,
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    expect(xml).toContain("<w:trPrChange ");
    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(expectedBase);
  });

  test("a changed table property is written as w:tblPrChange", async () => {
    const wider = {
      ...TWO_ROW_TABLE,
      properties: { width: { value: 7200, type: "dxa" }, borderSize: 4, justification: "center" },
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, TWO_ROW_TABLE, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, wider, OUTRO]);

    const {
      xml,
      accepted,
      target: expected,
      rejected,
      base: expectedBase,
    } = await roundTrip(base, target);

    expect(xml).toContain("<w:tblPrChange ");
    expect(accepted).toEqual(expected);
    expect(rejected).toEqual(expectedBase);
  });
});

const COLUMN_WIDTHS = [2400, 2000, 1600] as const;

type GeneratedCell = {
  content: string;
  width: number;
  gridSpan?: number;
  shadingFill?: string;
};

type BuildCellOptions = {
  column: number;
  gridSpan: number;
  content: string;
  shaded: boolean;
};

const buildCell = ({ column, gridSpan, content, shaded }: BuildCellOptions) => {
  const cell: GeneratedCell = {
    content,
    width: COLUMN_WIDTHS.slice(column, column + gridSpan).reduce((sum, width) => sum + width, 0),
  };
  if (gridSpan > 1) {
    cell.gridSpan = gridSpan;
  }
  if (shaded) {
    cell.shadingFill = "D9D9D9";
  }
  return cell;
};

const cellArbitrary = (column: number, spanAllowance: number): fc.Arbitrary<GeneratedCell> =>
  fc
    .record({
      content: fc.constantFrom("Alpha", "Beta", "Gamma", "Delta"),
      gridSpan: fc.integer({ min: 1, max: Math.max(1, spanAllowance) }),
      shaded: fc.boolean(),
    })
    .map(({ content, gridSpan, shaded }) => buildCell({ column, gridSpan, content, shaded }));

const buildRow = (cells: readonly GeneratedCell[], header: boolean, height: number | null) => {
  const row: { cells: readonly GeneratedCell[]; header?: boolean; height?: number } = { cells };
  if (header) {
    row.header = header;
  }
  if (height !== null) {
    row.height = height;
  }
  return row;
};

/** A row that spans the grid exactly, so the table stays well formed. */
const rowArbitrary = (): fc.Arbitrary<TableRow> =>
  fc
    .record({ header: fc.boolean(), height: fc.option(fc.integer({ min: 240, max: 900 })) })
    .chain(({ header, height }) =>
      cellArbitrary(0, COLUMN_WIDTHS.length).chain((first) => {
        const used = first.gridSpan ?? 1;
        const rest =
          used >= COLUMN_WIDTHS.length
            ? fc.constant<GeneratedCell[]>([])
            : fc.tuple(
                ...COLUMN_WIDTHS.slice(used).map((_width, index) => cellArbitrary(used + index, 1)),
              );
        return rest.map((cells) => buildRow([first, ...cells], header, height));
      }),
    );

const tableArbitrary = (): fc.Arbitrary<BodyItem> =>
  fc.array(rowArbitrary(), { minLength: 1, maxLength: 4 }).map((rows) => ({
    kind: "table",
    columnWidths: [...COLUMN_WIDTHS],
    properties: { width: { value: 6000, type: "dxa" }, borderSize: 4 },
    rows,
  }));

describe("table geometry round trip", () => {
  test("replaces a table whose paired cell spans cannot be revised in place", async () => {
    const baseTable = {
      kind: "table",
      columnWidths: [...COLUMN_WIDTHS],
      properties: { width: { value: 6000, type: "dxa" }, borderSize: 4 },
      rows: [
        {
          cells: [
            buildCell({ column: 0, gridSpan: 1, content: "Shared", shaded: false }),
            buildCell({ column: 1, gridSpan: 1, content: "Shared", shaded: false }),
            buildCell({ column: 2, gridSpan: 1, content: "Shared", shaded: false }),
          ],
        },
      ],
    } as const satisfies BodyItem;
    const targetTable = {
      ...baseTable,
      rows: [
        {
          cells: [
            buildCell({ column: 0, gridSpan: 2, content: "Shared", shaded: false }),
            buildCell({ column: 2, gridSpan: 1, content: "Shared", shaded: false }),
          ],
        },
      ],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification.status).toBe("verified");
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert"]);

    const compared = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(projectTableGeometry(compared.storyTables({ view: "final" }))).toEqual(
      projectTableGeometry((await FolioDocxReviewer.fromBuffer(target)).storyTables()),
    );
    expect(projectTableGeometry(compared.storyTables({ view: "original" }))).toEqual(
      projectTableGeometry((await FolioDocxReviewer.fromBuffer(base)).storyTables()),
    );
  });

  test("replaces a table when inserted rows carry a vertical merge", async () => {
    const kept = "The retained row identifies the same obligation on both sides.";
    const baseTable = {
      kind: "table",
      columnWidths: [2400, 2400],
      rows: [[kept, kept]],
    } as const satisfies BodyItem;
    const targetTable = {
      ...baseTable,
      rows: [
        ...baseTable.rows,
        [{ content: "New merged owner", verticalMerge: "restart" }, "Right upper cell"],
        [{ content: "", verticalMerge: "continue" }, "Right lower cell"],
      ],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification.status).toBe("verified");
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert"]);

    const compared = await FolioDocxReviewer.fromBuffer(result.value.buffer);
    expect(projectTableGeometry(compared.storyTables({ view: "final" }))).toEqual(
      projectTableGeometry((await FolioDocxReviewer.fromBuffer(target)).storyTables()),
    );
    expect(projectTableGeometry(compared.storyTables({ view: "original" }))).toEqual(
      projectTableGeometry((await FolioDocxReviewer.fromBuffer(base)).storyTables()),
    );
  });

  test("keeps a representable long row addition granular", async () => {
    const baseTable = {
      kind: "table",
      columnWidths: [2400, 2400],
      rows: [["Alpha", "Beta"]],
    } as const satisfies BodyItem;
    const targetTable = {
      ...baseTable,
      rows: [
        ...baseTable.rows,
        ["one two three four five six seven eight nine ten", "Additional owner"],
      ],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);

    const result = await compareDocx(base, target, OPTIONS);
    if (result.isErr()) {
      throw result.error;
    }
    expect(result.value.verification.status).toBe("verified");
    expect(result.value.changes.map(({ kind }) => kind)).toEqual(["table-row-insert"]);
  });

  test("does not replace a table when its template would lose a hyperlink", async () => {
    const linkedParagraph = {
      kind: "paragraph",
      text: [{ text: "Shared", href: "https://example.com" }],
    } as const satisfies BodyItem;
    const baseTable = {
      kind: "table",
      columnWidths: [2400, 2400],
      rows: [[{ content: [linkedParagraph] }, "Shared"]],
    } as const satisfies BodyItem;
    const targetTable = {
      ...baseTable,
      rows: [[{ content: [linkedParagraph], gridSpan: 2 }]],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);

    const strict = await compareDocx(base, target, OPTIONS);
    expect(strict.isErr()).toBe(true);
    if (strict.isErr()) {
      expect(strict.error._tag).toBe("CompareDocxRoundTripError");
    }

    const emitted = await compareDocx(base, target, { ...OPTIONS, mode: "bestEffort" });
    if (emitted.isErr()) {
      throw emitted.error;
    }
    expect(emitted.value.verification.status).toBe("unverified");
    expect(emitted.value.changes.map(({ kind }) => kind)).not.toContain("table-insert");
    expect(await documentXml(emitted.value.buffer)).toContain("<w:hyperlink ");
  });

  test("keeps a nested-table relocation explicitly unverified", async () => {
    const nested = {
      kind: "table",
      columnWidths: [2400],
      rows: [["Nested schedule"]],
    } as const satisfies BodyItem;
    const paragraph = (text: string) => ({ kind: "paragraph", text }) as const;
    const baseTable = {
      kind: "table",
      columnWidths: [2400, 2400],
      rows: [[{ content: [paragraph("Left"), nested, paragraph("After nested")] }, "Right"]],
    } as const satisfies BodyItem;
    const targetTable = {
      ...baseTable,
      rows: [["Left", { content: [paragraph("Right"), nested, paragraph("After nested")] }]],
    } as const satisfies BodyItem;
    const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
    const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);

    const strict = await compareDocx(base, target, OPTIONS);
    expect(strict.isErr()).toBe(true);
    if (strict.isErr()) {
      expect(strict.error._tag).toBe("CompareDocxRoundTripError");
    }

    const emitted = await compareDocx(base, target, { ...OPTIONS, mode: "bestEffort" });
    if (emitted.isErr()) {
      throw emitted.error;
    }
    expect(emitted.value.verification.status).toBe("unverified");
    if (emitted.value.verification.status === "unverified") {
      expect(emitted.value.verification.failures.map(({ cause }) => cause)).toContain("container");
    }
  });

  test(
    "accepting reproduces the target's table model and rejecting reproduces the base's",
    async () => {
      await fc.assert(
        fc.asyncProperty(tableArbitrary(), tableArbitrary(), async (baseTable, targetTable) => {
          const base = await buildBodySequenceDocx([INTRO, baseTable, OUTRO]);
          const target = await buildBodySequenceDocx([INTRO, targetTable, OUTRO]);
          const result = await compareDocx(base, target, OPTIONS);
          if (result.isErr()) {
            throw result.error;
          }
          expect(result.value.verification.status).toBe("verified");
          const compared = await FolioDocxReviewer.fromBuffer(result.value.buffer);
          expect(projectTableGeometry(compared.storyTables({ view: "final" }))).toEqual(
            projectTableGeometry((await FolioDocxReviewer.fromBuffer(target)).storyTables()),
          );
          expect(projectTableGeometry(compared.storyTables({ view: "original" }))).toEqual(
            projectTableGeometry((await FolioDocxReviewer.fromBuffer(base)).storyTables()),
          );
        }),
        propertyConfig(),
      );
    },
    propertyTestTimeout(),
  );
});
