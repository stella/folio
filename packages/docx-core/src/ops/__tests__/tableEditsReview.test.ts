import { describe, expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
  TableCell,
  TableMeasurement,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { compileEditorIntent } from "../editorIntent";
import { locateTableRow } from "../tableLocation";
import { tableEditParagraphDemand } from "../tableEdits";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  type DocumentOp,
  type TableIntentOperation,
} from "../types";

const id = (value: number): string => value.toString(16).toUpperCase().padStart(8, "0");
const paragraph = (value: number, text = "cell"): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  content: text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (value: number, text = "cell", width?: TableMeasurement): TableCell => ({
  type: "tableCell",
  ...(width === undefined ? {} : { formatting: { width } }),
  content: [paragraph(value, text)],
});
const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [paragraph(1, "before"), table, paragraph(2, "after")],
    },
  },
});
const appliedDocument = (document: Document, op: DocumentOp): Document => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value.document;
};
const expectGridCoverage = (table: Table): void => {
  const width = table.columnWidths?.length;
  if (width === undefined) throw new Error("The grid has explicit widths.");
  for (const row of table.rows) {
    const coverage =
      (row.formatting?.gridBefore ?? 0) +
      row.cells.reduce(
        (columns, current) =>
          columns +
          (current.structuralChange?.type === "tableCellDeletion"
            ? 0
            : (current.formatting?.gridSpan ?? 1)),
        0,
      ) +
      (row.formatting?.gridAfter ?? 0);
    expect(coverage).toBe(width);
  }
};
const tableFrom = (document: Document): Table => {
  const table = document.package.document.content.at(1);
  if (table?.type !== "table") throw new Error("The fixture table remains in place.");
  return table;
};
const paragraphIds = (document: Document): string[] => {
  const collect = (blocks: readonly BlockContent[]): string[] => {
    const ids: string[] = [];
    for (const block of blocks) {
      if (block.type === "paragraph") {
        if (block.paraId === undefined) throw new Error("Fixture paragraph id missing.");
        ids.push(block.paraId);
        continue;
      }
      if (block.type !== "table") continue;
      for (const row of block.rows) {
        for (const current of row.cells) {
          for (const paraId of collect(current.content)) ids.push(paraId);
        }
      }
    }
    return ids;
  };
  return collect(document.package.document.content);
};
const tableAddress = (blockId: string) => ({ story: OP_STORIES.MAIN, blockId });

describe("table edit review fixes", () => {
  test.each(["dxa", "pct"] as const)("split divides %s width exactly across unit cells", (type) => {
    const originalWidth = { type, value: 101 } satisfies TableMeasurement;
    const baseline = documentOf({
      type: "table",
      columnWidths: [130, 170],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              ...cell(100, "span", originalWidth),
              formatting: { width: originalWidth, gridSpan: 2 },
            },
          ],
        },
      ],
    });
    const after = appliedDocument(baseline, {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: [id(101)],
    });
    const table = tableFrom(after);
    const widths = table.rows.at(0)?.cells.map(({ formatting }) => formatting?.width);
    expect(table.rows.at(0)?.cells).toHaveLength(2);
    expect(widths).toEqual([
      { type, value: 51 },
      { type, value: 50 },
    ]);
    expect(widths?.reduce((total, width) => total + (width?.value ?? 0), 0)).toBe(
      originalWidth.value,
    );
    expectGridCoverage(table);
  });

  test.each(["dxa", "pct"] as const)("merge sums compatible top-row %s widths", (type) => {
    const table: Table = {
      type: "table",
      columnWidths: [80, 120],
      rows: [
        {
          type: "tableRow",
          cells: [cell(100, "left", { type, value: 40 }), cell(101, "right", { type, value: 61 })],
        },
      ],
    };
    const after = appliedDocument(documentOf(table), {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      top: 0,
      bottom: 1,
      left: 0,
      right: 2,
      newBlockIds: [],
    });
    const merged = tableFrom(after);
    expect(merged.rows.at(0)?.cells).toHaveLength(1);
    expect(merged.rows.at(0)?.cells.at(0)?.formatting?.width).toEqual({ type, value: 101 });
    expectGridCoverage(merged);
  });

  test("merge preserves the leading width when selected widths are incompatible", () => {
    const after = appliedDocument(
      documentOf({
        type: "table",
        columnWidths: [80, 120],
        rows: [
          {
            type: "tableRow",
            cells: [
              cell(100, "left", { type: "dxa", value: 40 }),
              cell(101, "right", { type: "pct", value: 61 }),
            ],
          },
        ],
      }),
      {
        ...tableAddress(id(100)),
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        top: 0,
        bottom: 1,
        left: 0,
        right: 2,
        newBlockIds: [],
      },
    );
    expect(tableFrom(after).rows.at(0)?.cells.at(0)?.formatting?.width).toEqual({
      type: "dxa",
      value: 40,
    });
  });

  test("split refuses fractional OOXML width units atomically", () => {
    const baseline = documentOf({
      type: "table",
      columnWidths: [130, 170],
      rows: [
        {
          type: "tableRow",
          cells: [
            {
              ...cell(100, "span", { type: "dxa", value: 101.5 }),
              formatting: { width: { type: "dxa", value: 101.5 }, gridSpan: 2 },
            },
          ],
        },
      ],
    });
    const result = applyDocumentOp(baseline, {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: [id(101)],
    });
    expect(result.isErr() && result.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    expect(baseline).toEqual(
      documentOf({
        type: "table",
        columnWidths: [130, 170],
        rows: [
          {
            type: "tableRow",
            cells: [
              {
                ...cell(100, "span", { type: "dxa", value: 101.5 }),
                formatting: { width: { type: "dxa", value: 101.5 }, gridSpan: 2 },
              },
            ],
          },
        ],
      }),
    );
  });

  test("a 1x1 split refuses as no change in direct and intent demand paths", () => {
    const document = documentOf({
      type: "table",
      columnWidths: [100],
      rows: [{ type: "tableRow", cells: [cell(100)] }],
    });
    const operation = {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: [],
    } satisfies DocumentOp;
    const direct = applyDocumentOp(document, operation);
    expect(direct.isErr() && direct.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE);

    const intent = {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
    } satisfies TableIntentOperation;
    const located = locateTableRow(document, intent);
    if (located.isErr()) throw located.error;
    const demand = tableEditParagraphDemand(located.value, intent);
    expect(demand.isErr() && demand.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE);
    const compiled = compileEditorIntent(document, {
      intent: { type: "table", operation: intent },
      mode: { type: "editing" },
    });
    expect(compiled.isErr() && compiled.error.reason).toBe(DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE);
  });

  test("a physically addressed vertical continuation refuses before id demand", () => {
    const document = documentOf({
      type: "table",
      columnWidths: [100],
      rows: [
        {
          type: "tableRow",
          cells: [{ ...cell(100), formatting: { vMerge: "restart" } }],
        },
        {
          type: "tableRow",
          cells: [{ ...cell(101), formatting: { vMerge: "continue" } }],
        },
      ],
    });
    const intent = {
      ...tableAddress(id(101)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
    } satisfies TableIntentOperation;
    const located = locateTableRow(document, intent);
    if (located.isErr()) throw located.error;
    const demand = tableEditParagraphDemand(located.value, intent);
    expect(demand.isErr() && demand.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    const compiled = compileEditorIntent(document, {
      intent: { type: "table", operation: intent },
      mode: { type: "editing" },
    });
    expect(compiled.isErr() && compiled.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    const direct = applyDocumentOp(document, { ...intent, newBlockIds: [] });
    expect(direct.isErr() && direct.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
    const owner = applyDocumentOp(document, {
      ...tableAddress(id(100)),
      type: DOCUMENT_OP_TYPES.SPLIT_CELL,
      newBlockIds: [],
    });
    expect(owner.isErr() && owner.error.reason).toBe(
      DOCUMENT_OP_REFUSAL_REASONS.STRUCTURE_MISMATCH,
    );
  });

  test("table intent allocates zero ids for empty vertical merge and rows minus one otherwise", () => {
    const emptyContinuation = documentOf({
      type: "table",
      columnWidths: [100],
      rows: [
        { type: "tableRow", cells: [cell(100, "owner")] },
        { type: "tableRow", cells: [cell(101, "")] },
      ],
    });
    const vertical = compileEditorIntent(emptyContinuation, {
      intent: {
        type: "table",
        operation: {
          ...tableAddress(id(100)),
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: 2,
          left: 0,
          right: 1,
        },
      },
      mode: { type: "editing" },
    });
    if (vertical.isErr()) throw vertical.error;
    const verticalOp = vertical.value.ops.at(0);
    if (verticalOp?.type !== DOCUMENT_OP_TYPES.MERGE_CELLS)
      throw new Error("The compiler emits the table merge operation.");
    expect(verticalOp.newBlockIds).toEqual([]);
    const appliedVertical = appliedDocument(emptyContinuation, verticalOp);
    const verticalTable = tableFrom(appliedVertical);
    expect(verticalTable.rows.at(0)?.cells.at(0)?.formatting?.vMerge).toBe("restart");
    expect(verticalTable.rows.at(1)?.cells.at(0)?.formatting?.vMerge).toBe("continue");
    expect(verticalTable.rows.at(0)?.cells.at(0)?.content).toEqual([paragraph(100, "owner")]);
    expect(verticalTable.rows.at(1)?.cells.at(0)?.content).toEqual([paragraph(101, "")]);
    expectGridCoverage(verticalTable);
    expect(new Set(paragraphIds(appliedVertical)).size).toBe(paragraphIds(appliedVertical).length);

    const general = documentOf({
      type: "table",
      columnWidths: [100, 100],
      rows: [
        { type: "tableRow", cells: [cell(200, "a"), cell(201, "b")] },
        { type: "tableRow", cells: [cell(202, "c"), cell(203, "d")] },
      ],
    });
    const rectangular = compileEditorIntent(general, {
      intent: {
        type: "table",
        operation: {
          ...tableAddress(id(200)),
          type: DOCUMENT_OP_TYPES.MERGE_CELLS,
          top: 0,
          bottom: 2,
          left: 0,
          right: 2,
        },
      },
      mode: { type: "editing" },
    });
    if (rectangular.isErr()) throw rectangular.error;
    const rectangularOp = rectangular.value.ops.at(0);
    if (rectangularOp?.type !== DOCUMENT_OP_TYPES.MERGE_CELLS)
      throw new Error("The compiler emits the rectangular merge operation.");
    expect(rectangularOp.newBlockIds).toHaveLength(1);
    const appliedRectangular = appliedDocument(general, rectangularOp);
    const rectangularTable = tableFrom(appliedRectangular);
    expect(rectangularTable.rows.at(0)?.cells.at(0)?.content).toEqual([
      paragraph(200, "a"),
      paragraph(201, "b"),
      paragraph(202, "c"),
      paragraph(203, "d"),
    ]);
    const continuationId = rectangularOp.newBlockIds.at(0);
    expect(rectangularTable.rows.at(1)?.cells.at(0)?.content).toEqual([
      { type: "paragraph", paraId: continuationId, content: [] },
    ]);
    expect(rectangularTable.rows.at(0)?.cells.at(0)?.formatting?.gridSpan).toBe(2);
    expect(rectangularTable.rows.at(1)?.cells.at(0)?.formatting?.gridSpan).toBe(2);
    expectGridCoverage(rectangularTable);
    const ids = paragraphIds(appliedRectangular);
    expect(new Set(ids).size).toBe(ids.length);
    expect(ids).toContain(rectangularOp.newBlockIds.at(0));
  });
});
