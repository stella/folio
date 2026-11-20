import { expect, test } from "bun:test";

import type {
  Document,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

let next = 0x10;
const cell = (
  text: string,
  change?: { type: "tableCellInsertion" | "tableCellDeletion"; id: number },
): TableCell => ({
  type: "tableCell",
  ...(change === undefined
    ? {}
    : {
        structuralChange: {
          type: change.type,
          info: { id: change.id, author: "Reviewer", date: "2026-11-20" },
        },
      }),
  content: [
    {
      type: "paragraph",
      paraId: (next++).toString(16).toUpperCase().padStart(8, "0"),
      content: [{ type: "run", content: [{ type: "text", text }] }],
    },
  ],
});

type Shape = { columnWidths: number[]; gridAfter?: number };

/** Rows `A` / the middle row / `D`, the middle row wider than the others. */
const tableDocument = (middle: TableCell[], shape: Shape): Document => {
  const row = (cells: TableCell[], after?: number): TableRow => ({
    type: "tableRow",
    ...(after === undefined ? {} : { formatting: { gridAfter: after } }),
    cells,
  });
  const table: Table = {
    type: "table",
    columnWidths: shape.columnWidths,
    rows: [
      row([cell("A")], shape.gridAfter),
      row(middle),
      row([cell("D")], shape.gridAfter),
    ],
  };
  return {
    package: {
      document: {
        content: [
          {
            type: "paragraph",
            paraId: "00000001",
            content: [
              { type: "run", content: [{ type: "text", text: "End" }] },
            ],
          },
          table,
          { type: "paragraph", paraId: "00000002", content: [] },
        ],
      },
    },
  };
};

/** Each row as its cells' texts with their spans, and its trailing grid columns. */
const rowsOf = (document: Document) => {
  const table = document.package.document.content.at(1);
  if (table?.type !== "table") throw new Error("The table stays in place.");
  return {
    columnWidths: table.columnWidths,
    rows: table.rows.map((row) => {
      const cells = row.cells.map(
        (entry) =>
          `${JSON.stringify(entry.content).match(/"text":"(\w)"/u)?.[1] ?? ""}${
            entry.formatting?.gridSpan === undefined
              ? ""
              : `*${String(entry.formatting.gridSpan)}`
          }`,
      );
      const gridAfter = row.formatting?.gridAfter;
      return gridAfter === undefined ? { cells } : { cells, gridAfter };
    }),
  };
};

const resolveAll = (document: Document, decision: RevisionDecision) => {
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [11, 12],
    decision,
  });
  if (result.isErr()) throw result.error;
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toEqual(document);
  return rowsOf(result.value.document);
};

const ONE_COLUMN = { columnWidths: [4000] };
const ins = { type: "tableCellInsertion", id: 11 } as const;
const del = { type: "tableCellDeletion", id: 12 } as const;

// A row that ends early holds the columns after it: removing the cell that
// made another row wider hands that column to the cell beside it.
const CASES = [
  {
    name: "a deleted extra cell",
    middle: () => [cell("B"), cell("C", del)],
    shape: ONE_COLUMN,
    accept: [["A"], ["B*2"], ["D"]],
    reject: [["A"], ["B", "C"], ["D"]],
  },
  {
    name: "an inserted extra cell",
    middle: () => [cell("B"), cell("C", ins)],
    shape: ONE_COLUMN,
    accept: [["A"], ["B", "C"], ["D"]],
    reject: [["A"], ["B*2"], ["D"]],
  },
  {
    name: "an inserted cell before a deleted one",
    middle: () => [cell("B", ins), cell("C", del)],
    shape: ONE_COLUMN,
    accept: [["A"], ["B*2"], ["D"]],
    reject: [["A"], ["C*2"], ["D"]],
  },
] as const;

for (const entry of CASES)
  for (const decision of Object.values(REVISION_DECISIONS))
    test(`a ragged table resolves ${entry.name} (${decision})`, () => {
      const resolved = resolveAll(
        tableDocument([...entry.middle()], entry.shape),
        decision,
      );
      expect(resolved.columnWidths).toEqual(entry.shape.columnWidths);
      const expected: readonly (readonly string[])[] =
        entry[decision === REVISION_DECISIONS.ACCEPT ? "accept" : "reject"];
      expect(resolved.rows.map(({ cells }) => cells)).toEqual(
        Array.from(expected, (cells) => Array.from(cells)),
      );
    });

test("a column only gridAfter holds stays when the cell in it goes", () => {
  const resolved = resolveAll(
    tableDocument([cell("B"), cell("C", del)], {
      columnWidths: [4000, 4000],
      gridAfter: 1,
    }),
    REVISION_DECISIONS.ACCEPT,
  );
  expect(resolved).toEqual({
    columnWidths: [4000, 4000],
    rows: [
      { cells: ["A"], gridAfter: 1 },
      { cells: ["B*2"] },
      { cells: ["D"], gridAfter: 1 },
    ],
  });
});

test("column edits refuse a table whose rows end early", () => {
  const result = applyDocumentOp(
    tableDocument([cell("B"), cell("C")], ONE_COLUMN),
    {
      type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "00000010",
      column: 1,
      width: 400,
      newBlockIds: ["00000030", "00000031", "00000032"],
    },
  );
  expect(result.isErr()).toBe(true);
});
