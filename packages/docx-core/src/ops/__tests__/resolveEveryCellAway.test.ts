/**
 * A decision that takes every cell of a table away (inserted cells rejected,
 * deleted cells accepted) takes the table with them, as removing its last row
 * does: a one-column table whose column was deleted cell by cell, or a table
 * whose every cell was inserted. Its inverse gives the table back exactly.
 */

import { expect, test } from "bun:test";

import type {
  Document,
  Table,
  TableStructuralChangeInfo,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");

type TableOfOptions = {
  rows: number;
  columns: number;
  change: TableStructuralChangeInfo["type"];
};
const tableOf = ({ rows, columns, change }: TableOfOptions): Table => ({
  type: "table",
  columnWidths: Array.from({ length: columns }, () => 900),
  rows: Array.from({ length: rows }, (_, row) => ({
    type: "tableRow",
    cells: Array.from({ length: columns }, (_cell, column) => ({
      type: "tableCell",
      structuralChange: {
        type: change,
        info: { id: 10 + row * columns + column, author: "Reviewer" },
      },
      content: [
        {
          type: "paragraph",
          paraId: id(0x100 + row * columns + column),
          content: [
            {
              type: "run",
              content: [{ type: "text", text: `cell ${row}:${column}` }],
            },
          ],
        },
      ],
    })),
  })),
});

const documentWith = (table: Table): Document => ({
  package: {
    document: {
      content: [
        { type: "paragraph", paraId: "00000001", content: [] },
        table,
        { type: "paragraph", paraId: "00000002", content: [] },
      ],
    },
  },
});

const cases = [
  { change: "tableCellDeletion", decision: REVISION_DECISIONS.ACCEPT },
  { change: "tableCellInsertion", decision: REVISION_DECISIONS.REJECT },
] as const;

const SHAPES = [
  { rows: 1, columns: 1 },
  { rows: 3, columns: 1 },
  { rows: 2, columns: 2 },
] as const;

test.each(
  cases.flatMap(({ change, decision }) =>
    SHAPES.map(({ rows, columns }) => ({ change, decision, rows, columns })),
  ),
)(
  "$decision of every $change ($rows x $columns) removes the table",
  ({ change, decision, rows, columns }) => {
    const table = tableOf({ rows, columns, change });
    const before = documentWith(table);
    const resolved = applyDocumentOp(before, {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: Array.from({ length: rows * columns }, (_, n) => 10 + n),
      decision,
    }).unwrap();
    expect(resolved.document.package.document.content).toEqual([
      { type: "paragraph", paraId: "00000001", content: [] },
      { type: "paragraph", paraId: "00000002", content: [] },
    ]);
    expect(contractViolation(resolved.document)).toBeUndefined();
    expect(
      applyDocumentOps(resolved.document, resolved.inverse).unwrap().document,
    ).toStrictEqual(before);
  },
);
