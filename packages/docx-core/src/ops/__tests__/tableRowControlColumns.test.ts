/**
 * A content control wrapping a table's rows keeps wrapping them through a
 * column or cell edit, direct or suggested, accepted or rejected: the edit
 * keeps every row, and every inverse puts the table back exactly.
 */
import { expect, test } from "bun:test";

import type { Document, Table, TableCell } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
} from "../types";

const cell = (paraId: string): TableCell => ({
  type: "tableCell",
  content: [
    {
      type: "paragraph",
      paraId,
      content: [{ type: "run", content: [{ type: "text", text: paraId }] }],
    },
  ],
});

const CONTROL = { sdtType: "richText" as const, id: 9, tag: "rows" };

/** Two rows of two cells, one content control around both rows. */
const wrapped = (): Document => ({
  package: {
    document: {
      content: [
        {
          type: "table",
          columnWidths: [1000, 1000],
          rows: [
            {
              type: "tableRow",
              cells: [cell("10000001"), cell("10000002")],
              contentControls: [CONTROL],
            },
            {
              type: "tableRow",
              cells: [cell("10000003"), cell("10000004")],
              contentControls: [CONTROL],
            },
          ],
        },
        {
          type: "paragraph",
          paraId: "00000009",
          content: [{ type: "run", content: [{ type: "text", text: "Tail" }] }],
        },
      ],
    },
  },
});

const tableIn = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("The table stays first.");
  return table;
};

const applied = (document: Document, op: DocumentOp): Document => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value.document;
};

const wrapsEveryRow = (document: Document) =>
  expect(tableIn(document).rows.map((row) => row.contentControls)).toEqual([
    [CONTROL],
    [CONTROL],
  ]);

const at = { story: OP_STORIES.MAIN, blockId: "10000001" };

test("a column or cell edit, direct or suggested, keeps a row content control around its rows", () => {
  const edits: [string, DocumentOp, boolean][] = [
    [
      "column insertion",
      {
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        ...at,
        column: 1,
        width: 1000,
        newBlockIds: ["20000001", "20000002"],
      },
      true,
    ],
    [
      "column deletion",
      { type: DOCUMENT_OP_TYPES.DELETE_COLUMN, ...at, column: 1 },
      true,
    ],
    [
      "cell merge",
      {
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        ...at,
        top: 0,
        bottom: 1,
        left: 0,
        right: 2,
        newBlockIds: [],
      },
      false,
    ],
  ];
  for (const [name, op, suggests] of edits) {
    const direct = applied(wrapped(), op);
    wrapsEveryRow(direct);
    if (!suggests) continue;
    const tracked = applied(wrapped(), {
      ...op,
      revision: { id: 70, author: "Reviewer", date: "2026-11-21T00:00:00Z" },
      newIds: {
        revision: Array.from({ length: 12 }, (_, index) => 71 + index),
      },
    });
    wrapsEveryRow(tracked);
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const resolved = applyDocumentOp(tracked, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds: Array.from({ length: 13 }, (_, index) => 70 + index),
        decision,
      });
      if (resolved.isErr())
        throw new Error(`${name}: ${resolved.error.message}`);
      wrapsEveryRow(resolved.value.document);
      expect(tableIn(resolved.value.document).rows[0]?.cells).toHaveLength(
        decision === REVISION_DECISIONS.REJECT
          ? 2
          : (tableIn(direct).rows[0]?.cells.length ?? 0),
      );
    }
  }
});
