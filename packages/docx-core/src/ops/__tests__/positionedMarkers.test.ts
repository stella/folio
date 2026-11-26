/**
 * A bookmark marker beside a table's rows, or beside a row's cells, keeps its
 * place among them when rows or cells come and go, and every edit's inverse
 * puts it back exactly.
 */
import { expect, test } from "bun:test";

import type {
  Document,
  PositionedBookmarkMarker,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
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

/** A start and an end of bookmark `id`, both standing at `index`. */
const marks = (id: number, index: number): PositionedBookmarkMarker[] => [
  { index, marker: { type: "bookmarkStart", id, name: `b${String(id)}` } },
  { index, marker: { type: "bookmarkEnd", id } },
];

const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [
        table,
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

const indices = (markers: readonly PositionedBookmarkMarker[] | undefined) =>
  markers?.map(({ index, marker }) => [marker.type, index]);

/** Rows 1, 2, 3, with bookmark 1 before row 2 and bookmark 2 after row 3. */
const betweenRows = (): Document =>
  documentOf({
    type: "table",
    bookmarks: [...marks(1, 1), ...marks(2, 3)],
    rows: ["10000001", "10000002", "10000003"].map((id): TableRow => ({
      type: "tableRow",
      cells: [cell(id)],
    })),
  });

const row = (blockId: string) => ({ story: OP_STORIES.MAIN, blockId });

test("a removed row leaves the markers beside it among the rows that remain", () => {
  for (const [blockId, expected] of [
    ["10000001", [0, 0, 2, 2]],
    ["10000002", [1, 1, 2, 2]],
    ["10000003", [1, 1, 2, 2]],
  ] as const) {
    const next = applied(betweenRows(), {
      type: DOCUMENT_OP_TYPES.DELETE_ROW,
      ...row(blockId),
    });
    expect(tableIn(next).bookmarks?.map(({ index }) => index)).toEqual([
      ...expected,
    ]);
  }
});

test("an added row goes after an end and before a start standing where it goes", () => {
  const next = applied(betweenRows(), {
    type: DOCUMENT_OP_TYPES.INSERT_ROW,
    ...row("10000001"),
    at: 1,
    row: { type: "tableRow", cells: [cell("10000004")] },
  });
  expect(indices(tableIn(next).bookmarks)).toEqual([
    ["bookmarkStart", 2],
    ["bookmarkEnd", 1],
    ["bookmarkStart", 4],
    ["bookmarkEnd", 4],
  ]);
});

test("an accepted row deletion removes the row, not the markers beside it", () => {
  const tracked = applied(betweenRows(), {
    type: DOCUMENT_OP_TYPES.DELETE_ROW,
    ...row("10000003"),
    revision: { id: 50, author: "Reviewer", date: "2026-11-21T00:00:00Z" },
    newIds: { revision: [51, 52, 53, 54] },
  });
  const accepted = applied(tracked, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [50, 51, 52, 53, 54],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  expect(tableIn(accepted).rows).toHaveLength(2);
  expect(tableIn(accepted).bookmarks?.map(({ index }) => index)).toEqual([
    1, 1, 2, 2,
  ]);
});

/** One row of cells 1 and 2 with bookmark 1 between them and 2 after them. */
const betweenCells = (): Document =>
  documentOf({
    type: "table",
    columnWidths: [1000, 1000],
    rows: [
      {
        type: "tableRow",
        bookmarks: [...marks(1, 1), ...marks(2, 2)],
        cells: [cell("10000001"), cell("10000002")],
      },
      { type: "tableRow", cells: [cell("10000003"), cell("10000004")] },
    ],
  });

const rowMarkers = (document: Document) =>
  tableIn(document)
    .rows.at(0)
    ?.bookmarks?.map(({ index }) => index);

test("a removed column leaves the markers beside its cells in the row", () => {
  const next = applied(betweenCells(), {
    type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
    ...row("10000001"),
    column: 1,
  });
  expect(rowMarkers(next)).toEqual([1, 1, 1, 1]);
});

test("an accepted column deletion removes the cells, not the markers beside them", () => {
  const tracked = applied(betweenCells(), {
    type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
    ...row("10000001"),
    column: 0,
    revision: { id: 60, author: "Reviewer", date: "2026-11-21T00:00:00Z" },
    newIds: { revision: [61, 62, 63, 64] },
  });
  const accepted = applied(tracked, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [60, 61, 62, 63, 64],
    decision: REVISION_DECISIONS.ACCEPT,
  });
  expect(tableIn(accepted).rows.at(0)?.cells).toHaveLength(1);
  expect(rowMarkers(accepted)).toEqual([0, 0, 1, 1]);
});

/** Two rows of two cells, bookmark 1 around the whole table. */
const aroundTable = (): Document =>
  documentOf({
    type: "table",
    columnWidths: [1000, 1000],
    bookmarks: [
      { index: 0, marker: { type: "bookmarkStart", id: 1, name: "b1" } },
      { index: 2, marker: { type: "bookmarkEnd", id: 1 } },
    ],
    rows: [
      { type: "tableRow", cells: [cell("10000001"), cell("10000002")] },
      { type: "tableRow", cells: [cell("10000003"), cell("10000004")] },
    ],
  });

test("a column or cell edit, direct or suggested, keeps the bookmarks beside the rows", () => {
  const stamp = (id: number) => ({
    revision: { id, author: "Reviewer", date: "2026-11-21T00:00:00Z" },
    newIds: {
      revision: Array.from({ length: 12 }, (_, index) => id + 1 + index),
    },
  });
  const edits: [string, DocumentOp, boolean][] = [
    [
      "column insertion",
      {
        type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
        ...row("10000001"),
        column: 1,
        width: 1000,
        newBlockIds: ["20000001", "20000002"],
      },
      true,
    ],
    [
      "column deletion",
      { type: DOCUMENT_OP_TYPES.DELETE_COLUMN, ...row("10000001"), column: 1 },
      true,
    ],
    [
      "cell merge",
      {
        type: DOCUMENT_OP_TYPES.MERGE_CELLS,
        ...row("10000001"),
        top: 0,
        bottom: 1,
        left: 0,
        right: 2,
        newBlockIds: [],
      },
      false,
    ],
  ];
  const kept = (document: Document) =>
    expect(indices(tableIn(document).bookmarks)).toEqual([
      ["bookmarkStart", 0],
      ["bookmarkEnd", 2],
    ]);
  for (const [name, op, suggests] of edits) {
    kept(applied(aroundTable(), op));
    if (!suggests) continue;
    const tracked = applied(aroundTable(), { ...op, ...stamp(70) });
    kept(tracked);
    const revisionIds = Array.from({ length: 13 }, (_, index) => 70 + index);
    for (const decision of [
      REVISION_DECISIONS.ACCEPT,
      REVISION_DECISIONS.REJECT,
    ]) {
      const resolved = applyDocumentOp(tracked, {
        type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
        story: OP_STORIES.MAIN,
        revisionIds,
        decision,
      });
      if (resolved.isErr())
        throw new Error(`${name}: ${resolved.error.message}`);
      kept(resolved.value.document);
    }
  }
});
