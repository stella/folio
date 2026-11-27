/**
 * Table edits batched on the canonical session, where one removes the
 * paragraph another is addressed by: each still finds its table, tracked as
 * directly, and whether one is refused does not depend on the mode.
 */
import { describe, expect, test } from "bun:test";

import { FolioDocxReviewer } from "../ai-edits/headless";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
} from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { BlockContent, Paragraph, Table } from "../types/document";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});

const cellId = (row: number, column: number): string => `100000${row}${column}`;

const tableOf = (
  rows: number,
  columns: number,
  textOf = (row: number, column: number) => `r${row}c${column}`,
): Table => ({
  type: "table",
  columnWidths: Array.from({ length: columns }, () => 2000),
  rows: Array.from({ length: rows }, (_row, row) => ({
    type: "tableRow",
    cells: Array.from({ length: columns }, (_cell, column) => ({
      type: "tableCell",
      content: [paragraph(cellId(row, column), textOf(row, column))],
    })),
  })),
});

const bytesOf = (table: Table): Promise<ArrayBuffer> => {
  const content: BlockContent[] = [
    paragraph("00000091", "Before"),
    table,
    paragraph("00000092", "After"),
  ];
  return createDocx({ package: { document: { content } } });
};

type BatchOptions = {
  bytes: ArrayBuffer;
  mode: "direct" | "tracked-changes";
  operations: FolioDocumentOperation[];
};

/** What a batch applied, and the words once everything is accepted. */
const accepted = async ({ bytes, mode, operations }: BatchOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(bytes.slice(0), {
    author: "Editor",
    session: "canonical",
  });
  const { applied, skipped } = reviewer.applyDocumentOperations(
    { version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION, mode, operations },
    { undefinedReferences: "refuse" },
  );
  const ids = new Set(applied.map(({ id }) => id));
  reviewer.acceptAll();
  return {
    applied: operations.flatMap(({ id }) => (ids.has(id) ? [id] : [])),
    skipped: skipped.map(({ id, message }) => ({ id, message })),
    texts: reviewer.getContent().map(({ text }) => text),
  };
};

const MODES = ["direct", "tracked-changes"] as const;

describe("column deletions of one table", () => {
  const COLUMNS = 3;
  const subsets = Array.from({ length: 1 << COLUMNS }, (_subset, mask) =>
    Array.from({ length: COLUMNS }, (_column, column) => column).filter(
      (column) => (mask & (1 << column)) !== 0,
    ),
  ).filter(({ length }) => length > 1);
  test.each(
    subsets.flatMap((columns) =>
      [columns, columns.toReversed()].flatMap((order) =>
        MODES.map((mode) => ({ order, mode })),
      ),
    ),
  )("all apply, in any order: $order ($mode)", async ({ order, mode }) => {
    const { applied, skipped, texts } = await accepted({
      bytes: await bytesOf(tableOf(2, COLUMNS)),
      mode,
      operations: order.map((column, index) => ({
        id: `op${index}`,
        type: "deleteTableColumn",
        blockId: cellId(0, column),
      })),
    });
    expect(skipped).toEqual([]);
    expect(applied).toEqual(order.map((_column, index) => `op${index}`));
    const kept = [0, 1].flatMap((row) =>
      Array.from({ length: COLUMNS }, (_column, column) => column)
        .filter((column) => !order.includes(column))
        .map((column) => `r${row}c${column}`),
    );
    expect(texts).toEqual(["Before", ...kept, "After"]);
  });
});

test.each(
  [0, 1, 2].flatMap((named) =>
    (["before", "after"] as const).flatMap((position) =>
      MODES.map((mode) => ({ named, position, mode })),
    ),
  ),
)(
  "a row inserted $position one the batch deletes, named by cell $named of it, applies ($mode)",
  async ({ named, position, mode }) => {
    for (const deletedBy of [0, 1, 2]) {
      const insertion: FolioDocumentOperation = {
        id: "insert",
        type: "insertTableRow",
        blockId: cellId(0, named),
        position,
        cellTexts: ["new"],
      };
      const deletion: FolioDocumentOperation = {
        id: "delete",
        type: "deleteTableRow",
        blockId: cellId(0, deletedBy),
      };
      for (const operations of [
        [insertion, deletion],
        [deletion, insertion],
      ]) {
        const { skipped, texts } = await accepted({
          bytes: await bytesOf(tableOf(2, 3)),
          mode,
          operations,
        });
        expect(skipped).toEqual([]);
        expect(texts).toEqual([
          "Before",
          "new",
          "",
          "",
          "r1c0",
          "r1c1",
          "r1c2",
          "After",
        ]);
      }
    }
  },
);

test("a column deletion after a merge that took its paragraph applies directly", async () => {
  const { applied, skipped } = await accepted({
    bytes: await bytesOf(tableOf(1, 2, () => "")),
    mode: "direct",
    operations: [
      {
        id: "merge",
        type: "mergeTableCells",
        blockId: cellId(0, 0),
        endBlockId: cellId(0, 1),
      },
      { id: "delete", type: "deleteTableColumn", blockId: cellId(0, 0) },
    ],
  });
  expect(skipped).toEqual([]);
  expect(applied).toEqual(["merge", "delete"]);
});

test("a split beside a merge that would close its row is refused in either mode", async () => {
  const split: FolioDocumentOperation = {
    id: "split",
    type: "splitTableCell",
    blockId: cellId(0, 0),
  };
  const merge: FolioDocumentOperation = {
    id: "merge",
    type: "mergeTableCells",
    blockId: cellId(0, 0),
    endBlockId: cellId(1, 1),
  };
  const table = tableOf(2, 2);
  const [first, second] = table.rows;
  const [start] = first?.cells ?? [];
  const [continued] = second?.cells ?? [];
  if (start === undefined || continued === undefined)
    throw new Error("The table has two rows.");
  start.formatting = { vMerge: "restart" };
  continued.formatting = { vMerge: "continue" };
  continued.content = [paragraph(cellId(1, 0), "")];
  const bytes = await bytesOf(table);
  for (const operations of [
    [split, merge],
    [merge, split],
  ]) {
    const direct = await accepted({ bytes, mode: "direct", operations });
    const tracked = await accepted({
      bytes,
      mode: "tracked-changes",
      operations,
    });
    expect(direct.applied).not.toContain("split");
    expect(tracked.applied).not.toContain("split");
  }
});

test.each(MODES)(
  "of two merges whose cells overlap, the later one is refused (%s)",
  async (mode) => {
    const pair: FolioDocumentOperation = {
      id: "pair",
      type: "mergeTableCells",
      blockId: cellId(0, 2),
      endBlockId: cellId(0, 3),
    };
    const row: FolioDocumentOperation = {
      id: "row",
      type: "mergeTableCells",
      blockId: cellId(0, 0),
      endBlockId: cellId(0, 3),
    };
    const bytes = await bytesOf(tableOf(1, 4, () => ""));
    for (const [first, second] of [
      [pair, row],
      [row, pair],
    ] as const) {
      const { applied, skipped } = await accepted({
        bytes,
        mode,
        operations: [first, second],
      });
      expect(applied).toEqual([first.id]);
      expect(skipped).toEqual([
        {
          id: second.id,
          message:
            "An earlier merge or split of the batch reshapes cells this merges or splits.",
        },
      ]);
    }
  },
);

test.each(MODES)(
  "an edit of a wide cell the batch deletes every column of is refused (%s)",
  async (mode) => {
    const table = tableOf(2, 3);
    const [first] = table.rows;
    const [wide, , last] = first?.cells ?? [];
    if (first === undefined || wide === undefined || last === undefined)
      throw new Error("The table has three columns.");
    wide.formatting = { gridSpan: 2 };
    first.cells = [wide, last];
    const bytes = await bytesOf(table);
    const edit: FolioDocumentOperation = {
      id: "edit",
      type: "replaceBlock",
      blockId: cellId(0, 0),
      text: "kept",
    };
    const deletions: FolioDocumentOperation[] = [0, 1].map((column) => ({
      id: `delete${column}`,
      type: "deleteTableColumn",
      blockId: cellId(1, column),
    }));
    // Either deletion alone narrows the cell; together they remove it, so
    // whichever of the edit and the second deletion comes later is refused.
    for (const [operations, refused, claimedBy] of [
      [[...deletions, edit], "edit", "delete1"],
      [[edit, ...deletions], "delete1", "edit"],
    ] as const) {
      const { applied, skipped, texts } = await accepted({
        bytes,
        mode,
        operations: [...operations],
      });
      expect(skipped).toContainEqual({
        id: refused,
        message: `operation "${claimedBy}", earlier in this batch, already claims its target.`,
      });
      expect(texts.includes("kept")).toBe(applied.includes("edit"));
    }
  },
);

test.each(MODES)(
  "a merge below a deleted row that starts a vertical merge is refused (%s)",
  async (mode) => {
    const table = tableOf(3, 2, (row, column) =>
      column === 1 && row < 2 ? `r${row}c${column}` : "",
    );
    for (const [index, row] of table.rows.entries()) {
      const [first] = row.cells;
      if (first === undefined) throw new Error("Every row has a cell.");
      first.formatting = { vMerge: index === 0 ? "restart" : "continue" };
    }
    const { applied, skipped } = await accepted({
      bytes: await bytesOf(table),
      mode,
      operations: [
        {
          id: "merge",
          type: "mergeTableCells",
          blockId: cellId(1, 1),
          endBlockId: cellId(2, 1),
        },
        { id: "delete", type: "deleteTableRow", blockId: cellId(0, 0) },
      ],
    });
    expect(applied).toEqual(["delete"]);
    expect(skipped.map(({ id }) => id)).toEqual(["merge"]);
  },
);
