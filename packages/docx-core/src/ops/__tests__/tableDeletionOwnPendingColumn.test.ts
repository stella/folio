/**
 * A tracked deletion of a table holding its author's own pending column: the
 * column goes outright, as deleting over one's own insertion removes it, and
 * the rest of the table is suggested for deletion. A column another author
 * suggested keeps the refusal.
 */
import { expect, test } from "bun:test";

import type {
  BlockContent,
  Document,
  Paragraph,
  Table,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const original = (): Document => {
  const table: Table = {
    type: "table",
    columnWidths: [900, 900],
    rows: [0, 1].map((row) => ({
      type: "tableRow" as const,
      cells: [0, 1].map((cell) => ({
        type: "tableCell" as const,
        content: [paragraph(`0000001${row * 2 + cell}`, `Cell ${row}${cell}.`)],
      })),
    })),
  };
  const content: BlockContent[] = [
    paragraph("00000001", "Before."),
    table,
    paragraph("00000002", "After."),
  ];
  return { package: { document: { content } } };
};

const stamp = (id: number, author: string) => ({
  id,
  author,
  date: "2026-02-03T04:05:06Z",
});

const withPendingColumn = (author: string) => {
  const inserted = applyDocumentOp(original(), {
    type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
    story: OP_STORIES.MAIN,
    blockId: "00000010",
    column: 2,
    width: 700,
    newBlockIds: ["00000020", "00000021"],
    revision: stamp(50, author),
    newIds: { revision: Array.from({ length: 16 }, (_, index) => 51 + index) },
  });
  if (inserted.isErr()) throw inserted.error;
  return inserted.value.document;
};

const deletion = {
  type: DOCUMENT_OP_TYPES.DELETE_TABLE,
  story: OP_STORIES.MAIN,
  blockId: "00000010",
  revision: stamp(100, "Reviewer"),
  newIds: { revision: Array.from({ length: 16 }, (_, index) => 101 + index) },
} as const;

const tableIn = (document: Document) => {
  const table = document.package.document.content.find(
    (block): block is Table => block.type === "table",
  );
  if (table === undefined) throw new Error("The document has a table.");
  return table;
};

test("a tracked table deletion removes the author's own pending column first", () => {
  const document = withPendingColumn("Reviewer");
  const deleted = applyDocumentOp(document, deletion);
  if (deleted.isErr()) throw deleted.error;
  const table = tableIn(deleted.value.document);
  expect(table.propertyChanges ?? []).toEqual([]);
  expect(
    table.rows.map(({ cells, structuralChange }) => [
      cells.length,
      structuralChange?.type,
    ]),
  ).toEqual([
    [2, "tableRowDeletion"],
    [2, "tableRowDeletion"],
  ]);

  // Rejecting the deletion leaves the table as it was before the column.
  const rejected = applyDocumentOps(deleted.value.document, [
    {
      type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
      story: OP_STORIES.MAIN,
      revisionIds: deleted.value.revisions,
      decision: REVISION_DECISIONS.REJECT,
    },
  ]);
  if (rejected.isErr()) throw rejected.error;
  expect(tableIn(rejected.value.document).rows).toEqual(
    tableIn(original()).rows,
  );

  const undone = applyDocumentOps(
    deleted.value.document,
    deleted.value.inverse,
  );
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
});

test("another author's pending column keeps the table deletion refused", () => {
  const result = applyDocumentOp(withPendingColumn("Other"), deletion);
  expect(result.isErr()).toBe(true);
});
