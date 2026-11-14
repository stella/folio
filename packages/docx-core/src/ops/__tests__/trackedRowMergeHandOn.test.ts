/**
 * A suggested deletion of the row a vertical merge starts keeps the merged
 * cell's content: it stays unmarked while pending, accepting hands the merge
 * on to the next row exactly as deleting the row directly does, and
 * rejecting gives the table back.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { removalOrphansAMerge, rowsWithoutRow } from "../tableRowRemoval";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  type RevisionDecision,
} from "../types";

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (
  paraId: string,
  text: string,
  vMerge?: "restart" | "continue",
): TableCell => ({
  type: "tableCell",
  ...(vMerge === undefined ? {} : { formatting: { vMerge } }),
  content: [paragraph(paraId, text)],
});
const row = (...cells: TableCell[]): TableRow => ({ type: "tableRow", cells });

/** A merge in the first column over `merged` rows, then one plain row. */
const mergeTable = (merged: number): Table => ({
  type: "table",
  columnWidths: [2000, 2000],
  rows: [
    ...Array.from({ length: merged }, (_, index) =>
      row(
        cell(
          `000000${index + 1}0`,
          index === 0 ? "A1" : "",
          index === 0 ? "restart" : "continue",
        ),
        cell(`000000${index + 1}1`, `B${index + 1}`),
      ),
    ),
    row(cell("00000090", "A9"), cell("00000091", "B9")),
  ],
});
const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "Before"),
        table,
        paragraph("00000002", "After"),
      ],
    },
  },
});
const firstTable = (document: Document): Table => {
  const table = document.package.document.content.find(
    (block) => block.type === "table",
  );
  if (table?.type !== "table") throw new Error("The table remains.");
  return table;
};

const revision = { id: 100, author: "Reviewer", date: "2026-05-06T07:08:09Z" };
const newIds = {
  revision: Array.from({ length: 32 }, (_value, index) => 101 + index),
};
const applied = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  const undone = applyDocumentOps(result.value.document, result.value.inverse);
  if (undone.isErr()) throw undone.error;
  expect(undone.value.document).toStrictEqual(document);
  return result.value;
};
const resolve = (
  document: Document,
  revisionIds: readonly number[],
  decision: RevisionDecision,
) =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
const deleteFirstRow = (tracked: boolean) =>
  ({
    type: DOCUMENT_OP_TYPES.DELETE_ROW,
    story: OP_STORIES.MAIN,
    blockId: "00000011",
    ...(tracked ? { revision, newIds } : {}),
  }) as const satisfies DocumentOp;

describe("a suggested deletion of the row a vertical merge starts", () => {
  for (const merged of [2, 3])
    test(`keeps the merged cell's words over ${merged} rows, pending, accepted and rejected`, () => {
      const document = documentOf(mergeTable(merged));
      const result = applied(document, deleteFirstRow(true));
      const pending = firstTable(result.document).rows.at(0);
      expect(pending?.structuralChange?.type).toBe("tableRowDeletion");
      // The merged cell outlives the row: nothing in it is marked.
      expect(pending?.cells.at(0)).toStrictEqual(
        firstTable(document).rows.at(0)?.cells.at(0),
      );
      // The row's own cell is marked as the row is.
      const own = pending?.cells.at(1)?.content.at(0);
      if (own?.type !== "paragraph")
        throw new Error("The cell holds a paragraph.");
      expect(own.pPrMark?.kind).toBe("del");
      expect(own.content.at(0)?.type).toBe("deletion");

      const direct = applied(document, deleteFirstRow(false));
      const accepted = resolve(
        result.document,
        result.revisions,
        REVISION_DECISIONS.ACCEPT,
      ).document;
      expect(accepted).toStrictEqual(direct.document);
      const first = firstTable(accepted).rows.at(0)?.cells.at(0);
      expect(first?.content.at(0)).toStrictEqual(paragraph("00000010", "A1"));
      expect(first?.formatting?.vMerge).toBe(
        merged > 2 ? "restart" : undefined,
      );
      expect(
        resolve(result.document, result.revisions, REVISION_DECISIONS.REJECT)
          .document,
      ).toStrictEqual(document);
    });

  test("a suggested table deletion still marks the merged cell", () => {
    const document = documentOf(mergeTable(2));
    const result = applied(document, {
      type: DOCUMENT_OP_TYPES.DELETE_TABLE,
      story: OP_STORIES.MAIN,
      blockId: "00000011",
      revision,
      newIds,
    });
    const owner = firstTable(result.document)
      .rows.at(0)
      ?.cells.at(0)
      ?.content.at(0);
    if (owner?.type !== "paragraph")
      throw new Error("The cell holds a paragraph.");
    expect(owner.content.at(0)?.type).toBe("deletion");
  });
});

describe("taking a row out of a table", () => {
  test("hands the merges it starts on, where dropping it alone orphans them", () => {
    const table = mergeTable(3);
    const dropped = table.rows.slice(1);
    expect(removalOrphansAMerge(table, dropped)).toBe(true);
    const handed = rowsWithoutRow(table, 0);
    expect(removalOrphansAMerge(table, handed)).toBe(false);
    expect(handed.at(0)?.cells.at(0)?.content).toStrictEqual([
      paragraph("00000010", "A1"),
    ]);
    // A row that only continues a merge goes without handing anything on.
    expect(rowsWithoutRow(table, 1)).toStrictEqual([
      table.rows[0],
      table.rows[2],
      table.rows[3],
    ]);
  });
});
