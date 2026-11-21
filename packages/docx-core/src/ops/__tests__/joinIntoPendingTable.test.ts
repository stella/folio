/**
 * A tracked join through a paragraph whose mark already goes runs into the
 * table after it. A table under review cannot receive what runs in, a refusal
 * of its own; any other table that cannot is a shape no tracked join records
 * yet.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  Table,
  TableRow,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});

const tableOf = (row: TableRow): Table => ({
  type: "table",
  rows: [row],
});

const cellRow = (cellParagraph: Paragraph): TableRow => ({
  type: "tableRow",
  cells: [{ type: "tableCell", content: [cellParagraph] }],
});

/** The first and second paragraph, the second's mark pending deletion, then `table`. */
const documentBefore = (table: Table): Document => ({
  package: {
    document: {
      content: [
        paragraph("00000001", "first"),
        {
          ...paragraph("00000002", "second"),
          pPrMark: { kind: "del", info: { id: 5, ...OTHER } },
        },
        table,
        paragraph("00000004", "after"),
      ],
    },
  },
});

const joinReason = (table: Table) => {
  const joined = applyDocumentOp(documentBefore(table), {
    type: DOCUMENT_OP_TYPES.JOIN_BLOCKS,
    story: OP_STORIES.MAIN,
    blockId: "00000001",
    nextBlockId: "00000002",
    newIds: { revision: [11, 12, 13] },
    revision: { id: 10, author: "Reviewer", date: "2026-02-03T04:05:06Z" },
  });
  return joined.isErr() ? joined.error.reason : undefined;
};

describe("a tracked join that would run into the table after it", () => {
  test("is refused by its own reason when the table's first row is under review", () => {
    expect(
      joinReason(
        tableOf({
          ...cellRow(paragraph("00000003", "cell")),
          structuralChange: {
            type: "tableRowInsertion",
            info: { id: 6, ...OTHER },
          },
        }),
      ),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.JOIN_INTO_PENDING_TABLE);
  });

  test("is untrackable when the table's first paragraph ends a section", () => {
    expect(
      joinReason(
        tableOf(
          cellRow({
            ...paragraph("00000003", "cell"),
            sectionProperties: { pageWidth: 11906 },
          }),
        ),
      ),
    ).toBe(DOCUMENT_OP_REFUSAL_REASONS.UNTRACKABLE);
  });
});
