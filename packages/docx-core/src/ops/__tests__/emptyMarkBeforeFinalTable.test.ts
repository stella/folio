import { expect, test } from "bun:test";

import type { BlockContent, Document, Paragraph } from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

// A package other producers write: the story ends with a table inserted after
// a paragraph whose deletion, mark included, is pending.
const info = (id: number) => ({ id, author: "Reviewer", date: "2026-11-18" });
const text = (paraId: string, value: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: value }] }],
});
const deleted: Paragraph = {
  type: "paragraph",
  paraId: "00000002",
  content: [
    {
      type: "deletion",
      info: info(2),
      content: [
        { type: "run", content: [{ type: "deletedText", text: "Clause two" }] },
      ],
    },
  ],
  pPrMark: { kind: "del", info: info(3) },
};
const insertedCell = (paraId: string, value: string, id: number) => ({
  type: "tableCell" as const,
  content: [
    {
      type: "paragraph" as const,
      paraId,
      content: [
        {
          type: "insertion" as const,
          info: info(id),
          content: [
            {
              type: "run" as const,
              content: [{ type: "text" as const, text: value }],
            },
          ],
        },
      ],
    },
  ],
});
const insertedTable: BlockContent = {
  type: "table",
  rows: [
    {
      type: "tableRow",
      structuralChange: { type: "tableRowInsertion", info: info(1) },
      cells: [insertedCell("00000011", "Term", 0)],
    },
  ],
};
const tableOf = (paraId: string): BlockContent => ({
  type: "table",
  rows: [
    {
      type: "tableRow",
      cells: [{ type: "tableCell", content: [text(paraId, "Term")] }],
    },
  ],
});
const plainTable = tableOf("00000011");

const resolve = (
  content: BlockContent[],
  revisionIds: number[],
  decision: (typeof REVISION_DECISIONS)[keyof typeof REVISION_DECISIONS],
) =>
  applyDocumentOp({ package: { document: { content } } } satisfies Document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });

for (const table of ["pending", "plain"] as const)
  test(`accepting an emptied deletion before a story-final ${table} table removes the paragraph`, () => {
    const result = resolve(
      [
        text("00000001", "Body"),
        deleted,
        table === "pending" ? insertedTable : plainTable,
      ],
      table === "pending" ? [0, 1, 2, 3] : [2, 3],
      REVISION_DECISIONS.ACCEPT,
    );
    if (result.isErr()) throw result.error;
    expect(result.value.document.package.document.content).toEqual([
      text("00000001", "Body"),
      plainTable,
    ]);
  });

test("accepting a deleted mark of words before a story-final table runs them into its first cell", () => {
  const marked: Paragraph = {
    ...text("00000001", "Body "),
    pPrMark: { kind: "del", info: info(4) },
  };
  const result = resolve(
    [text("00000005", "Lead"), marked, deleted, plainTable],
    [2, 3, 4],
    REVISION_DECISIONS.ACCEPT,
  );
  if (result.isErr()) throw result.error;
  const content = result.value.document.package.document.content;
  expect(content).toHaveLength(2);
  const [, table] = content;
  if (table?.type !== "table") throw new Error("The table stands alone.");
  const cell = table.rows[0]?.cells[0]?.content[0];
  expect(cell?.type === "paragraph" && cell.paraId).toBe("00000011");
  expect(JSON.stringify(cell)).toContain("Body ");
  expect(JSON.stringify(cell)).toContain("Term");
});

test("accepting an emptied deletion between two tables at the story's end removes it", () => {
  const result = resolve(
    [text("00000001", "Body"), tableOf("00000012"), deleted, plainTable],
    [2, 3],
    REVISION_DECISIONS.ACCEPT,
  );
  if (result.isErr()) throw result.error;
  expect(result.value.document.package.document.content).toEqual([
    text("00000001", "Body"),
    tableOf("00000012"),
    plainTable,
  ]);
});
