import { expect, test } from "bun:test";
import type { Node } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { resolvedDocumentView } from "@stll/docx-core/ops";

import { resolveAllChangesInHeadlessState } from "../__tests__/reference-editor/comments";
import type {
  Document,
  Paragraph,
  TableCell,
  TableRow,
} from "../types/document";
import { toProseDoc } from "./conversion/toProseDoc";

const paragraphInfo = { id: 1, author: "Reviewer" };
const receiverInfo = { id: 2, author: "Reviewer" };
const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (content: Paragraph): TableCell => ({
  type: "tableCell",
  content: [content],
});
const row = (...cells: TableCell[]): TableRow => ({
  type: "tableRow",
  cells,
});
const blocks = (doc: Node) => {
  const result: string[] = [];
  doc.forEach((block) =>
    result.push(`${block.type.name}:${block.textContent}`),
  );
  return result;
};

const pending = cell(paragraph("00000002", "Pending. "));
const standing = cell(paragraph("00000003", "Standing."));
const addedRow = {
  ...row(pending),
  structuralChange: { type: "tableRowInsertion", info: receiverInfo },
} as const satisfies TableRow;
const deletedRow = {
  ...row(pending),
  structuralChange: { type: "tableRowDeletion", info: receiverInfo },
} as const satisfies TableRow;
const deletedCell = {
  ...pending,
  structuralChange: { type: "tableCellDeletion", info: receiverInfo },
} as const satisfies TableCell;

test.each([
  {
    description: "rejects an inserted row in a surviving table",
    decision: "reject",
    rows: [addedRow, row(standing)],
    body: ["paragraph:Outside. ", "table:Standing."],
  },
  {
    description: "rejects a whole inserted table before a standing table",
    decision: "reject",
    rows: [addedRow],
    nextRows: [row(standing)],
    body: ["table:Outside. Standing."],
  },
  {
    description: "accepts an inserted receiver row",
    decision: "accept",
    rows: [addedRow, row(standing)],
    body: ["paragraph:Outside. ", "table:Pending. Standing."],
  },
  {
    description: "accepts a deleted receiver row",
    decision: "accept",
    rows: [deletedRow, row(standing)],
    body: ["paragraph:Outside. ", "table:Standing."],
  },
  {
    description: "rejects a deleted receiver row",
    decision: "reject",
    rows: [deletedRow],
    body: ["table:Outside. Pending. "],
  },
  {
    description: "rejects an inserted receiver cell",
    decision: "reject",
    rows: [
      row(
        {
          ...pending,
          structuralChange: { type: "tableCellInsertion", info: receiverInfo },
        },
        standing,
      ),
    ],
    body: ["paragraph:Outside. ", "table:Standing."],
  },
  {
    description: "accepts a deleted receiver cell",
    decision: "accept",
    rows: [row(deletedCell, standing)],
    body: ["paragraph:Outside. ", "table:Standing."],
  },
  {
    description: "rejects a deleted receiver cell",
    decision: "reject",
    rows: [row(deletedCell, standing)],
    body: ["table:Outside. Pending. Standing."],
  },
] satisfies {
  description: string;
  decision: "accept" | "reject";
  rows: TableRow[];
  nextRows?: TableRow[];
  body: string[];
}[])("markup resolution $description", ({ decision, rows, nextRows, body }) => {
  const source: Document = {
    package: {
      document: {
        content: [
          {
            ...paragraph("00000001", "Outside. "),
            pPrMark: {
              kind: decision === "reject" ? "ins" : "del",
              info: paragraphInfo,
            },
          },
          { type: "table", rows },
          ...(nextRows === undefined
            ? []
            : [{ type: "table" as const, rows: nextRows }]),
          paragraph("00000004", "Tail."),
        ],
      },
    },
  };
  const state = EditorState.create({ doc: toProseDoc(source) });
  const reference = resolveAllChangesInHeadlessState(state, decision);
  const native = resolvedDocumentView(source, decision);
  expect(native.unresolved).toEqual([]);
  expect(blocks(toProseDoc(native.document))).toEqual(blocks(reference.doc));
  expect(blocks(reference.doc)).toEqual([...body, "paragraph:Tail."]);
});
