import { expect, test } from "bun:test";
import type { Node } from "prosemirror-model";
import { EditorState } from "prosemirror-state";

import { resolvedMarkupDocumentView } from "@stll/docx-core/ops";

import { resolveAllChangesInHeadlessState } from "../__tests__/reference-editor/comments";
import { projectReferenceMarkupView } from "../__tests__/reference-editor/markupViewProjection";
import type {
  Document,
  Paragraph,
  TableCell,
  TableRow,
} from "../types/document";
import { toProseDoc } from "./conversion/toProseDoc";
import {
  type MarkupViewProjection,
  projectMarkupView,
} from "./markupViewProjection";

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
  const native = resolvedMarkupDocumentView(source, decision);
  expect(native.unresolved).toEqual([]);
  expect(blocks(toProseDoc(native.document))).toEqual(blocks(reference.doc));
  expect(blocks(reference.doc)).toEqual([...body, "paragraph:Tail."]);
});

test("simple markup bars follow surviving table membership changes and joined text", () => {
  const info = (id: number) => ({ id, author: "Reviewer" });
  const source: Document = {
    package: {
      document: {
        content: [
          {
            ...paragraph("00000001", "Joined "),
            pPrMark: { kind: "del", info: info(1) },
          },
          paragraph("00000002", "paragraph"),
          {
            type: "table",
            propertyChanges: [{ type: "tablePropertyChange", info: info(2) }],
            rows: [
              {
                ...row(
                  cell(paragraph("00000003", "Added row first cell")),
                  cell(paragraph("00000004", "Added row second cell")),
                ),
                structuralChange: { type: "tableRowInsertion", info: info(3) },
              },
              {
                ...row(
                  {
                    ...cell(paragraph("00000005", "Added cell")),
                    structuralChange: {
                      type: "tableCellInsertion",
                      info: info(4),
                    },
                  },
                  {
                    ...cell(paragraph("00000006", "Cell property only")),
                    propertyChanges: [
                      { type: "tableCellPropertyChange", info: info(5) },
                    ],
                  },
                ),
                propertyChanges: [
                  { type: "tableRowPropertyChange", info: info(6) },
                ],
              },
              {
                ...row(
                  cell({
                    type: "paragraph",
                    paraId: "00000007",
                    content: [
                      {
                        type: "deletion",
                        info: info(7),
                        content: [
                          {
                            type: "run",
                            content: [{ type: "text", text: "Removed" }],
                          },
                        ],
                      },
                    ],
                  }),
                  cell(paragraph("00000008", "Removed too")),
                ),
                structuralChange: { type: "tableRowDeletion", info: info(8) },
              },
              row(
                cell(paragraph("00000009", "Standing first cell")),
                cell(paragraph("0000000A", "Standing second cell")),
              ),
            ],
          },
          paragraph("0000000B", "Tail"),
        ],
      },
    },
  };
  const state = EditorState.create({ doc: toProseDoc(source) });
  const expected = marked(projectReferenceMarkupView(state, "simple-markup"));
  expect(expected).toEqual([
    "Joined paragraph",
    "Added row first cell",
    "Added row second cell",
    "Added cell",
  ]);
  expect(marked(projectMarkupView(state, "simple-markup"))).toEqual(expected);
});

const marked = (view: MarkupViewProjection): string[] => {
  if (view.type === "authored") return [];
  const found: string[] = [];
  view.doc.descendants((node, position) => {
    if (node.type.name !== "paragraph") return true;
    if (
      view.changeBarRanges.some(
        ({ from, to }) => from < position + node.nodeSize && to > position,
      )
    )
      found.push(node.textContent);
    return false;
  });
  return found;
};

test.each(["owner", "continuation"] as const)(
  "simple markup marks resized and reviewed vertical spans when removing the %s row",
  (removedRow) => {
    const owner = (paraId: string, text: string): TableCell => ({
      ...cell(paragraph(paraId, text)),
      formatting: { vMerge: "restart" },
    });
    const continuation = (paraId: string): TableCell => ({
      ...cell({ type: "paragraph", paraId, content: [] }),
      formatting: { vMerge: "continue" },
    });
    const source: Document = {
      package: {
        document: {
          content: [
            {
              type: "table",
              rows: [
                {
                  ...row(
                    owner("00000001", "Resized span"),
                    cell(paragraph("00000002", "Other cell in the first row")),
                  ),
                  ...(removedRow === "owner"
                    ? { structuralChange: deletedRow.structuralChange }
                    : {}),
                },
                {
                  ...row(
                    continuation("00000003"),
                    cell(paragraph("00000004", "Other cell in the next row")),
                  ),
                  ...(removedRow === "continuation"
                    ? { structuralChange: deletedRow.structuralChange }
                    : {}),
                },
              ],
            },
            {
              type: "table",
              rows: [
                row(
                  owner("00000005", "Reviewed continuation owner"),
                  cell(paragraph("00000006", "Standing second table")),
                ),
                row(
                  {
                    ...continuation("00000007"),
                    structuralChange: {
                      type: "tableCellMerge",
                      info: paragraphInfo,
                      verticalMerge: "continue",
                      verticalMergeOriginal: "rest",
                    },
                  },
                  cell(paragraph("00000008", "Standing next row")),
                ),
              ],
            },
            paragraph("00000009", "Tail"),
          ],
        },
      },
    };
    const state = EditorState.create({ doc: toProseDoc(source) });
    const reference = projectReferenceMarkupView(state, "simple-markup");
    const candidate = projectMarkupView(state, "simple-markup");
    expect(marked(reference)).toEqual([
      "Resized span",
      "Reviewed continuation owner",
    ]);
    expect(marked(candidate)).toEqual(marked(reference));
  },
);
