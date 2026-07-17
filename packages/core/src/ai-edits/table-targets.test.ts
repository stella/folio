import { describe, expect, test } from "bun:test";
import { type Node as PMNode, Schema } from "prosemirror-model";
import { TableMap } from "prosemirror-tables";

import {
  findEnclosingTableBoundary,
  findEnclosingTableCell,
  findEnclosingTableRow,
  tableRectangleCutsMergedCell,
} from "./table-targets";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: {
      content: "text*",
      group: "block",
      attrs: { paraId: { default: null } },
    },
    text: {},
    table: {
      content: "tableRow+",
      group: "block",
      tableRole: "table",
    },
    tableRow: {
      content: "tableCell+",
      tableRole: "row",
    },
    tableCell: {
      content: "block+",
      tableRole: "cell",
      attrs: {
        colspan: { default: 1 },
        rowspan: { default: 1 },
        colwidth: { default: null },
      },
    },
  },
});

const paragraph = (paraId: string): PMNode =>
  schema.node("paragraph", { paraId }, [schema.text(paraId)]);

const cell = (content: PMNode[], attrs?: { colspan?: number; rowspan?: number }): PMNode =>
  schema.node("tableCell", attrs, content);

const findParagraphPosition = (doc: PMNode, paraId: string): number => {
  let position: number | null = null;
  doc.descendants((node, nodePosition) => {
    if (node.attrs["paraId"] === paraId) {
      position = nodePosition;
      return false;
    }
  });
  if (position === null) {
    throw new Error(`Missing paragraph ${paraId}.`);
  }
  return position;
};

describe("table targets", () => {
  test("resolves the nearest table, row, and cell for a nested anchor", () => {
    const nestedTable = schema.node("table", null, [
      schema.node("tableRow", null, [cell([paragraph("nested")])]),
    ]);
    const outerTable = schema.node("table", null, [
      schema.node("tableRow", null, [cell([paragraph("outer"), nestedTable])]),
    ]);
    const doc = schema.node("doc", null, [outerTable]);
    const blockFrom = findParagraphPosition(doc, "nested");

    const boundary = findEnclosingTableBoundary(doc, blockFrom);
    const row = findEnclosingTableRow(doc, blockFrom);
    const targetCell = findEnclosingTableCell(doc, blockFrom);

    expect(boundary).not.toBeNull();
    expect(doc.nodeAt(boundary?.before ?? -1)).toEqual(nestedTable);
    expect(row?.table).toEqual(nestedTable);
    expect(row?.rowIndex).toBe(0);
    expect(targetCell).toMatchObject({
      tablePosition: boundary?.before,
      leftColumnIndex: 0,
      rightColumnIndex: 1,
      topRowIndex: 0,
      bottomRowIndex: 1,
    });
  });

  test("returns the full grid rectangle for a spanning target cell", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        cell([paragraph("span")], { colspan: 2 }),
        cell([paragraph("side")]),
      ]),
    ]);
    const doc = schema.node("doc", null, [table]);

    expect(findEnclosingTableCell(doc, findParagraphPosition(doc, "span"))).toMatchObject({
      leftColumnIndex: 0,
      rightColumnIndex: 2,
      topRowIndex: 0,
      bottomRowIndex: 1,
    });
  });

  test("detects rectangle boundaries that cut through a merged cell", () => {
    const table = schema.node("table", null, [
      schema.node("tableRow", null, [
        cell([paragraph("span")], { colspan: 2 }),
        cell([paragraph("top-right")]),
      ]),
      schema.node("tableRow", null, [
        cell([paragraph("bottom-left")]),
        cell([paragraph("bottom-center")]),
        cell([paragraph("bottom-right")]),
      ]),
    ]);
    const map = TableMap.get(table);

    expect(tableRectangleCutsMergedCell(map, { left: 1, top: 0, right: 3, bottom: 2 })).toBe(true);
    expect(tableRectangleCutsMergedCell(map, { left: 0, top: 0, right: 3, bottom: 2 })).toBe(false);
  });
});
