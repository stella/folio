import { describe, expect, test } from "bun:test";
import { Schema } from "prosemirror-model";

import {
  planTableMutations,
  type TableMutationPlanTarget,
  type TableRectangle,
} from "./table-mutation-plan";

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "text*", group: "block" },
    text: {},
    table: { content: "row+", group: "block", tableRole: "table" },
    row: { content: "cell+", tableRole: "row" },
    cell: {
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

const grid = (rows: number, columns: number) =>
  schema.node(
    "table",
    null,
    Array.from({ length: rows }, () =>
      schema.node(
        "row",
        null,
        Array.from({ length: columns }, () =>
          schema.node("cell", null, [schema.node("paragraph")]),
        ),
      ),
    ),
  );

const firstTable = grid(4, 3);
const doc = schema.node("doc", null, [firstTable, grid(4, 3)]);
const otherTablePosition = firstTable.nodeSize;

type Candidate = {
  item: string;
  operationId: string;
  target: TableMutationPlanTarget;
};

const rectangle = (left: number, top: number, right: number, bottom: number): TableRectangle => ({
  left,
  top,
  right,
  bottom,
});

const candidate = (item: string, target: TableMutationPlanTarget): Candidate => ({
  item,
  operationId: item,
  target,
});

describe("table mutation planning", () => {
  test("rejects cell-shape edits that share a table with structural edits", () => {
    const plan = planTableMutations(
      [
        candidate("insert-row", { type: "tableStructure", tablePosition: 0 }),
        candidate("merge", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 0, 1, 2),
        }),
        candidate("split", {
          type: "splitCell",
          tablePosition: otherTablePosition,
          rectangle: rectangle(0, 0, 1, 2),
        }),
        candidate("delete-column", { type: "tableStructure", tablePosition: otherTablePosition }),
      ],
      doc,
    );

    expect(plan.executable).toEqual(["insert-row", "delete-column"]);
    expect(plan.skipped).toEqual([
      { id: "merge", reason: "unsupportedBlock" },
      { id: "split", reason: "unsupportedBlock" },
    ]);
  });

  test("rejects merge and split combinations on the same table", () => {
    const plan = planTableMutations(
      [
        candidate("merge", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 0, 1, 2),
        }),
        candidate("split", {
          type: "splitCell",
          tablePosition: 0,
          rectangle: rectangle(1, 0, 2, 2),
        }),
      ],
      doc,
    );

    expect(plan.executable).toEqual([]);
    expect(plan.skipped).toEqual([
      { id: "merge", reason: "unsupportedBlock" },
      { id: "split", reason: "unsupportedBlock" },
    ]);
  });

  test("distinguishes duplicate, overlapping, and disjoint rectangles", () => {
    const plan = planTableMutations(
      [
        candidate("first", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 0, 1, 2),
        }),
        candidate("duplicate", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 0, 1, 2),
        }),
        candidate("overlap", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 1, 1, 3),
        }),
        candidate("disjoint", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(1, 0, 2, 2),
        }),
        candidate("other-table", {
          type: "mergeCells",
          tablePosition: otherTablePosition,
          rectangle: rectangle(0, 1, 1, 3),
        }),
        candidate("ordinary-edit", { type: "none" }),
      ],
      doc,
    );

    expect(plan.executable).toEqual(["first", "disjoint", "other-table", "ordinary-edit"]);
    expect(plan.skipped).toEqual([
      { id: "duplicate", reason: "noopOperation" },
      { id: "overlap", reason: "unsupportedBlock" },
    ]);
  });

  test("disjoint vertical merges keep a physical cell in every shared row", () => {
    for (const rows of [2, 3, 4]) {
      for (const columns of [2, 3, 4]) {
        for (const order of [
          Array.from({ length: columns }, (_, column) => column),
          Array.from({ length: columns }, (_, column) => columns - column - 1),
        ]) {
          const table = grid(rows, columns);
          const nested = schema.node("doc", null, [
            schema.node("table", null, [
              schema.node("row", null, [
                schema.node("cell", null, [
                  schema.node("paragraph"),
                  table,
                  schema.node("paragraph"),
                ]),
              ]),
            ]),
          ]);
          let tablePosition = -1;
          nested.descendants((node, position) => {
            if (node === table) tablePosition = position;
          });
          expect(tablePosition).toBeGreaterThan(0);
          const merges = order.map((column) =>
            candidate(`column-${column}`, {
              type: "mergeCells",
              tablePosition,
              rectangle: rectangle(column, 0, column + 1, rows),
            }),
          );
          const plan = planTableMutations(merges, nested);
          expect(plan.executable).toEqual(merges.slice(0, -1).map(({ item }) => item));
          expect(plan.skipped).toEqual([
            { id: merges.at(-1)!.operationId, reason: "unsupportedBlock" },
          ]);
        }
      }
    }
  });

  test("keeps a row-removing merge alone so other row coordinates cannot drift", () => {
    const table = grid(4, 2);
    const plan = planTableMutations(
      [
        candidate("remove-row", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 0, 2, 2),
        }),
        candidate("below", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(0, 2, 2, 3),
        }),
      ],
      schema.node("doc", null, [table]),
    );
    expect(plan.executable).toEqual(["remove-row"]);
    expect(plan.skipped).toEqual([{ id: "below", reason: "unsupportedBlock" }]);
  });

  test("counts spanning cells by their physical origins when deciding row survival", () => {
    const cell = (attrs: { colspan?: number; rowspan?: number }) =>
      schema.node("cell", attrs, [schema.node("paragraph")]);
    const table = schema.node("table", null, [
      schema.node("row", null, [cell({ rowspan: 2 }), cell({ colspan: 2 }), cell({})]),
      schema.node("row", null, [cell({ colspan: 2 }), cell({})]),
    ]);
    const plan = planTableMutations(
      [
        candidate("wide", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(1, 0, 3, 2),
        }),
        candidate("narrow", {
          type: "mergeCells",
          tablePosition: 0,
          rectangle: rectangle(3, 0, 4, 2),
        }),
      ],
      schema.node("doc", null, [table]),
    );
    expect(plan.executable).toEqual(["wide"]);
    expect(plan.skipped).toEqual([{ id: "narrow", reason: "unsupportedBlock" }]);
  });
});
