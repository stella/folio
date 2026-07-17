import { describe, expect, test } from "bun:test";

import {
  planTableMutations,
  type TableMutationPlanTarget,
  type TableRectangle,
} from "./table-mutation-plan";

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
    const plan = planTableMutations([
      candidate("insert-row", { type: "tableStructure", tablePosition: 10 }),
      candidate("merge", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      candidate("split", {
        type: "splitCell",
        tablePosition: 20,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      candidate("delete-column", { type: "tableStructure", tablePosition: 20 }),
    ]);

    expect(plan.executable).toEqual(["insert-row", "delete-column"]);
    expect(plan.skipped).toEqual([
      { id: "merge", reason: "unsupportedBlock" },
      { id: "split", reason: "unsupportedBlock" },
    ]);
  });

  test("rejects merge and split combinations on the same table", () => {
    const plan = planTableMutations([
      candidate("merge", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      candidate("split", {
        type: "splitCell",
        tablePosition: 10,
        rectangle: rectangle(1, 0, 2, 2),
      }),
    ]);

    expect(plan.executable).toEqual([]);
    expect(plan.skipped).toEqual([
      { id: "merge", reason: "unsupportedBlock" },
      { id: "split", reason: "unsupportedBlock" },
    ]);
  });

  test("distinguishes duplicate, overlapping, and disjoint rectangles", () => {
    const plan = planTableMutations([
      candidate("first", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      candidate("duplicate", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      candidate("overlap", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(0, 1, 1, 3),
      }),
      candidate("disjoint", {
        type: "mergeCells",
        tablePosition: 10,
        rectangle: rectangle(1, 0, 2, 2),
      }),
      candidate("other-table", {
        type: "mergeCells",
        tablePosition: 20,
        rectangle: rectangle(0, 1, 1, 3),
      }),
      candidate("ordinary-edit", { type: "none" }),
    ]);

    expect(plan.executable).toEqual(["first", "disjoint", "other-table", "ordinary-edit"]);
    expect(plan.skipped).toEqual([
      { id: "duplicate", reason: "noopOperation" },
      { id: "overlap", reason: "unsupportedBlock" },
    ]);
  });
});
