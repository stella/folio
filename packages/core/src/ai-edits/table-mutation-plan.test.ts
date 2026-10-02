import { describe, expect, test } from "bun:test";
import fc from "fast-check";
import { Schema } from "prosemirror-model";

import { assertProperty } from "../../../../test/property-testing";

import {
  planTableMutations,
  tableMergeFoldedRows,
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

const schema = new Schema({
  nodes: {
    doc: { content: "block+" },
    paragraph: { content: "text*", group: "block" },
    text: {},
    table: { content: "tableRow+", group: "block", tableRole: "table" },
    tableRow: { content: "tableCell+", tableRole: "row" },
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

const planningTable = schema.node(
  "table",
  null,
  Array.from({ length: 3 }, () =>
    schema.node(
      "tableRow",
      null,
      Array.from({ length: 3 }, () => schema.node("tableCell", null, [schema.node("paragraph")])),
    ),
  ),
);

type MergeCandidateOptions = { item: string; tablePosition: number; rectangle: TableRectangle };

const mergeCandidate = ({ item, tablePosition, rectangle: bounds }: MergeCandidateOptions) =>
  candidate(item, {
    type: "mergeCells",
    tablePosition,
    rectangle: bounds,
    foldedRows: tableMergeFoldedRows(planningTable, bounds),
  });

describe("table mutation planning", () => {
  test("disjoint merges retain a starting cell in every shared row, in any order", () => {
    assertProperty(
      fc.property(
        fc.array(fc.integer({ min: 1, max: 3 }), { minLength: 2, maxLength: 6 }),
        fc.integer({ min: 2, max: 5 }),
        fc.nat(),
        fc.boolean(),
        (widths, height, rotation, reverse) => {
          const table = schema.node(
            "table",
            null,
            Array.from({ length: height }, () =>
              schema.node(
                "tableRow",
                null,
                widths.map((colspan) =>
                  schema.node("tableCell", { colspan }, [schema.node("paragraph")]),
                ),
              ),
            ),
          );
          let left = 0;
          const candidates = widths.map((width, index) => {
            const bounds = rectangle(left, 0, left + width, height);
            left += width;
            return candidate(String(index), {
              type: "mergeCells",
              tablePosition: 10,
              rectangle: bounds,
              foldedRows: tableMergeFoldedRows(table, bounds),
            });
          });
          const start = rotation % candidates.length;
          const rotated = [...candidates.slice(start), ...candidates.slice(0, start)];
          const ordered = reverse ? rotated.toReversed() : rotated;
          const plan = planTableMutations(ordered);
          expect(plan.executable).toEqual(ordered.slice(0, -1).map(({ item }) => item));
          expect(plan.skipped).toEqual([
            { id: ordered.at(-1)?.operationId, reason: "unsupportedBlock" },
          ]);
        },
      ),
      { numRuns: 100 },
    );
  });

  test("rejects cell-shape edits that share a table with structural edits", () => {
    const plan = planTableMutations([
      candidate("insert-row", { type: "tableStructure", tablePosition: 10 }),
      mergeCandidate({
        item: "merge",
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
      mergeCandidate({
        item: "merge",
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
      mergeCandidate({
        item: "first",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      mergeCandidate({
        item: "duplicate",
        tablePosition: 10,
        rectangle: rectangle(0, 0, 1, 2),
      }),
      mergeCandidate({
        item: "overlap",
        tablePosition: 10,
        rectangle: rectangle(0, 1, 1, 3),
      }),
      mergeCandidate({
        item: "disjoint",
        tablePosition: 10,
        rectangle: rectangle(1, 0, 2, 2),
      }),
      mergeCandidate({
        item: "other-table",
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
