import { GlobalRegistrator } from "@happy-dom/global-registrator";
import { panic } from "better-result";
import { afterAll, beforeAll, expect, test } from "bun:test";
import fc from "fast-check";
import type { Node as PMNode } from "prosemirror-model";
import { EditorState } from "prosemirror-state";
import { EditorView } from "prosemirror-view";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import { serializeTableFormatting } from "../docx/serializer/tableSerializer";
import { expectTableAttrs, expectTableCellAttrs } from "./attrs";
import { tableAttrsToFormatting } from "./conversion/fromProseDoc";
import { schema } from "./schema";
import { commitColumnResize, commitRightEdgeResize } from "./tableResize";

beforeAll(() => GlobalRegistrator.register());
afterAll(() => GlobalRegistrator.unregister());

const partitions = (cuts: readonly boolean[]) => {
  const spans: number[] = [];
  let span = 1;
  for (const cut of cuts) {
    if (cut) {
      spans.push(span);
      span = 1;
    } else span++;
  }
  spans.push(span);
  return spans;
};

const resizeCase = fc
  .array(fc.integer({ min: 300, max: 4000 }), { minLength: 2, maxLength: 8 })
  .chain((widths) =>
    fc.record({
      widths: fc.constant(widths),
      cuts: fc.array(fc.boolean(), {
        minLength: widths.length - 1,
        maxLength: widths.length - 1,
      }),
      column: fc.integer({ min: 0, max: widths.length - 2 }),
      newLeft: fc.integer({ min: 300, max: 4000 }),
      newRight: fc.integer({ min: 300, max: 4000 }),
    }),
  );

const tableDocument = (widths: number[], spans: number[]) => {
  const rows = [
    spans,
    [widths.length],
    widths.map(() => 1),
    widths.map(() => 1),
  ].map((rowSpans, rowIndex) => {
    let column = 0;
    return schema.node(
      "tableRow",
      { height: 900 },
      rowSpans.map((colspan, cellIndex) => {
        const width = widths
          .slice(column, column + colspan)
          .reduce((total, value) => total + value, 0);
        column += colspan;
        return schema.node(
          "tableCell",
          {
            colspan,
            width,
            widthType: "dxa",
            colwidth: Array.from({ length: colspan }, () => 30),
            ...(rowIndex === 2 && cellIndex === 0
              ? { _omittedGridSlot: "before" }
              : {}),
            ...(rowIndex === 3 && cellIndex === rowSpans.length - 1
              ? { _omittedGridSlot: "after" }
              : {}),
          },
          schema.node(
            "paragraph",
            null,
            (rowIndex === 2 && cellIndex === 0) ||
              (rowIndex === 3 && cellIndex === rowSpans.length - 1)
              ? undefined
              : schema.text("cell"),
          ),
        );
      }),
    );
  });
  return schema.node(
    "doc",
    null,
    schema.node(
      "table",
      { columnWidths: widths, _originalFormatting: { layout: "autofit" } },
      rows,
    ),
  );
};

type AssertResizeOptions = {
  before: PMNode;
  after: PMNode;
  widths: number[];
  firstColumn: number;
  lastColumn: number;
};

const assertResize = ({
  before,
  after,
  widths,
  firstColumn,
  lastColumn,
}: AssertResizeOptions) => {
  const beforeTable = before.firstChild;
  const table = after.firstChild;
  if (!beforeTable || !table) panic("A resize fixture must contain its table.");
  expect(expectTableAttrs(table).columnWidths).toEqual(widths);
  expect(
    serializeTableFormatting(tableAttrsToFormatting(expectTableAttrs(table))),
  ).toContain('<w:tblLayout w:type="fixed"/>');
  table.forEach((row, _rowOffset, rowIndex) => {
    expect(row.attrs).toEqual(beforeTable.child(rowIndex).attrs);
    let column = 0;
    row.forEach((cell, _cellOffset, cellIndex) => {
      const original = beforeTable.child(rowIndex).child(cellIndex);
      const attrs = expectTableCellAttrs(cell);
      const originalAttrs = expectTableCellAttrs(original);
      const end = column + attrs.colspan;
      const overlaps = column <= lastColumn && end > firstColumn;
      if (overlaps) {
        expect(attrs.width).toBe(
          widths.slice(column, end).reduce((total, value) => total + value, 0),
        );
        expect(attrs.widthType).toBe("dxa");
        expect(attrs.colwidth).toBeUndefined();
        expect(cell.attrs["colwidth"]).toBeNull();
      } else expect(cell).toBe(original);
      expect(attrs.colspan).toBe(originalAttrs.colspan);
      expect(attrs._omittedGridSlot).toBe(originalAttrs._omittedGridSlot);
      expect(cell.content).toBe(original.content);
      column = end;
    });
    expect(column).toBe(widths.length);
  });
};

for (const type of ["pair", "rightEdge"] as const) {
  test(
    `${type} resize preserves every cell's covered grid width`,
    () => {
      assertProperty(
        fc.property(
          resizeCase,
          ({ widths, cuts, column, newLeft, newRight }) => {
            const before = tableDocument(widths, partitions(cuts));
            let dispatches = 0;
            const view = new EditorView(document.createElement("div"), {
              state: EditorState.create({ doc: before }),
              dispatchTransaction(transaction) {
                dispatches++;
                view.updateState(view.state.apply(transaction));
              },
            });
            try {
              const expected = [...widths];
              const firstColumn = type === "pair" ? column : widths.length - 1;
              const lastColumn = type === "pair" ? column + 1 : firstColumn;
              expected[firstColumn] = newLeft;
              switch (type) {
                case "pair":
                  expected[lastColumn] = newRight;
                  commitColumnResize(view, {
                    pmStart: 0,
                    colIdx: column,
                    newLeft,
                    newRight,
                  });
                  break;
                case "rightEdge":
                  commitRightEdgeResize(view, {
                    pmStart: 0,
                    colIdx: firstColumn,
                    newWidth: newLeft,
                  });
                  break;
                default: {
                  const exhaustive: never = type;
                  panic(`Unknown resize type ${exhaustive}`);
                }
              }
              expect(dispatches).toBe(1);
              assertResize({
                before,
                after: view.state.doc,
                widths: expected,
                firstColumn,
                lastColumn,
              });
            } finally {
              view.destroy();
            }
          },
        ),
        {
          id: `${type} resize preserves every cell's covered grid width`,
          numRuns: 80,
          examples: [
            [
              {
                widths: [1000, 1000, 1000],
                cuts: [false, true],
                column: 1,
                newLeft: 1200,
                newRight: 800,
              },
            ],
            [
              {
                widths: [1000, 1000, 1000],
                cuts: [false, true],
                column: 0,
                newLeft: 1200,
                newRight: 800,
              },
            ],
          ],
        },
      );
    },
    propertyTestTimeout(20_000),
  );
}
