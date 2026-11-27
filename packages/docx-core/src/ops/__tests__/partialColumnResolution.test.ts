/**
 * A tracked column insertion of which only some cells are accepted: rejecting
 * the rest keeps the grid columns the accepted cells hold, and each row that
 * loses its inserted cell gives those columns to the cell before it, or to
 * the cell after it when the inserted cell started the row.
 */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type {
  Document,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { tableGrid } from "../tableGrid";
import { DOCUMENT_OP_TYPES, OP_STORIES, REVISION_DECISIONS } from "../types";

const cell = (paraId: string, inserted?: number): TableCell => ({
  type: "tableCell",
  content: [{ type: "paragraph", paraId, content: [] }],
  ...(inserted === undefined
    ? {}
    : {
        structuralChange: {
          type: "tableCellInsertion",
          info: { id: inserted, author: "Reviewer" },
        },
      }),
});

/** Three rows with a column inserted at `at`; `gridChange` records the earlier grid. */
const tableWithInsertedColumn = (at: 0 | 1, gridChange: boolean): Table => ({
  type: "table",
  ...(gridChange
    ? { formatting: { gridChange: { id: 9, columnWidths: [500, 700] } } }
    : {}),
  columnWidths: at === 0 ? [300, 500, 700] : [500, 300, 700],
  rows: [1, 2, 3].map((row) => {
    const kept = [cell(`0000${row}0A0`), cell(`0000${row}0B0`)];
    const added = cell(`0000${row}0C0`, row);
    return {
      type: "tableRow",
      cells: at === 0 ? [added, ...kept] : [kept[0]!, added, kept[1]!],
    };
  }),
});

const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [table, { type: "paragraph", paraId: "00000001", content: [] }],
    },
  },
});

const resolve = (
  document: Document,
  revisionIds: number[],
  decision: "accept" | "reject",
): Document => {
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  });
  if (result.isErr()) throw result.error;
  return result.value.document;
};

const tableOf = (document: Document): Table => {
  const [table] = document.package.document.content;
  if (table?.type !== "table") throw new Error("The table remains first.");
  return table;
};

const shape = (table: Table) =>
  table.rows.map((row) =>
    row.cells.map((each) => {
      const [paragraph] = each.content;
      return `${paragraph?.type === "paragraph" ? paragraph.paraId : "?"}:${each.formatting?.gridSpan ?? 1}${each.structuralChange === undefined ? "" : "*"}`;
    }),
  );

describe("rejecting the rest of a partly accepted column insertion", () => {
  for (const gridChange of [false, true])
    test(`keeps the accepted cell's column (grid change recorded: ${gridChange})`, () => {
      const document = documentOf(tableWithInsertedColumn(1, gridChange));
      const accepted = resolve(document, [2], REVISION_DECISIONS.ACCEPT);
      const rejected = tableOf(
        resolve(
          accepted,
          [1, 3, ...(gridChange ? [9] : [])],
          REVISION_DECISIONS.REJECT,
        ),
      );
      expect(shape(rejected)).toEqual([
        ["000010A0:2", "000010B0:1"],
        ["000020A0:1", "000020C0:1", "000020B0:1"],
        ["000030A0:2", "000030B0:1"],
      ]);
      expect(rejected.columnWidths).toEqual([500, 300, 700]);
      expect(rejected.formatting?.gridChange).toBeUndefined();
      expect(tableGrid(rejected, DOCUMENT_OP_TYPES.SET_TABLE).isOk()).toBe(
        true,
      );
    });

  test("gives a first-column cell's column to the cell after it", () => {
    const document = documentOf(tableWithInsertedColumn(0, false));
    const accepted = resolve(document, [2], REVISION_DECISIONS.ACCEPT);
    expect(
      shape(tableOf(resolve(accepted, [1, 3], REVISION_DECISIONS.REJECT))),
    ).toEqual([
      ["000010A0:2", "000010B0:1"],
      ["000020C0:1", "000020A0:1", "000020B0:1"],
      ["000030A0:2", "000030B0:1"],
    ]);
  });

  test("removes the column when every row rejects its cell", () => {
    const document = documentOf(tableWithInsertedColumn(1, true));
    const rejected = tableOf(
      resolve(document, [1, 2, 3, 9], REVISION_DECISIONS.REJECT),
    );
    expect(shape(rejected)).toEqual([
      ["000010A0:1", "000010B0:1"],
      ["000020A0:1", "000020B0:1"],
      ["000030A0:1", "000030B0:1"],
    ]);
    expect(rejected.columnWidths).toEqual([500, 700]);
  });
});

test(
  "partly rejecting cell insertion records preserves vertical merges across omission positions",
  () => {
    assertProperty(
      fc.property(
        fc.constantFrom("before", "after", "interior"),
        fc.integer({ min: 1, max: 4 }),
        fc.constantFrom("unset", "dxa", "pct"),
        (edge, span, measurement) => {
          const grid = edge === "before" ? "gridBefore" : "gridAfter";
          const width = edge === "before" ? "widthBefore" : "widthAfter";
          const omitted = measurement === "unset" ? 0 : 1;
          const hasNextCell = edge === "interior";
          const rowFormatting =
            measurement === "unset"
              ? undefined
              : { [grid]: omitted, [width]: { type: measurement, value: 180 } };
          const table: Table = {
            type: "table",
            columnWidths: Array.from(
              { length: span + 1 + omitted + Number(hasNextCell) },
              (_, slot) => 300 + slot * 100,
            ),
            rows: [0, 1, 2].map((row) => {
              const added: TableCell = {
                ...cell(`0000${row}0A0`, row + 1),
                formatting: { gridSpan: span },
              };
              const merged: TableCell = {
                ...cell(`0000${row}0B0`),
                ...(row === 2
                  ? {}
                  : {
                      formatting: {
                        vMerge: row === 0 ? "restart" : "continue",
                      },
                    }),
              };
              const cells =
                edge === "before" ? [added, merged] : [merged, added];
              if (hasNextCell) cells.push(cell(`0000${row}0C0`));
              const result: TableRow = {
                type: "tableRow",
                cells,
              };
              if (rowFormatting !== undefined)
                result.formatting = rowFormatting;
              return result;
            }),
          };
          const source = resolve(
            documentOf(table),
            [3],
            REVISION_DECISIONS.ACCEPT,
          );
          const settled = applyDocumentOp(source, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [1, 2],
            decision: REVISION_DECISIONS.REJECT,
          }).unwrap();
          const resolved = tableOf(settled.document);
          expect(resolved.columnWidths).toEqual(table.columnWidths);
          const settledFormatting = hasNextCell
            ? rowFormatting
            : { ...rowFormatting, [grid]: omitted + span };
          expect(resolved.rows.map((row) => row.formatting)).toEqual([
            settledFormatting,
            settledFormatting,
            rowFormatting,
          ]);
          const widenedFollowing = hasNextCell ? [{ gridSpan: span + 1 }] : [];
          const retainedRow =
            edge === "before"
              ? [{ gridSpan: span }, undefined]
              : [undefined, { gridSpan: span }];
          if (hasNextCell) retainedRow.push(undefined);
          expect(
            resolved.rows.map(({ cells }) =>
              cells.map(({ formatting }) => formatting),
            ),
          ).toEqual([
            [{ vMerge: "restart" }, ...widenedFollowing],
            [{ vMerge: "continue" }, ...widenedFollowing],
            retainedRow,
          ]);
          expect(tableGrid(resolved, DOCUMENT_OP_TYPES.SET_TABLE).isOk()).toBe(
            true,
          );
          expect(
            applyDocumentOps(settled.document, settled.inverse).unwrap()
              .document,
          ).toStrictEqual(source);
        },
      ),
      {
        id: "partial cell insertion review preserves vertical merges across omissions",
      },
    );
  },
  propertyTestTimeout(10_000),
);

test(
  "selective cell review keeps a blank interior slot between vertical merges",
  () => {
    assertProperty(
      fc.property(
        fc.integer({ min: 1, max: 4 }),
        fc.constantFrom(REVISION_DECISIONS.ACCEPT, REVISION_DECISIONS.REJECT),
        (span, decision) => {
          const middle = {
            ...cell("000010B0"),
            formatting: { gridSpan: span },
            content: [
              {
                type: "paragraph",
                paraId: "000010B0",
                content: [
                  {
                    type: "run",
                    content: [{ type: "text", text: "Reviewed text" }],
                  },
                ],
              },
            ],
            structuralChange: {
              type:
                decision === "accept"
                  ? "tableCellDeletion"
                  : "tableCellInsertion",
              info: { id: 7, author: "Reviewer" },
            },
          } satisfies TableCell;
          const table = {
            type: "table",
            columnWidths: Array.from({ length: span + 3 }, () => 600),
            rows: [
              {
                type: "tableRow",
                cells: [
                  { ...cell("000010A0"), formatting: { vMerge: "restart" } },
                  middle,
                  { ...cell("000010C0"), formatting: { vMerge: "restart" } },
                  cell("000010D0"),
                ],
              },
              {
                type: "tableRow",
                cells: [
                  { ...cell("000020A0"), formatting: { vMerge: "continue" } },
                  { ...cell("000020B0"), formatting: { gridSpan: span } },
                  { ...cell("000020C0"), formatting: { vMerge: "continue" } },
                  cell("000020D0"),
                ],
              },
            ],
          } satisfies Table;
          const source = documentOf(table);
          const settled = applyDocumentOp(source, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds: [7],
            decision,
          }).unwrap();
          const resolved = tableOf(settled.document);
          expect(resolved.rows.at(0)?.cells.at(1)).toEqual({
            ...cell("000010B0"),
            ...(span === 1 ? {} : { formatting: { gridSpan: span } }),
          });
          expect(
            tableGrid(resolved, DOCUMENT_OP_TYPES.SET_TABLE).unwrap().width,
          ).toBe(span + 3);
          expect(resolved.rows.at(0)?.cells.at(2)).toEqual(
            table.rows.at(0)?.cells.at(2),
          );
          expect(resolved.rows.at(1)).toEqual(table.rows.at(1));
          expect(
            applyDocumentOps(settled.document, settled.inverse).unwrap()
              .document,
          ).toStrictEqual(source);
        },
      ),
      {
        id: "selective cell review keeps a blank interior slot between vertical merges",
      },
    );
  },
  propertyTestTimeout(10_000),
);
