import { panic } from "better-result";
import { expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Table } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { DOCUMENT_OP_TYPES } from "../types";

const cellArbitrary = fc.record({
  span: fc.integer({ min: 1, max: 4 }),
  merge: fc.option(fc.constantFrom("restart" as const, "continue" as const), {
    nil: undefined,
  }),
});

const irregularTable = fc
  .array(fc.array(cellArbitrary, { minLength: 1, maxLength: 3 }), {
    minLength: 1,
    maxLength: 3,
  })
  .map((rows): Table => {
    let id = 0;
    // This short first row guarantees the source is irregular; the others
    // exercise oversized spans and unmatched vertical continuations too.
    const cells = [[{ span: 1, merge: undefined }], ...rows];
    return {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      rows: cells.map((row) => ({
        type: "tableRow",
        cells: row.map(({ span, merge }) => ({
          type: "tableCell",
          formatting: {
            gridSpan: span,
            ...(merge === undefined ? {} : { vMerge: merge }),
          },
          content: [
            {
              type: "paragraph",
              paraId: (++id).toString(16).padStart(8, "0"),
              content: [],
            },
          ],
        })),
      })),
    };
  });

const operations = [
  {
    type: DOCUMENT_OP_TYPES.SET_TABLE_PROPS,
    patch: { justification: "center" },
  },
  { type: DOCUMENT_OP_TYPES.SET_ROW_PROPS, patch: { header: true } },
  {
    type: DOCUMENT_OP_TYPES.SET_CELL_PROPS,
    patch: { verticalAlign: "bottom" },
  },
] as const;

const tableOf = (document: Document) => {
  const table = document.package.document.content.at(0);
  return table?.type === "table" ? table : panic("The generated table exists.");
};

test(
  "property edits preserve every irregular table fact and have exact serialized inverses",
  () => {
    assertProperty(
      fc.property(irregularTable, (table) => {
        for (const historical of [false, true]) {
          const active = structuredClone(table);
          if (historical) {
            const cell =
              active.rows.at(0)?.cells.at(0) ?? panic("The first cell exists.");
            cell.propertyChanges = [
              {
                type: "tableCellPropertyChange",
                info: {
                  id: 17,
                  author: "Reviewer",
                  date: "2026-03-04T04:00:00Z",
                },
                previousFormatting: { gridSpan: 2 },
              },
            ];
          }
          const source: Document = {
            package: { document: { content: [active] } },
          };
          for (const operation of operations)
            for (const revision of [
              undefined,
              { id: 50, author: "Reviewer", date: "2026-03-04T05:00:00Z" },
            ]) {
              const applied = applyDocumentOp(source, {
                ...operation,
                story: "main",
                blockId: "00000001",
                ...(revision === undefined ? {} : { revision }),
              }).unwrap();
              const expected = structuredClone(source);
              const target = tableOf(expected);
              switch (operation.type) {
                case DOCUMENT_OP_TYPES.SET_TABLE_PROPS:
                  target.formatting = {
                    ...target.formatting,
                    ...operation.patch,
                  };
                  if (revision !== undefined)
                    target.propertyChanges = [
                      { type: "tablePropertyChange", info: revision },
                    ];
                  break;
                case DOCUMENT_OP_TYPES.SET_ROW_PROPS: {
                  const row =
                    target.rows.at(0) ?? panic("The first row exists.");
                  row.formatting = { ...row.formatting, ...operation.patch };
                  if (revision !== undefined)
                    row.propertyChanges = [
                      { type: "tableRowPropertyChange", info: revision },
                    ];
                  break;
                }
                case DOCUMENT_OP_TYPES.SET_CELL_PROPS: {
                  const cell =
                    target.rows.at(0)?.cells.at(0) ??
                    panic("The first cell exists.");
                  const pending = cell.propertyChanges?.at(0);
                  const previousFormatting =
                    pending === undefined
                      ? cell.formatting
                      : pending.previousFormatting;
                  cell.formatting = { ...cell.formatting, ...operation.patch };
                  if (revision !== undefined)
                    cell.propertyChanges = [
                      {
                        type: "tableCellPropertyChange",
                        info: revision,
                        ...(previousFormatting === undefined
                          ? {}
                          : { previousFormatting }),
                      },
                    ];
                  break;
                }
                default: {
                  const unreachable: never = operation;
                  return unreachable;
                }
              }
              expect(applied.document).toStrictEqual(expected);
              const inverse = applied.inverse.at(0);
              if (inverse?.type !== DOCUMENT_OP_TYPES.SET_TABLE)
                panic("A table property inverse restores the table.");
              for (const scope of ["table", "row", "cell"] as const) {
                const forged = structuredClone(inverse.table);
                const info = { id: 80, author: "Reviewer" };
                switch (scope) {
                  case "table":
                    forged.propertyChanges = [
                      {
                        type: "tablePropertyChange",
                        info,
                        previousFormatting: {
                          gridChange: { id: 81, columnWidths: [123] },
                        },
                      },
                    ];
                    break;
                  case "row": {
                    const row =
                      forged.rows.at(0) ?? panic("The first row exists.");
                    row.propertyChanges = [
                      {
                        type: "tableRowPropertyChange",
                        info,
                        currentFormatting: { gridBefore: 3 },
                      },
                    ];
                    break;
                  }
                  case "cell": {
                    const cell =
                      forged.rows.at(0)?.cells.at(0) ??
                      panic("The first cell exists.");
                    cell.propertyChanges = [
                      {
                        type: "tableCellPropertyChange",
                        info,
                        previousFormatting: { gridSpan: 99 },
                      },
                    ];
                    break;
                  }
                  default: {
                    const unreachable: never = scope;
                    return unreachable;
                  }
                }
                expect(
                  applyDocumentOp(applied.document, {
                    ...inverse,
                    table: forged,
                  }).isErr(),
                ).toBe(true);
              }
              const undo = applyDocumentOps(
                applied.document,
                JSON.parse(JSON.stringify(applied.inverse)),
              ).unwrap();
              expect(undo.document).toStrictEqual(source);
              expect(
                applyDocumentOps(undo.document, undo.inverse).unwrap().document,
              ).toStrictEqual(applied.document);
            }
          // A property edit never grants permission for a topology operation.
          expect(
            applyDocumentOp(source, {
              type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
              story: "main",
              blockId: "00000001",
              column: 0,
            }).isErr(),
          ).toBe(true);
        }
      }),
      {
        id: "property edits preserve irregular geometry and exact inverses",
        numRuns: 80,
      },
    );
  },
  propertyTestTimeout(30_000),
);
