/**
 * A direct cell merge, cell split or column edit on a table with pending
 * review records keeps every record it does not remove with its cell, under
 * its id: accepting or rejecting them all afterwards gives the direct edit
 * applied to the accepted or rejected table, and the inverse restores the
 * table exactly. A tracked edit of the same table refuses. A record a direct
 * edit cannot keep exact refuses the edit, with the reason it names.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../../model/document";
import { applyDocumentOp } from "../apply";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  type TableEditOp,
} from "../types";

const INFO = { author: "Reviewer", date: "2026-01-01T00:00:00Z" };

/** The paragraph id of the cell in `row` and `column`. */
const at = (row: number, column: number) => `000${row}${column}0A0`;
const FRESH = Array.from(
  { length: 10 },
  (_, index) => `00F00${String(index)}A0`,
);

const paragraph = (paraId: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text: paraId }] }],
});
const cell = (paraId: string, extra: Partial<TableCell> = {}): TableCell => ({
  type: "tableCell",
  content: [paragraph(paraId)],
  ...extra,
});
const row = (cells: TableCell[], extra: Partial<TableRow> = {}): TableRow => ({
  type: "tableRow",
  cells,
  ...extra,
});

type CellExtra = (row: number, column: number) => Partial<TableCell>;
/** Three rows of three cells, then a row whose first cell spans two columns. */
const rows = (extra: CellExtra = () => ({})): TableRow[] => [
  ...[0, 1, 2].map((r) =>
    row([0, 1, 2].map((c) => cell(at(r, c), extra(r, c)))),
  ),
  row([
    cell(at(3, 0), {
      formatting: { gridSpan: 2, width: { type: "dxa", value: 4000 } },
      ...extra(3, 0),
    }),
    cell(at(3, 2), extra(3, 2)),
  ]),
];
const WIDTHS = [2000, 2000, 2000];

const OP_NAMES = [
  "mergeAcross",
  "mergeDown",
  "split",
  "insertColumn",
  "insertColumns",
  "deleteColumn",
] as const;
type OpName = (typeof OP_NAMES)[number];

const OPS = {
  mergeAcross: {
    type: DOCUMENT_OP_TYPES.MERGE_CELLS,
    top: 0,
    bottom: 1,
    left: 0,
    right: 2,
  },
  mergeDown: {
    type: DOCUMENT_OP_TYPES.MERGE_CELLS,
    top: 0,
    bottom: 2,
    left: 0,
    right: 1,
  },
  split: { type: DOCUMENT_OP_TYPES.SPLIT_CELL },
  insertColumn: {
    type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
    column: 1,
    width: 1500,
  },
  insertColumns: {
    type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
    column: 1,
    width: 1500,
    count: 2,
  },
  deleteColumn: { type: DOCUMENT_OP_TYPES.DELETE_COLUMN, column: 0 },
} as const satisfies Record<OpName, { type: TableEditOp["type"] }>;

const opFor = (name: OpName, table: Table): TableEditOp => {
  const spec = OPS[name];
  const address = { story: OP_STORIES.MAIN, blockId: at(0, 0) };
  switch (spec.type) {
    case DOCUMENT_OP_TYPES.MERGE_CELLS:
      return {
        ...address,
        ...spec,
        newBlockIds: FRESH.slice(0, spec.bottom - spec.top - 1),
      };
    case DOCUMENT_OP_TYPES.SPLIT_CELL:
      return {
        ...spec,
        story: OP_STORIES.MAIN,
        blockId: at(3, 0),
        newBlockIds: FRESH.slice(0, 1),
      };
    case DOCUMENT_OP_TYPES.INSERT_COLUMN:
      // Every row takes a new cell but one whose first cell the column crosses.
      return {
        ...address,
        ...spec,
        newBlockIds: FRESH.slice(
          0,
          table.rows.filter(
            ({ cells }) => (cells.at(0)?.formatting?.gridSpan ?? 1) === 1,
          ).length * ("count" in spec ? spec.count : 1),
        ),
      };
    case DOCUMENT_OP_TYPES.DELETE_COLUMN:
      return { ...address, ...spec };
    default: {
      const unreachable: never = spec;
      return unreachable;
    }
  }
};

type Outcome =
  | {
      type: "applies";
      /** The record ids that go with the cells the edit removes. */
      removed?: readonly number[];
    }
  | { type: "refused"; reason: string; message: string };
const APPLIES = { type: "applies" } as const satisfies Outcome;
const refusedAs = (
  message: string,
  reason: string = DOCUMENT_OP_REFUSAL_REASONS.PENDING_TABLE_CHANGE,
) => ({ type: "refused", reason, message }) as const satisfies Outcome;

type Shape = { table: Table; outcome: Record<OpName, Outcome> };

const ALL_APPLY = {
  mergeAcross: APPLIES,
  mergeDown: APPLIES,
  split: APPLIES,
  insertColumn: APPLIES,
  insertColumns: APPLIES,
  deleteColumn: APPLIES,
} as const satisfies Record<OpName, Outcome>;

/** A pending insertion of the second row and deletion of the third. */
const ROW_CHANGES: readonly (TableRow["structuralChange"] | undefined)[] = [
  undefined,
  { type: "tableRowInsertion", info: { id: 6, ...INFO } },
  { type: "tableRowDeletion", info: { id: 7, ...INFO } },
];

const SHAPES = {
  "a pending table property change": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      formatting: { justification: "center" },
      propertyChanges: [
        {
          type: "tablePropertyChange",
          info: { id: 9, ...INFO },
          previousFormatting: {},
        },
      ],
      rows: rows(),
    },
    outcome: ALL_APPLY,
  },
  "a pending grid change": {
    table: {
      type: "table",
      columnWidths: [2000, 3000, 4000],
      formatting: { gridChange: { id: 9, columnWidths: [3000, 3000, 3000] } },
      rows: rows(),
    },
    outcome: ALL_APPLY,
  },
  "a pending row insertion and a pending row deletion": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      rows: rows().map((each, index) => {
        const structuralChange = ROW_CHANGES[index];
        return structuralChange === undefined
          ? each
          : Object.assign({}, each, { structuralChange });
      }),
    },
    // Merging a row that stands with one that a rejection would remove
    // would keep the removed row's words in the row that stands.
    outcome: {
      ...ALL_APPLY,
      mergeDown: refusedAs(
        "Resolve the pending row insertions and deletions the merge crosses before merging.",
      ),
    },
  },
  "pending cell and row property changes": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      rows: [
        ...rows((r, c) => {
          if (r === 0 && c === 0)
            return {
              formatting: { verticalAlign: "center" },
              propertyChanges: [
                {
                  type: "tableCellPropertyChange",
                  info: { id: 9, ...INFO },
                  previousFormatting: {},
                },
              ],
            };
          if (r === 3 && c === 0)
            return {
              formatting: {
                gridSpan: 2,
                width: { type: "dxa", value: 4000 },
                verticalAlign: "bottom",
              },
              propertyChanges: [
                {
                  type: "tableCellPropertyChange",
                  info: { id: 10, ...INFO },
                  previousFormatting: {
                    gridSpan: 2,
                    width: { type: "dxa", value: 4000 },
                  },
                },
              ],
            };
          return {};
        }),
        row([cell(at(4, 1)), cell(at(4, 2))], {
          formatting: { gridBefore: 1, cantSplit: true },
          propertyChanges: [
            {
              type: "tableRowPropertyChange",
              info: { id: 11, ...INFO },
              previousFormatting: { gridBefore: 1 },
            },
          ],
        }),
      ],
    },
    outcome: {
      ...ALL_APPLY,
      // The split's new cell takes the cell's pending vertical alignment,
      // which only a record of its own could keep pending.
      split: refusedAs(
        "Resolve the pending property changes of the split cell before splitting its columns.",
      ),
      deleteColumn: { type: "applies", removed: [9] },
    },
  },
  "a pending column deletion": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      rows: rows((r, c) =>
        c === 2
          ? {
              structuralChange: {
                type: "tableCellDeletion",
                info: { id: 20 + r, ...INFO },
              },
            }
          : {},
      ),
    },
    outcome: ALL_APPLY,
  },
  "a pending column insertion": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      formatting: { gridChange: { id: 9, columnWidths: [2000, 2000] } },
      rows: rows((r, c) =>
        c === 2
          ? {
              structuralChange: {
                type: "tableCellInsertion",
                info: { id: 20 + r, ...INFO },
              },
            }
          : {},
      ),
    },
    outcome: ALL_APPLY,
  },
  // The cell a split inserted beside the cell it narrowed: rejecting them
  // widens the cell back, so the row stands on another grid.
  "a pending cell insertion that takes part of a column": {
    table: {
      type: "table",
      columnWidths: WIDTHS,
      rows: [
        ...rows().slice(0, 3),
        row([
          cell(at(3, 0), {
            formatting: { width: { type: "dxa", value: 2000 } },
            propertyChanges: [
              {
                type: "tableCellPropertyChange",
                info: { id: 10, ...INFO },
                previousFormatting: {
                  gridSpan: 2,
                  width: { type: "dxa", value: 4000 },
                },
              },
            ],
          }),
          cell(at(3, 1), {
            formatting: { width: { type: "dxa", value: 2000 } },
            structuralChange: {
              type: "tableCellInsertion",
              info: { id: 11, ...INFO },
            },
          }),
          cell(at(3, 2)),
        ]),
      ],
    },
    outcome: {
      ...ALL_APPLY,
      split: refusedAs(
        "A one-row, one-column cell has nothing to split.",
        DOCUMENT_OP_REFUSAL_REASONS.NO_CHANGE,
      ),
      insertColumn: refusedAs(
        "Resolve the pending cell insertions and deletions that take part of a column before changing its columns.",
      ),
      insertColumns: refusedAs(
        "Resolve the pending cell insertions and deletions that take part of a column before changing its columns.",
      ),
      deleteColumn: refusedAs(
        "Resolve the pending cell insertions and deletions that take part of a column before changing its columns.",
      ),
    },
  },
} as const satisfies Record<string, Shape>;

const documentOf = (table: Table): Document => ({
  package: {
    document: { content: [table, paragraph("00000001")] },
  },
});
const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(0);
  if (table?.type !== "table") throw new Error("The table stays first.");
  return table;
};

/** Every revision id the table's records hold. */
const revisionIds = (node: unknown): number[] => {
  if (Array.isArray(node)) return node.flatMap(revisionIds);
  if (node === null || typeof node !== "object") return [];
  const own: number[] = [];
  const info: unknown = Reflect.get(node, "info");
  if (info !== null && typeof info === "object") {
    const id: unknown = Reflect.get(info, "id");
    if (typeof id === "number") own.push(id);
  }
  const gridChange: unknown = Reflect.get(node, "gridChange");
  if (gridChange !== null && typeof gridChange === "object") {
    const id: unknown = Reflect.get(gridChange, "id");
    if (typeof id === "number") own.push(id);
  }
  return [
    ...own,
    ...Object.entries(node)
      .filter(([key]) => key !== "info" && key !== "gridChange")
      .flatMap(([, value]) => revisionIds(value)),
  ];
};

/** The table with every new paragraph's id read as one: a row a decision removes takes its own. */
const withoutFreshIds = (table: Table): unknown =>
  JSON.parse(JSON.stringify(table).replaceAll(/"00F00\dA0"/g, `"new"`));

const apply = (document: Document, op: DocumentOp) => {
  const result = applyDocumentOp(document, op);
  if (result.isErr()) throw result.error;
  return result.value;
};

const resolveAll = (document: Document, decision: "accept" | "reject") =>
  apply(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...new Set(revisionIds(document))],
    decision,
  }).document;

const ascending = (left: number, right: number) => left - right;

describe.each(Object.entries(SHAPES))(
  "a table with %s",
  (_, { table, outcome }) => {
    test.each(OP_NAMES)("a direct %s", (opName) => {
      const original = documentOf(structuredClone(table));
      const op = opFor(opName, table);
      const expected = outcome[opName];
      const snapshot = structuredClone(original);
      const result = applyDocumentOp(original, op);
      expect(original).toStrictEqual(snapshot);
      if (expected.type === "refused") {
        expect(result.isErr() && result.error.reason).toBe(expected.reason);
        expect(result.isErr() && result.error.message).toBe(expected.message);
        return;
      }
      if (result.isErr()) throw result.error;
      const edited = result.value.document;
      const removed = new Set(expected.removed ?? []);
      expect(revisionIds(edited).toSorted(ascending)).toEqual(
        revisionIds(original)
          .filter((id) => !removed.has(id))
          .toSorted(ascending),
      );
      for (const decision of ["accept", "reject"] as const) {
        const resolved = resolveAll(edited, decision);
        expect(revisionIds(resolved)).toEqual([]);
        const before = resolveAll(original, decision);
        const want = apply(before, opFor(opName, tableOf(before))).document;
        expect(withoutFreshIds(tableOf(resolved))).toEqual(
          withoutFreshIds(tableOf(want)),
        );
      }
      let undone = edited;
      for (const inverse of result.value.inverse)
        undone = apply(undone, inverse).document;
      expect(undone).toStrictEqual(original);
    });

    test("a tracked merge refuses", () => {
      const result = applyDocumentOp(documentOf(structuredClone(table)), {
        ...opFor("mergeDown", table),
        revision: { id: 900, ...INFO },
      });
      expect(result.isErr() && result.error.message).toBe(
        "Resolve table reviews before changing cells.",
      );
    });
  },
);
