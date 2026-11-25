/**
 * A direct cell merge, cell split or column edit on a table with pending
 * review records applies on both engines. On the canonical session it keeps
 * every record it does not remove with its cell, under its id: saved,
 * reopened and resolved, it gives the direct edit applied to the accepted or
 * rejected table, and undo restores the model. Where the default engine's
 * result differs, it is stated per engine. A tracked column insertion beside
 * the same records keeps the canonical refusal.
 */
import { describe, expect, test } from "bun:test";

import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { createDocx } from "../docx/rezip";
import type { HeadlessReviewerSession } from "../types/canonicalCapabilities";
import type {
  Document,
  Paragraph,
  Table,
  TableCell,
  TableRow,
} from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIEditOperation } from "./types";

const SESSIONS = [
  "canonical",
] as const satisfies readonly HeadlessReviewerSession[];

const AUTHOR = "Reviewer";
const OTHER = { author: "Other", date: "2026-01-01T00:00:00Z" };

/** The paragraph id of the cell in `row` and `column`. */
const at = (row: number, column: number) => `1000${row}${column}00`;

const paragraph = (paraId: string, text: string): Paragraph => ({
  type: "paragraph",
  paraId,
  content: [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (row: number, column: number, extra: Partial<TableCell> = {}) => ({
  type: "tableCell" as const,
  content: [paragraph(at(row, column), `${"ABC"[column]}${row}`)],
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
  ...[0, 1, 2].map((r) => row([0, 1, 2].map((c) => cell(r, c, extra(r, c))))),
  row([
    cell(3, 0, {
      formatting: { gridSpan: 2, width: { type: "dxa", value: 4000 } },
      ...extra(3, 0),
    }),
    cell(3, 2, extra(3, 2)),
  ]),
];

const OPERATIONS = {
  mergeAcross: {
    id: "op",
    type: "mergeTableCells",
    blockId: at(0, 0),
    endBlockId: at(0, 1),
  },
  mergeDown: {
    id: "op",
    type: "mergeTableCells",
    blockId: at(0, 0),
    rowCount: 2,
  },
  split: { id: "op", type: "splitTableCell", blockId: at(3, 0) },
  insertColumn: {
    id: "op",
    type: "insertTableColumn",
    blockId: at(0, 0),
    position: "after",
  },
  deleteColumn: { id: "op", type: "deleteTableColumn", blockId: at(0, 0) },
} as const satisfies Record<string, FolioAIEditOperation>;
type OperationName = keyof typeof OPERATIONS;
const OPERATION_NAMES = [
  "mergeAcross",
  "mergeDown",
  "split",
  "insertColumn",
  "deleteColumn",
] as const satisfies readonly OperationName[];

type Outcome =
  | {
      /** Applied; resolving all commutes with the edit; undo restores. */
      type: "exact";
      /** The record ids that go with the cells the edit removes. */
      removed?: readonly number[];
      /** The widths rejecting gives, where they are not the commuted ones. */
      rejectedWidths?: readonly number[];
    }
  /** Applied; how the result departs from an exact one is stated beside it. */
  | { type: "applied" }
  | { type: "refused"; message: string };
const EXACT = { type: "exact" } as const satisfies Outcome;
const APPLIED = { type: "applied" } as const satisfies Outcome;

type PerEngine = Record<
  OperationName,
  Record<HeadlessReviewerSession, Outcome>
>;
const everywhere = (outcome: Outcome) =>
  ({ legacy: outcome, canonical: outcome }) as const;
const ALL_EXACT = {
  mergeAcross: everywhere(EXACT),
  mergeDown: everywhere(EXACT),
  split: everywhere(EXACT),
  insertColumn: everywhere(EXACT),
  deleteColumn: everywhere(EXACT),
} as const satisfies PerEngine;
/** The default engine keeps a grid change pending through every decision. */
const LEGACY_KEEPS_GRID_CHANGE = {
  mergeAcross: { legacy: APPLIED, canonical: EXACT },
  mergeDown: { legacy: APPLIED, canonical: EXACT },
  split: { legacy: APPLIED, canonical: EXACT },
  insertColumn: { legacy: APPLIED, canonical: EXACT },
  deleteColumn: { legacy: APPLIED, canonical: EXACT },
} as const satisfies PerEngine;

const COLUMNS_REFUSED = "Resolve table reviews before changing its columns.";
const GRID_CHANGE_PENDING =
  "The table's columns already have a pending change; insert every column of one change in one operation.";

type Shape = {
  table: Table;
  outcome: PerEngine;
  /** A tracked column insertion beside the records, on the canonical session. */
  trackedColumn: { type: "applied" } | { type: "refused"; message: string };
};

/** A pending insertion of the second row and deletion of the third. */
const ROW_CHANGES: readonly (TableRow["structuralChange"] | undefined)[] = [
  undefined,
  { type: "tableRowInsertion", info: { id: 6, ...OTHER } },
  { type: "tableRowDeletion", info: { id: 7, ...OTHER } },
];

const SHAPES = {
  "a pending table property change": {
    table: {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      formatting: { justification: "center" },
      propertyChanges: [
        {
          type: "tablePropertyChange",
          info: { id: 9, ...OTHER },
          previousFormatting: {},
        },
      ],
      rows: rows(),
    },
    outcome: ALL_EXACT,
    trackedColumn: { type: "refused", message: COLUMNS_REFUSED },
  },
  "a pending grid change": {
    table: {
      type: "table",
      columnWidths: [2000, 3000, 4000],
      formatting: { gridChange: { id: 9, columnWidths: [3000, 3000, 3000] } },
      rows: rows(),
    },
    outcome: {
      ...LEGACY_KEEPS_GRID_CHANGE,
      // The new column keeps the width it was inserted with, its
      // neighbour's, when the earlier grid comes back.
      insertColumn: {
        legacy: APPLIED,
        canonical: { type: "exact", rejectedWidths: [3000, 2000, 3000, 3000] },
      },
    },
    trackedColumn: { type: "refused", message: GRID_CHANGE_PENDING },
  },
  "a pending row insertion and a pending row deletion": {
    table: {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      rows: rows().map((each, index) => {
        const structuralChange = ROW_CHANGES[index];
        return structuralChange === undefined
          ? each
          : Object.assign({}, each, { structuralChange });
      }),
    },
    outcome: {
      ...ALL_EXACT,
      // The default engine moves the inserted row's words into the row
      // above, where rejecting the row keeps them.
      mergeDown: {
        legacy: APPLIED,
        canonical: {
          type: "refused",
          message:
            "Resolve the pending row insertions and deletions the merge crosses before merging.",
        },
      },
    },
    // A row's own insertion or deletion keeps its mark beside a tracked
    // column.
    trackedColumn: { type: "applied" },
  },
  "pending cell property changes": {
    table: {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      rows: rows((r, c) => {
        if (r === 0 && c === 0)
          return {
            formatting: { verticalAlign: "center" },
            propertyChanges: [
              {
                type: "tableCellPropertyChange",
                info: { id: 9, ...OTHER },
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
                info: { id: 10, ...OTHER },
                previousFormatting: {
                  gridSpan: 2,
                  width: { type: "dxa", value: 4000 },
                },
              },
            ],
          };
        return {};
      }),
    },
    outcome: {
      ...ALL_EXACT,
      // The default engine gives the new cell a copy of the split cell's
      // record under the same id.
      split: {
        legacy: APPLIED,
        canonical: {
          type: "refused",
          message:
            "Resolve the pending property changes of the split cell before splitting its columns.",
        },
      },
      deleteColumn: everywhere({ type: "exact", removed: [9] }),
    },
    trackedColumn: { type: "refused", message: COLUMNS_REFUSED },
  },
  "a pending column deletion": {
    table: {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      rows: rows((r, c) =>
        c === 2
          ? {
              structuralChange: {
                type: "tableCellDeletion",
                info: { id: 20 + r, ...OTHER },
              },
            }
          : {},
      ),
    },
    outcome: ALL_EXACT,
    trackedColumn: { type: "refused", message: COLUMNS_REFUSED },
  },
  "a pending column insertion": {
    table: {
      type: "table",
      columnWidths: [2000, 2000, 2000],
      formatting: { gridChange: { id: 9, columnWidths: [2000, 2000] } },
      rows: rows((r, c) =>
        c === 2
          ? {
              structuralChange: {
                type: "tableCellInsertion",
                info: { id: 20 + r, ...OTHER },
              },
            }
          : {},
      ),
    },
    outcome: LEGACY_KEEPS_GRID_CHANGE,
    trackedColumn: { type: "refused", message: GRID_CHANGE_PENDING },
  },
} as const satisfies Record<string, Shape>;

const source = async (table: Table) => {
  const document: Document = createEmptyDocument();
  document.package.document.content = [
    paragraph("00000001", "Before."),
    structuredClone(table),
    paragraph("00000002", "After."),
  ];
  return await createDocx(document);
};

const open = (buffer: ArrayBuffer, session: HeadlessReviewerSession) =>
  FolioDocxReviewer.fromBuffer(buffer.slice(0), { author: AUTHOR, session });

const run = (
  reviewer: FolioDocxReviewer,
  operation: FolioAIEditOperation,
  mode: "direct" | "tracked-changes",
) =>
  reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode,
      operations: [operation],
    },
    { undefinedReferences: "refuse" },
  );

const textOf = (node: unknown): string => {
  if (Array.isArray(node)) return node.map(textOf).join("");
  if (node === null || typeof node !== "object") return "";
  if (Reflect.get(node, "type") === "text") {
    const text: unknown = Reflect.get(node, "text");
    return typeof text === "string" ? text : "";
  }
  return Object.values(node).map(textOf).join("");
};

/** What a decision leaves of the table: its widths, properties and cells. */
const tableShape = (reviewer: FolioDocxReviewer) => {
  const table = reviewer
    .toDocument()
    .package.document.content.find((block) => block.type === "table");
  if (table?.type !== "table") throw new Error("The table remains.");
  const {
    sourceXml: _source,
    gridSourceXml: _grid,
    ...formatting
  } = table.formatting ?? {};
  return {
    widths: table.columnWidths,
    formatting,
    rows: table.rows.map(({ formatting: rowFormatting, cells }) => ({
      formatting: rowFormatting,
      cells: cells.map(({ content, formatting: cellFormatting }) => {
        const { sourceXml: _cell, ...rest } = cellFormatting ?? {};
        return { text: textOf(content), formatting: rest };
      }),
    })),
  };
};

const changeIds = (reviewer: FolioDocxReviewer) =>
  reviewer
    .getChanges()
    .map(({ id }) => id)
    .toSorted((left, right) => left - right);

describe.each(SESSIONS)("%s: a direct table edit", (session) => {
  describe.each(Object.entries(SHAPES))(
    "beside %s",
    (_, { table, outcome }) => {
      test.each(OPERATION_NAMES)("%s", async (name) => {
        const operation = OPERATIONS[name];
        const expected = outcome[name][session];
        const buffer = await source(table);
        const reviewer = await open(buffer, session);
        const before = changeIds(reviewer);
        const model = reviewer.toDocument();
        const result = run(reviewer, operation, "direct");
        if (expected.type === "refused") {
          expect(result.applied).toEqual([]);
          expect(result.skipped.map(({ message }) => message)).toEqual([
            expected.message,
          ]);
          return;
        }
        expect(result.skipped).toEqual([]);
        expect(result.applied.map(({ id }) => id)).toEqual(["op"]);
        if (expected.type === "applied") return;
        const removed = new Set(expected.removed ?? []);
        expect(changeIds(reviewer)).toEqual(
          before.filter((id) => !removed.has(id)),
        );
        const saved = await reviewer.toBuffer();
        for (const decision of ["accept", "reject"] as const) {
          const resolved = await open(saved, session);
          if (decision === "accept") resolved.acceptAll();
          else resolved.rejectAll();
          expect(resolved.getChanges()).toEqual([]);
          const want = await open(buffer, session);
          if (decision === "accept") want.acceptAll();
          else want.rejectAll();
          const commuted = run(want, operation, "direct");
          expect(commuted.skipped).toEqual([]);
          const shape = tableShape(want);
          expect(tableShape(resolved)).toEqual(
            decision === "reject" && expected.rejectedWidths !== undefined
              ? { ...shape, widths: [...expected.rejectedWidths] }
              : shape,
          );
        }
        if (result.undoHandle === null) throw new Error("No undo handle.");
        expect(
          reviewer.undoDocumentOperations(result.undoHandle),
        ).toMatchObject({
          status: "undone",
        });
        expect(reviewer.toDocument()).toEqual(model);
      });
    },
  );
});

describe("canonical: a tracked column insertion", () => {
  test.each(Object.entries(SHAPES))(
    "beside %s keeps its outcome",
    async (_, { table, trackedColumn }) => {
      const reviewer = await open(await source(table), "canonical");
      // After the last column: a suggested column crossing a span refuses
      // before any record is read.
      const result = run(
        reviewer,
        { ...OPERATIONS.insertColumn, blockId: at(0, 2) },
        "tracked-changes",
      );
      expect({
        applied: result.applied.map(({ id }) => id),
        skipped: result.skipped.map(({ message }) => message),
      }).toEqual(
        trackedColumn.type === "applied"
          ? { applied: ["op"], skipped: [] }
          : { applied: [], skipped: [trackedColumn.message] },
      );
    },
  );
});
