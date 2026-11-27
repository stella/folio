/**
 * Every table operation, on random tables with merged cells, through the
 * public reviewer, applied directly and tracked.
 *
 * A table operation's own tests build one table, apply one operation, and
 * read the live editor state back. That checks the operation against what it
 * meant to do, not against what the rest of the system makes of the result,
 * and every table-operation defect so far lived in the gap: a value the
 * operation dropped while reporting success, a row the live model held and
 * the reopened package did not, a tracked result whose accept left a
 * different table from the direct one. So the properties here are about the
 * gap, for any table and any operation:
 *
 *   (a) every value an applied operation supplied is in the saved document;
 *   (b) the direct result equals the tracked result accepted, or both refuse;
 *   (c) the tracked result rejected equals the original;
 *   (d) every saved table is well formed, and the live snapshot, the reopened
 *       snapshot, the package and its Markdown agree on every cell's spans.
 *
 * Each result is saved and reopened before it is read, and a tracked one is
 * resolved only after a save and reopen, as a reviewer receiving the file
 * would. Tables mix horizontal and vertical merges, and one cell may hold a
 * nested table the operation can target instead.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../test/property-testing";
import {
  buildTableDocx,
  readReviewerTables,
  type TableReading,
  tableReadingProblems,
  type TableSpec,
  type TableSpecCell,
} from "../__tests__/tableOperationDocument";
import {
  FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
  type FolioDocumentOperation,
  type FolioDocumentOperationResult,
} from "../document-operations";
import { FolioDocxReviewer } from "./headless";
import type { FolioAIBlock } from "./types";

/** A drawn cell size, and whether the cell is left blank (a tracked merge takes only blank cells). */
type Rectangle = { rowSpan: number; columnSpan: number; blank: boolean };

/**
 * A partition of a `rows` x `columns` grid into merged cells, filled row-major:
 * each free slot starts a cell as wide and as tall as the drawn sizes allow
 * without covering a slot already taken. Every row starts at least one cell,
 * because a package row of nothing but continuations is not one the readers
 * agree on (see `checkTableGrid`), so a table like that is not an input.
 */
const partition = (
  rows: number,
  columns: number,
  sizes: readonly Rectangle[],
  label: string,
): TableSpec | null => {
  const taken = new Set<string>();
  const cells: TableSpecCell[] = [];
  let draw = 0;
  for (let row = 0; row < rows; row++) {
    for (let column = 0; column < columns; column++) {
      if (taken.has(`${row}:${column}`)) {
        continue;
      }
      const size = sizes[draw % sizes.length] ?? {
        rowSpan: 1,
        columnSpan: 1,
        blank: false,
      };
      draw += 1;
      let columnSpan = 1;
      while (
        columnSpan < size.columnSpan &&
        column + columnSpan < columns &&
        !taken.has(`${row}:${column + columnSpan}`)
      ) {
        columnSpan += 1;
      }
      const rowSpan = Math.min(size.rowSpan, rows - row);
      for (let r = row; r < row + rowSpan; r++) {
        for (let c = column; c < column + columnSpan; c++) {
          taken.add(`${r}:${c}`);
        }
      }
      cells.push({
        row,
        column,
        rowSpan,
        columnSpan,
        text: size.blank ? "" : `${label}${cells.length}`,
      });
    }
  }
  for (let row = 0; row < rows; row++) {
    if (!cells.some((cell) => cell.row === row)) {
      return null;
    }
  }
  return { rows, columns, cells };
};

const rectangleArbitrary = fc.record({
  rowSpan: fc.oneof(
    { weight: 3, arbitrary: fc.constant(1) },
    { weight: 2, arbitrary: fc.integer({ min: 2, max: 3 }) },
  ),
  columnSpan: fc.oneof(
    { weight: 3, arbitrary: fc.constant(1) },
    { weight: 2, arbitrary: fc.integer({ min: 2, max: 3 }) },
  ),
  blank: fc.oneof(
    { weight: 3, arbitrary: fc.constant(false) },
    { weight: 1, arbitrary: fc.constant(true) },
  ),
});

const tableSpecArbitrary = (
  label: string,
  maxRows: number,
  maxColumns: number,
): fc.Arbitrary<TableSpec> =>
  fc
    .record({
      rows: fc.integer({ min: 1, max: maxRows }),
      columns: fc.integer({ min: 1, max: maxColumns }),
      sizes: fc.array(rectangleArbitrary, { minLength: 1, maxLength: 12 }),
    })
    .map(({ rows, columns, sizes }) => partition(rows, columns, sizes, label))
    .filter((spec): spec is TableSpec => spec !== null);

/** An outer table, one of whose cells may hold a nested table. */
const documentArbitrary: fc.Arbitrary<TableSpec> = fc
  .record({
    outer: tableSpecArbitrary("c", 4, 4),
    nested: fc.option(tableSpecArbitrary("n", 3, 3), { nil: undefined }),
    host: fc.nat(),
  })
  .map(({ outer, nested, host }) => {
    if (!nested) {
      return outer;
    }
    const cells = [...outer.cells];
    const index = host % cells.length;
    const hostCell = cells[index];
    if (hostCell) {
      cells[index] = Object.assign({}, hostCell, { nested });
    }
    return { rows: outer.rows, columns: outer.columns, cells };
  });

type OperationKind =
  | "insertTableRow"
  | "deleteTableRow"
  | "insertTableColumn"
  | "deleteTableColumn"
  | "mergeTableCells"
  | "splitTableCell"
  | "replaceBlock"
  | "insertAfterBlock"
  | "deleteBlock";

/** An operation, with its anchors named as cells of the spec until the document is open. */
type OperationPlan = {
  anchor: number;
  other: number;
  kind: OperationKind;
  position: "before" | "after";
  valueCount: number | undefined;
};

const planArbitrary: fc.Arbitrary<OperationPlan> = fc.record({
  anchor: fc.nat(),
  other: fc.nat(),
  kind: fc.constantFrom<OperationKind>(
    "insertTableRow",
    "deleteTableRow",
    "insertTableColumn",
    "deleteTableColumn",
    "mergeTableCells",
    "splitTableCell",
    "replaceBlock",
    "insertAfterBlock",
    "deleteBlock",
  ),
  position: fc.constantFrom("before", "after"),
  valueCount: fc.option(fc.integer({ min: 0, max: 5 }), { nil: undefined }),
});

/** A document whose only table cell holds `table`. */
const onlyCellOf = (table: TableSpec): TableSpec => ({
  rows: 1,
  columns: 1,
  cells: [
    { row: 0, column: 0, rowSpan: 1, columnSpan: 1, text: "", nested: table },
  ],
});

/** A table of two to `maxRows` rows with at least one vertical merge. */
const verticallyMergedTableArbitrary = (
  label: string,
  maxRows: number,
): fc.Arbitrary<TableSpec> =>
  fc
    .record({
      rows: fc.integer({ min: 2, max: maxRows }),
      columns: fc.integer({ min: 1, max: 3 }),
      sizes: fc.array(
        fc.record({
          rowSpan: fc.integer({ min: 1, max: 3 }),
          columnSpan: fc.oneof(
            { weight: 3, arbitrary: fc.constant(1) },
            { weight: 1, arbitrary: fc.constant(2) },
          ),
          blank: fc.boolean(),
        }),
        { minLength: 1, maxLength: 8 },
      ),
    })
    .map(({ rows, columns, sizes }) => partition(rows, columns, sizes, label))
    .filter(
      (spec): spec is TableSpec =>
        spec?.cells.some((cell) => cell.rowSpan > 1) === true,
    );

/**
 * A table built around vertical merges, top-level or as the only cell of an
 * outer table, so every anchor but one lands in it.
 */
const mergedDocumentArbitrary: fc.Arbitrary<TableSpec> = fc
  .record({
    table: verticallyMergedTableArbitrary("n", 3),
    nested: fc.boolean(),
  })
  .map(({ table, nested }) => (nested ? onlyCellOf(table) : table));

/**
 * As {@link mergedDocumentArbitrary}, with taller tables, and also nested in
 * a cell of an outer table with vertical merges of its own.
 */
const nestedMergedDocumentArbitrary: fc.Arbitrary<TableSpec> = fc
  .record({
    table: verticallyMergedTableArbitrary("n", 4),
    outer: verticallyMergedTableArbitrary("c", 3),
    nesting: fc.constantFrom("none", "only", "beside"),
    host: fc.nat(),
  })
  .map(({ table, outer, nesting, host }): TableSpec => {
    if (nesting === "none") return table;
    if (nesting === "only") return onlyCellOf(table);
    const cells = [...outer.cells];
    const index = host % cells.length;
    const hostCell = cells[index];
    if (hostCell) cells[index] = { ...hostCell, nested: table };
    return { rows: outer.rows, columns: outer.columns, cells };
  });

const batchPlanArbitrary = (
  kinds: readonly { weight: number; kinds: OperationKind[] }[],
  maxLength: number,
): fc.Arbitrary<OperationPlan[]> =>
  fc.array(
    fc.record({
      anchor: fc.nat(),
      other: fc.nat(),
      kind: fc.oneof(
        ...kinds.map(({ weight, kinds: choices }) => ({
          weight,
          arbitrary: fc.constantFrom(...choices),
        })),
      ),
      position: fc.constantFrom("before", "after"),
      valueCount: fc.option(fc.integer({ min: 0, max: 3 }), {
        nil: undefined,
      }),
    }),
    { minLength: 2, maxLength },
  );

/** Two or three edits of one merged table, mostly of its rows. */
const mergedBatchArbitrary = batchPlanArbitrary(
  [
    { weight: 4, kinds: ["insertTableRow", "deleteTableRow"] },
    {
      weight: 1,
      kinds: [
        "insertTableColumn",
        "deleteTableColumn",
        "mergeTableCells",
        "splitTableCell",
        "replaceBlock",
        "deleteBlock",
      ],
    },
  ],
  3,
);

/** Two to four edits of merged tables, rows, columns and cells mixed. */
const mixedBatchArbitrary = batchPlanArbitrary(
  [
    { weight: 3, kinds: ["insertTableRow", "deleteTableRow"] },
    { weight: 2, kinds: ["insertTableColumn", "deleteTableColumn"] },
    {
      weight: 2,
      kinds: [
        "mergeTableCells",
        "splitTableCell",
        "replaceBlock",
        "insertAfterBlock",
        "deleteBlock",
      ],
    },
  ],
  4,
);

/**
 * One operation, or two in one batch: each is resolved against the document
 * as it was read, and the batch applies them from the end backwards, so a
 * second operation at the first one's boundary checks that neither lands on
 * what the other moved.
 */
const batchArbitrary = fc.array(planArbitrary, { minLength: 1, maxLength: 2 });

/** A cell of the spec, and the snapshot table it is read from (0 outer, 1 nested). */
type AnchorCell = {
  cell: TableSpecCell;
  tableIndex: number;
  siblings: readonly TableSpecCell[];
};

const anchorCells = (spec: TableSpec): AnchorCell[] =>
  spec.cells.flatMap((cell) => [
    { cell, tableIndex: 0, siblings: spec.cells },
    ...(cell.nested?.cells.map((nested) => ({
      cell: nested,
      tableIndex: 1,
      siblings: cell.nested?.cells ?? [],
    })) ?? []),
  ]);

type PlannedOperation = {
  operation: FolioDocumentOperation;
  /** The texts the operation supplies, each of which must reach the document. */
  values: string[];
  anchor: AnchorCell;
  other: TableSpecCell;
};

/**
 * The operation a plan names. Anchors are the first paragraph of a cell, found
 * by where the snapshot says the cell is rather than by its text, because a
 * cell can be blank.
 */
const buildOperation = (
  spec: TableSpec,
  plan: OperationPlan,
  blocks: readonly FolioAIBlock[],
  id: string,
): PlannedOperation => {
  const anchors = anchorCells(spec);
  const anchor = anchors[plan.anchor % anchors.length]!;
  const other = anchor.siblings[plan.other % anchor.siblings.length]!;
  const blockIdAt = (tableIndex: number, cell: TableSpecCell): string => {
    const block = blocks.find(
      ({ table }) =>
        table?.tableIndex === tableIndex &&
        table.rowIndex === cell.row &&
        table.gridColumnIndex === cell.column &&
        table.paragraphIndex === 0,
    );
    if (!block) {
      throw new Error(
        `no block starts cell ${cell.row}:${cell.column} of table ${tableIndex}`,
      );
    }
    return block.id;
  };
  const blockId = blockIdAt(anchor.tableIndex, anchor.cell);
  const values = Array.from(
    { length: plan.valueCount ?? 0 },
    (_, index) => `${id}v${index}v`,
  );
  const planned = (
    operation: FolioDocumentOperation,
    supplied: string[] = [],
  ) => ({
    operation,
    values: supplied,
    anchor,
    other,
  });
  switch (plan.kind) {
    case "insertTableRow":
    case "insertTableColumn":
      return planned(
        {
          id,
          type: plan.kind,
          blockId,
          position: plan.position,
          ...(plan.valueCount !== undefined && { cellTexts: values }),
        },
        values,
      );
    case "deleteTableRow":
    case "deleteTableColumn":
    case "splitTableCell":
    case "deleteBlock":
      return planned({ id, type: plan.kind, blockId });
    case "mergeTableCells":
      return planned({
        id,
        type: plan.kind,
        blockId,
        endBlockId: blockIdAt(anchor.tableIndex, other),
      });
    case "replaceBlock": {
      const text = `${id}replacedv`;
      return planned({ id, type: plan.kind, blockId, text }, [text]);
    }
    case "insertAfterBlock": {
      const text = `${id}insertedv`;
      return planned({ id, type: plan.kind, blockId, text }, [text]);
    }
  }
};

const SESSIONS = ["canonical"] as const;

describe.each(SESSIONS)("%s session", (session) => {
  const open = (bytes: ArrayBuffer) =>
    FolioDocxReviewer.fromBuffer(bytes, { author: "Tester", session });

  type Run = {
    planned: PlannedOperation[];
    reviewer: FolioDocxReviewer;
    result: FolioDocumentOperationResult;
  };

  /** The plans as one batch, each resolved against the document as it was read. */
  const run = async (
    base: ArrayBuffer,
    spec: TableSpec,
    plans: readonly OperationPlan[],
    mode: "direct" | "tracked-changes",
  ): Promise<Run> => {
    const reviewer = await open(base);
    const blocks = reviewer.getContent();
    const planned = plans.map((plan, index) =>
      buildOperation(spec, plan, blocks, `op${index}`),
    );
    const result = reviewer.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode,
        operations: planned.map(({ operation }) => operation),
      },
      { undefinedReferences: "refuse" },
    );
    return { planned, reviewer, result };
  };

  const appliedIds = ({ result }: Run): Set<string> =>
    new Set(result.applied.map(({ id }) => id));

  const readTablesAt = async (
    stage: string,
    reviewer: FolioDocxReviewer,
  ): Promise<TableReading> => {
    try {
      return await readReviewerTables(reviewer);
    } catch (cause) {
      throw new Error(`Table property failed while reading ${stage}`, {
        cause,
      });
    }
  };

  const resolve = async (
    reviewer: FolioDocxReviewer,
    resolution: "accept" | "reject",
  ): Promise<TableReading> => {
    const reopened = await open(await reviewer.toBuffer());
    if (resolution === "accept") {
      reopened.acceptAll();
    } else {
      reopened.rejectAll();
    }
    return readTablesAt(resolution, reopened);
  };

  /** What a caller sees of a document: its blocks' text and every table's cells. */
  const visible = (reading: TableReading) => ({
    texts: reading.texts,
    tables: reading.snapshot,
  });

  /**
   * Whether tracked mode may refuse what direct mode applies. Each case is a
   * shape the tracked-change vocabulary cannot record, and nothing else is:
   *
   * - A merge tracks paragraph moves and geometry. The generator also puts
   *   nested tables and paragraphs with pending edits in cells that would
   *   need to move as complete blocks.
   * - A split is the reverse, so only a vertical merge splits tracked.
   * Pending table reviews also block edits that could overwrite their record
   * or give it a different effect. Those refusals use a separate typed reason.
   */
  const trackedRefusalIsExpected = ({
    planned,
    plan,
    paragraphEdits,
  }: {
    planned: PlannedOperation;
    plan: OperationPlan;
    paragraphEdits: ReadonlySet<TableSpecCell>;
  }): boolean => {
    const { anchor, other } = planned;
    const cell = anchor.cell;
    const siblings = anchor.siblings;
    switch (plan.kind) {
      case "mergeTableCells": {
        const top = Math.min(cell.row, other.row);
        const bottom = Math.max(
          cell.row + cell.rowSpan,
          other.row + other.rowSpan,
        );
        const left = Math.min(cell.column, other.column);
        const right = Math.max(
          cell.column + cell.columnSpan,
          other.column + other.columnSpan,
        );
        return siblings.some(
          (candidate) =>
            (candidate.nested !== undefined || paragraphEdits.has(candidate)) &&
            (candidate.row !== top || candidate.column !== left) &&
            candidate.row < bottom &&
            candidate.row + candidate.rowSpan > top &&
            candidate.column < right &&
            candidate.column + candidate.columnSpan > left,
        );
      }
      case "splitTableCell":
        return cell.columnSpan > 1;
      default:
        return false;
    }
  };

  const checkCase = async (
    spec: TableSpec,
    plans: readonly OperationPlan[],
  ): Promise<void> => {
    const base = await buildTableDocx(spec);
    const original = await readTablesAt("original", await open(base));
    expect(tableReadingProblems(original)).toEqual([]);

    const direct = await run(base, spec, plans, "direct");
    const tracked = await run(base, spec, plans, "tracked-changes");
    const directApplied = appliedIds(direct);
    const trackedApplied = appliedIds(tracked);

    const operationIds = direct.planned.map(({ operation }) => operation.id);
    for (const { result } of [direct, tracked]) {
      const skippedIds = new Set(result.skipped.map(({ id }) => id));
      expect(result.skipped.map(({ id }) => id)).toEqual(
        operationIds.filter((id) => skippedIds.has(id)),
      );
    }

    const directReading = await readTablesAt("direct", direct.reviewer);
    expect(tableReadingProblems(directReading)).toEqual([]);
    if (directApplied.size > 0) {
      // (a) Every value an applied operation supplied landed.
      const text = directReading.texts.join("\n");
      for (const { operation, values } of direct.planned) {
        for (const value of directApplied.has(operation.id) ? values : []) {
          expect(text).toContain(value);
        }
      }
    } else {
      // A refused batch leaves the document alone, and says why.
      expect(direct.result.issues).toHaveLength(plans.length);
      expect(visible(directReading)).toEqual(visible(original));
    }

    const rejected = await resolve(tracked.reviewer, "reject");
    expect(tableReadingProblems(rejected)).toEqual([]);
    // (c) Whatever tracked mode did, rejecting it restores the original.
    expect(visible(rejected)).toEqual(visible(original));

    const disagreeing = tracked.planned.flatMap((planned, index) =>
      directApplied.has(planned.operation.id) ===
      trackedApplied.has(planned.operation.id)
        ? []
        : [{ planned, plan: plans[index]! }],
    );
    if (disagreeing.length > 0) {
      // (b) Both modes agree on whether each operation is allowed, bar the
      // refusals the tracked vocabulary forces.
      const paragraphEdits = new Set(
        tracked.planned.flatMap(({ operation, anchor }) =>
          trackedApplied.has(operation.id) &&
          (operation.type === "deleteBlock" ||
            operation.type === "replaceBlock" ||
            operation.type === "insertAfterBlock")
            ? [anchor.cell]
            : [],
        ),
      );
      for (const { planned, plan } of disagreeing) {
        const id = planned.operation.id;
        const pendingTableReview =
          session === "canonical" &&
          tracked.planned.some((edit) => {
            const { operation } = edit;
            if (
              !trackedApplied.has(operation.id) ||
              edit.anchor.tableIndex !== planned.anchor.tableIndex
            )
              return false;
            const type = operation.type;
            const column =
              type === "insertTableColumn" || type === "deleteTableColumn";
            const cell =
              type === "splitTableCell" || type === "mergeTableCells";
            if (
              plan.kind === "insertTableColumn" ||
              plan.kind === "deleteTableColumn"
            )
              return column || cell;
            const row = type === "insertTableRow" || type === "deleteTableRow";
            if (plan.kind === "splitTableCell") return column || cell || row;
            if (plan.kind !== "mergeTableCells") return false;
            if (column || cell) return true;
            if (!row) return false;
            const top = Math.min(planned.anchor.cell.row, planned.other.row);
            const bottom = Math.max(
              planned.anchor.cell.row + planned.anchor.cell.rowSpan,
              planned.other.row + planned.other.rowSpan,
            );
            const anchorRow = edit.anchor.cell.row;
            if (type === "deleteTableRow")
              return anchorRow >= top && anchorRow < bottom;
            const insertedAt =
              anchorRow + (operation.position === "after" ? 1 : 0);
            return insertedAt > top && insertedAt < bottom;
          });
        expect({
          plans,
          direct: direct.result.skipped,
          tracked: tracked.result.skipped,
          expected:
            directApplied.has(id) &&
            tracked.result.skipped.some(
              (skip) =>
                skip.id === id &&
                ((skip.reason === "pendingTableChange" && pendingTableReview) ||
                  (skip.reason === "unsupportedBlock" &&
                    trackedRefusalIsExpected({
                      planned,
                      plan,
                      paragraphEdits,
                    }))),
            ),
        }).toMatchObject({ expected: true });
      }
    }
    if (trackedApplied.size === 0) {
      if (disagreeing.length === 0) {
        expect(direct.result.skipped.map(({ reason }) => reason)).toEqual(
          tracked.result.skipped.map(({ reason }) => reason),
        );
      }
      return;
    }
    const accepted = await resolve(tracked.reviewer, "accept");
    expect(tableReadingProblems(accepted)).toEqual([]);
    // (b) Accepting the tracked result is the direct result for the operations
    // both modes applied, even when tracked mode refused another operation.
    if (disagreeing.length === 0) {
      expect(visible(accepted)).toEqual(visible(directReading));
      return;
    }
    const shared = await open(base);
    const sharedOperations = direct.planned
      .filter(({ operation }) => trackedApplied.has(operation.id))
      .map(({ operation }) => operation);
    const sharedResult = shared.applyDocumentOperations(
      {
        version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
        mode: "direct",
        operations: sharedOperations,
      },
      { undefinedReferences: "refuse" },
    );
    expect(new Set(sharedResult.applied.map(({ id }) => id))).toEqual(
      trackedApplied,
    );
    const sharedReading = await readTablesAt("direct shared", shared);
    expect(tableReadingProblems(sharedReading)).toEqual([]);
    expect(visible(accepted)).toEqual(visible(sharedReading));
  };

  describe("table operations on merged tables", () => {
    test(
      "keep every value, agree across modes and readers, and reject cleanly",
      async () => {
        // The seeds this property once failed on replay first, from
        // test/property-seeds/.
        await assertProperty(
          fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
          {
            numRuns: 150,
            id: "keep every value, agree across modes and readers, and reject cleanly",
          },
        );
      },
      propertyTestTimeout(120_000),
    );

    // Row edits beside, inside and at the ends of vertical merges, batched so
    // one edit lands where another changed a merge.
    test(
      "keep vertical merges whole across batched row edits",
      async () => {
        await assertProperty(
          fc.asyncProperty(
            mergedDocumentArbitrary,
            mergedBatchArbitrary,
            checkCase,
          ),
          {
            numRuns: 100,
            id: "keep vertical merges whole across batched row edits",
          },
        );
      },
      propertyTestTimeout(120_000),
    );

    // Row, column and cell edits batched across merged tables at both
    // levels, so an edit lands where another changed a merge of either.
    test(
      "keep merges whole across mixed batches on nested merged tables",
      async () => {
        await assertProperty(
          fc.asyncProperty(
            nestedMergedDocumentArbitrary,
            mixedBatchArbitrary,
            checkCase,
          ),
          {
            numRuns: 100,
            id: "keep merges whole across mixed batches on nested merged tables",
          },
        );
      },
      propertyTestTimeout(120_000),
    );
  });
});
