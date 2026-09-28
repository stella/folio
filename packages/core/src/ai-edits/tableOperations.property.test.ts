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

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  buildTableDocx,
  readReviewerTables,
  tableReadingProblems,
  type TableReading,
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
      const size = sizes[draw % sizes.length] ?? { rowSpan: 1, columnSpan: 1, blank: false };
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

/**
 * One operation, or two in one batch: each is resolved against the document
 * as it was read, and the batch applies them from the end backwards, so a
 * second operation at the first one's boundary checks that neither lands on
 * what the other moved.
 */
const batchArbitrary = fc.array(planArbitrary, { minLength: 1, maxLength: 2 });

/** A cell of the spec, and the snapshot table it is read from (0 outer, 1 nested). */
type AnchorCell = { cell: TableSpecCell; tableIndex: number; siblings: readonly TableSpecCell[] };

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
      throw new Error(`no block starts cell ${cell.row}:${cell.column} of table ${tableIndex}`);
    }
    return block.id;
  };
  const blockId = blockIdAt(anchor.tableIndex, anchor.cell);
  const values = Array.from({ length: plan.valueCount ?? 0 }, (_, index) => `${id}v${index}v`);
  const planned = (operation: FolioDocumentOperation, supplied: string[] = []) => ({
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

const open = (bytes: ArrayBuffer) => FolioDocxReviewer.fromBuffer(bytes, { author: "Tester" });

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
  const planned = plans.map((plan, index) => buildOperation(spec, plan, blocks, `op${index}`));
  const result = reviewer.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode,
    operations: planned.map(({ operation }) => operation),
  });
  return { planned, reviewer, result };
};

const appliedIds = ({ result }: Run): Set<string> => new Set(result.applied.map(({ id }) => id));

const readTablesAt = async (stage: string, reviewer: FolioDocxReviewer): Promise<TableReading> => {
  try {
    return await readReviewerTables(reviewer);
  } catch (cause) {
    throw new Error(`Table property failed while reading ${stage}`, { cause });
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
const visible = (reading: TableReading) => ({ texts: reading.texts, tables: reading.snapshot });

/**
 * Whether tracked mode may refuse what direct mode applies. Each case is a
 * shape the tracked-change vocabulary cannot record, and nothing else is:
 *
 * - A merge is tracked as `w:vMerge` under `w:tcPr/w:cellMerge`, so only a
 *   single column of single, unmerged cells merges tracked, and only into
 *   blank cells, since the record moves no content. Nor can it record a row
 *   going with the merge, which the direct merge does when it takes a row's
 *   only cell of its own.
 * - A split is the reverse, so only a vertical merge splits tracked.
 * - A column inserted through a horizontal merge widens it, and a column
 *   deleted through one narrows it: a `w:gridSpan` change has no tracked form.
 */
const trackedRefusalIsExpected = (planned: PlannedOperation, plan: OperationPlan): boolean => {
  const { anchor, other } = planned;
  const cell = anchor.cell;
  const siblings = anchor.siblings;
  switch (plan.kind) {
    case "mergeTableCells": {
      const top = Math.min(cell.row, other.row);
      const bottom = Math.max(cell.row + cell.rowSpan, other.row + other.rowSpan);
      const left = Math.min(cell.column, other.column);
      const right = Math.max(cell.column + cell.columnSpan, other.column + other.columnSpan);
      const inside = siblings.filter(
        (candidate) =>
          candidate.row < bottom &&
          candidate.row + candidate.rowSpan > top &&
          candidate.column < right &&
          candidate.column + candidate.columnSpan > left,
      );
      const rowsLeftEmpty = Array.from(
        { length: bottom - top - 1 },
        (_, index) => top + 1 + index,
      ).some(
        (row) =>
          !siblings.some((candidate) => candidate.row === row && !inside.includes(candidate)),
      );
      return (
        right - left !== 1 ||
        rowsLeftEmpty ||
        inside.some(
          (candidate) =>
            candidate.rowSpan !== 1 ||
            candidate.columnSpan !== 1 ||
            candidate.nested !== undefined ||
            (candidate.row !== top && candidate.text !== ""),
        )
      );
    }
    case "splitTableCell":
      return cell.columnSpan > 1;
    case "insertTableColumn": {
      const boundary = plan.position === "before" ? cell.column : cell.column + cell.columnSpan;
      return siblings.some(
        (candidate) =>
          candidate.column < boundary && candidate.column + candidate.columnSpan > boundary,
      );
    }
    case "deleteTableColumn":
      return siblings.some(
        (candidate) =>
          candidate.columnSpan > 1 &&
          candidate.column <= cell.column &&
          candidate.column + candidate.columnSpan > cell.column,
      );
    default:
      return false;
  }
};

const checkCase = async (spec: TableSpec, plans: readonly OperationPlan[]): Promise<void> => {
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
    directApplied.has(planned.operation.id) === trackedApplied.has(planned.operation.id)
      ? []
      : [{ planned, plan: plans[index]! }],
  );
  if (disagreeing.length > 0) {
    // (b) Both modes agree on whether each operation is allowed, bar the
    // refusals the tracked vocabulary forces.
    for (const { planned, plan } of disagreeing) {
      const id = planned.operation.id;
      expect({
        plans,
        direct: direct.result.skipped,
        tracked: tracked.result.skipped,
        expected:
          directApplied.has(id) &&
          tracked.result.skipped.some(
            (skip) => skip.id === id && skip.reason === "unsupportedBlock",
          ) &&
          trackedRefusalIsExpected(planned, plan),
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
  const sharedResult = shared.applyDocumentOperations({
    version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
    mode: "direct",
    operations: sharedOperations,
  });
  expect(new Set(sharedResult.applied.map(({ id }) => id))).toEqual(trackedApplied);
  const sharedReading = await readTablesAt("direct shared", shared);
  expect(tableReadingProblems(sharedReading)).toEqual([]);
  expect(visible(accepted)).toEqual(visible(sharedReading));
};

describe("table operations on merged tables", () => {
  test(
    "keep every value, agree across modes and readers, and reject cleanly",
    async () => {
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({ numRuns: 150, seed: -1930932135 }),
      );
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({ numRuns: 150, seed: 58022172 }),
      );
      // A refused column insertion and a no-op deletion were reported in opposite orders.
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({
          numRuns: 150,
          seed: 58022172,
          path: "282:1:0:1:0:0:0:0:0:1:1:5:5:5:6:6:6",
        }),
      );
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({ numRuns: 150, seed: -324071034 }),
      );
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({ numRuns: 150, seed: -1330713042 }),
      );
      await fc.assert(
        fc.asyncProperty(documentArbitrary, batchArbitrary, checkCase),
        propertyConfig({ numRuns: 150 }),
      );
    },
    propertyTestTimeout(120_000),
  );
});
