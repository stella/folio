/**
 * Partial structural review was missing from the generated sequence space:
 * it exposes unstable repeated resolutions and save/reopen identity drift.
 */
import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { assertProperty, propertyTestTimeout } from "../../../../test/property-testing";
import {
  assertReviewRoundTrip,
  canonicalReviewModel,
} from "../../../../test/consumer-scenarios/support/review-comparison";
import { buildTableDocx, type TableSpec } from "../__tests__/tableOperationDocument";
import { FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION } from "../document-operations";
import { FolioDocxReviewer } from "./headless";

const operationArbitrary = fc.oneof(
  fc.record({
    rows: fc.integer({ min: 2, max: 4 }),
    columns: fc.integer({ min: 1, max: 3 }),
    operation: fc.constant("insertTableColumn" as const),
  }),
  fc.record({
    rows: fc.integer({ min: 2, max: 4 }),
    columns: fc.integer({ min: 2, max: 3 }),
    operation: fc.constant("deleteTableColumn" as const),
  }),
);

type ColumnOperation = fc.ArbitraryValue<typeof operationArbitrary>["operation"];

const partialReviewArbitrary = operationArbitrary.chain(({ rows, columns, operation }) =>
  fc
    .shuffledSubarray(
      Array.from({ length: rows }, (_, index) => index),
      {
        minLength: 1,
        maxLength: rows - 1,
      },
    )
    .chain((selectedRows) =>
      fc.constantFrom("accept", "reject").map((resolution) => ({
        rows,
        columns,
        operation,
        selectedRows,
        resolution,
      })),
    ),
);

const tableSpec = (rows: number, columns: number): TableSpec => ({
  rows,
  columns,
  cells: Array.from({ length: rows * columns }, (_, index) => {
    const row = Math.floor(index / columns);
    const column = index % columns;
    return { row, column, rowSpan: 1, columnSpan: 1, text: `r${row}c${column}` };
  }),
});

type ReviewerOptions = { rows: number; columns: number; operation: ColumnOperation };

const reviewerFor = async ({ rows, columns, operation }: ReviewerOptions) => {
  const reviewer = await FolioDocxReviewer.fromBuffer(
    await buildTableDocx(tableSpec(rows, columns)),
  );
  const anchor = reviewer.getContent().find(({ text }) => text === "r0c0");
  if (!anchor) throw new Error("rectangular table fixture must expose its first cell");
  const result = reviewer.applyDocumentOperations(
    {
      version: FOLIO_DOCUMENT_OPERATION_CONTRACT_VERSION,
      mode: "tracked-changes",
      operations: [
        {
          id: "column",
          type: operation,
          blockId: anchor.id,
          ...(operation === "insertTableColumn" && {
            position: "after" as const,
            cellTexts: Array.from({ length: rows }, (_, row) => `new${row}`),
          }),
        },
      ],
    },
    { undefinedReferences: "refuse" },
  );
  expect(result.skipped).toEqual([]);
  expect(result.applied).toHaveLength(1);
  return reviewer;
};

describe("partial tracked table-column review", () => {
  test(
    "generated rectangular columns keep repeated resolutions and save/reopen results stable",
    async () => {
      await assertProperty(
        fc.asyncProperty(
          partialReviewArbitrary,
          async ({ rows, columns, operation, selectedRows, resolution }) => {
            const reviewer = await reviewerFor({ rows, columns, operation });
            const changeType = operation === "insertTableColumn" ? "cellInserted" : "cellDeleted";
            const changes = reviewer.getChanges().filter(({ type }) => type === changeType);
            expect(changes).toHaveLength(rows);

            for (const row of selectedRows) {
              const change = changes[row];
              if (!change) throw new Error(`missing structural cell change for row ${row}`);
              const resolved =
                resolution === "accept"
                  ? reviewer.acceptChange(change)
                  : reviewer.rejectChange(change);
              expect(resolved).toBe(true);
            }
            expect(reviewer.getChanges().filter(({ type }) => type === changeType)).toHaveLength(
              rows - selectedRows.length,
            );

            for (const decision of ["accept", "reject"] as const) {
              const first = canonicalReviewModel(reviewer.toResolvedDocument(decision));
              const second = canonicalReviewModel(reviewer.toResolvedDocument(decision));
              expect(second).toEqual(first);
            }

            const reopened = await FolioDocxReviewer.fromBuffer(await reviewer.toBuffer());
            await assertReviewRoundTrip({
              live: reviewer,
              reopened,
              context: `${operation}, ${rows}x${columns}, partial ${resolution}`,
            });
          },
        ),
        { id: "partial tracked table-column review round trip", numRuns: 30 },
      );
    },
    propertyTestTimeout(30_000),
  );
});
