/**
 * Several adjacent columns inserted by one operation share the table's one
 * grid change: the edit equals that many single insertions, and resolving its
 * records together gives the direct edit back (accept) or the original
 * (reject). A second tracked insertion beside the pending grid change is
 * refused, since one `w:tblGridChange` cannot record two changes that resolve
 * apart.
 */

import { panic } from "better-result";
import { expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import {
  assertProperty,
  propertyTestTimeout,
} from "../../../../../test/property-testing";
import type { Document, Table, TableRow } from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import { DOCUMENT_OP_REFUSAL_REASONS } from "../refusal";
import {
  DOCUMENT_OP_TYPES,
  type DocumentOp,
  OP_STORIES,
  REVISION_DECISIONS,
  toOpEnvelope,
} from "../types";

setDefaultTimeout(propertyTestTimeout(30_000));

const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const rowOf = (start: number, width: number): TableRow => ({
  type: "tableRow",
  cells: Array.from({ length: width }, (_, column) => ({
    type: "tableCell",
    content: [
      {
        type: "paragraph",
        paraId: id(start + column),
        content: [
          {
            type: "run",
            content: [{ type: "text", text: id(start + column) }],
          },
        ],
      },
    ],
  })),
});
const stamp = (revisionId: number) => ({
  id: revisionId,
  author: "Reviewer",
  date: "2026-10-02T10:00:00Z",
});
const applied = (document: Document, op: DocumentOp) => {
  const before = structuredClone(document);
  const wire = JSON.parse(JSON.stringify(toOpEnvelope(op)));
  expect(wire.op).toStrictEqual(op);
  const result = applyDocumentOp(document, wire.op).unwrap();
  expect(contractViolation(result.document)).toBeUndefined();
  const inverse = JSON.parse(JSON.stringify(result.inverse));
  expect(
    applyDocumentOps(result.document, inverse).unwrap().document,
  ).toStrictEqual(before);
  return result;
};

/** The ids of the new cells: row `row`'s cell of new column `column`. */
const newCellId = (row: number, column: number) =>
  id(0x30000 + row * 16 + column);

test("adjacent columns inserted by one operation resolve together and equal single insertions", async () => {
  await assertProperty(
    fc.property(
      fc.record({
        rows: fc.integer({ min: 1, max: 4 }),
        columns: fc.integer({ min: 1, max: 4 }),
        at: fc.nat(),
        count: fc.integer({ min: 1, max: 3 }),
      }),
      ({ rows, columns, at, count }) => {
        const table: Table = {
          type: "table",
          formatting: { layout: "fixed" },
          columnWidths: Array.from({ length: columns }, () => 900),
          rows: Array.from({ length: rows }, (_, index) =>
            rowOf(0x10000 + index * 16, columns),
          ),
        };
        const baseline: Document = {
          package: {
            document: {
              content: [
                { type: "paragraph", paraId: "00000001", content: [] },
                table,
                { type: "paragraph", paraId: "00000002", content: [] },
              ],
            },
          },
        };
        const column = at % (columns + 1);
        const blockId = id(0x10000);
        const insertion = {
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          story: OP_STORIES.MAIN,
          blockId,
          column,
          width: 550,
          count,
          newBlockIds: Array.from({ length: rows * count }, (_, index) =>
            newCellId(Math.floor(index / count), index % count),
          ),
        } as const satisfies DocumentOp;
        const direct = applied(baseline, insertion).document;
        let single = baseline;
        for (let added = 0; added < count; added++)
          single = applied(single, {
            type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
            story: OP_STORIES.MAIN,
            blockId,
            column: column + added,
            width: 550,
            newBlockIds: Array.from({ length: rows }, (_, row) =>
              newCellId(row, added),
            ),
          }).document;
        expect(direct).toStrictEqual(single);

        const pending = applied(baseline, {
          ...insertion,
          revision: stamp(200),
          newIds: {
            revision: Array.from(
              { length: rows * count + 1 },
              (_, index) => 201 + index,
            ),
          },
        });
        const revisionIds =
          pending.revisions ?? panic("Expected column revisions");
        expect(revisionIds).toHaveLength(rows * count + 2);
        const resolve = (decision: "accept" | "reject") =>
          applied(pending.document, {
            type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
            story: OP_STORIES.MAIN,
            revisionIds,
            decision,
          }).document;
        expect(resolve(REVISION_DECISIONS.ACCEPT)).toStrictEqual(direct);
        expect(resolve(REVISION_DECISIONS.REJECT)).toStrictEqual(baseline);

        const second = applyDocumentOp(pending.document, {
          type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
          story: OP_STORIES.MAIN,
          blockId,
          column: 0,
          width: 550,
          newBlockIds: Array.from({ length: rows }, (_, row) =>
            id(0x40000 + row),
          ),
          revision: stamp(300),
          newIds: {
            revision: Array.from(
              { length: rows + 1 },
              (_, index) => 301 + index,
            ),
          },
        });
        expect(second.isErr() && second.error.reason).toBe(
          DOCUMENT_OP_REFUSAL_REASONS.PENDING_GRID_CHANGE,
        );
      },
    ),
    { numRuns: 40, id: "multi-column insertion resolves as one change" },
  );
});
