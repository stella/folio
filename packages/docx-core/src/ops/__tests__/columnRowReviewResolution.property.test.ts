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
const tableIn = (document: Document) => {
  const block = document.package.document.content.at(1);
  return block?.type === "table" ? block : panic("Expected fixture table");
};
const rowId = (row: TableRow) => {
  const paragraph = row.cells.at(0)?.content.at(0);
  return paragraph?.type === "paragraph" && paragraph.paraId !== undefined
    ? paragraph.paraId
    : panic("Expected fixture row anchor");
};
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
  expect(document).toStrictEqual(before);
  return result;
};
type ResolveOptions = {
  document: Document;
  revisionIds: readonly number[];
  decision: "accept" | "reject";
};
const resolved = ({ document, revisionIds, decision }: ResolveOptions) =>
  applied(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds,
    decision,
  }).document;

const decisions = Object.values(REVISION_DECISIONS);
const cases = (["insert", "delete"] as const).flatMap((rowKind) =>
  (["insert", "delete"] as const).flatMap((columnKind) =>
    decisions.flatMap((rowDecision) =>
      decisions.map((columnDecision) => ({
        rowKind,
        columnKind,
        rowDecision,
        columnDecision,
      })),
    ),
  ),
);

test.each(cases)(
  "$rowKind row / $columnKind column resolve $rowDecision/$columnDecision in either order with exact wire inverses",
  async ({ rowKind, columnKind, rowDecision, columnDecision }) => {
    await assertProperty(
      fc.property(
        fc.record({
          rows: fc.integer({ min: 2, max: 4 }),
          columns: fc.integer({ min: 2, max: 4 }),
          at: fc.nat(),
          column: fc.nat(),
        }),
        ({ rows, columns, at, column }) => {
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
          const rowIndex = at % rows;
          const target = rowId(
            table.rows.at(rowIndex) ?? panic("Expected selected row"),
          );
          const rowOp =
            rowKind === "insert"
              ? ({
                  type: DOCUMENT_OP_TYPES.INSERT_ROW,
                  story: OP_STORIES.MAIN,
                  blockId: target,
                  at: rowIndex,
                  row: rowOf(0x20000, columns),
                } as const satisfies DocumentOp)
              : ({
                  type: DOCUMENT_OP_TYPES.DELETE_ROW,
                  story: OP_STORIES.MAIN,
                  blockId: target,
                } as const satisfies DocumentOp);
          const pendingRow = applied(baseline, {
            ...rowOp,
            revision: stamp(100),
            newIds: {
              revision: Array.from({ length: 32 }, (_, index) => 101 + index),
            },
          });
          const rowRevisions =
            pendingRow.revisions ?? panic("Expected row revisions");
          expect(rowRevisions.length).toBeGreaterThan(0);
          const columnIds = new Map(
            tableIn(pendingRow.document).rows.map((row, index) => [
              rowId(row),
              id(0x30000 + index),
            ]),
          );
          const columnOp = (document: Document): DocumentOp => {
            const current = tableIn(document);
            const blockId = rowId(
              current.rows.at(0) ?? panic("Expected surviving row"),
            );
            if (columnKind === "delete")
              return {
                type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
                story: OP_STORIES.MAIN,
                blockId,
                column: column % columns,
              };
            return {
              type: DOCUMENT_OP_TYPES.INSERT_COLUMN,
              story: OP_STORIES.MAIN,
              blockId,
              column: column % columns,
              width: 550,
              newBlockIds: current.rows.map(
                (row) =>
                  columnIds.get(rowId(row)) ??
                  panic("Expected row's new column identity"),
              ),
            };
          };
          const pending = applied(pendingRow.document, {
            ...columnOp(pendingRow.document),
            revision: stamp(200),
            newIds: {
              revision: Array.from({ length: 32 }, (_, index) => 201 + index),
            },
          });
          const columnRevisions =
            pending.revisions ?? panic("Expected column revisions");
          expect(columnRevisions.length).toBeGreaterThan(0);
          const rowFirst = resolved({
            document: resolved({
              document: pending.document,
              revisionIds: rowRevisions,
              decision: rowDecision,
            }),
            revisionIds: columnRevisions,
            decision: columnDecision,
          });
          const columnFirst = resolved({
            document: resolved({
              document: pending.document,
              revisionIds: columnRevisions,
              decision: columnDecision,
            }),
            revisionIds: rowRevisions,
            decision: rowDecision,
          });
          expect(rowFirst).toStrictEqual(columnFirst);
          let expected =
            rowDecision === REVISION_DECISIONS.ACCEPT
              ? applied(baseline, rowOp).document
              : baseline;
          if (columnDecision === REVISION_DECISIONS.ACCEPT)
            expected = applied(expected, columnOp(expected)).document;
          expect(rowFirst).toStrictEqual(expected);
        },
      ),
      {
        numRuns: 12,
        id: `${rowKind} row / ${columnKind} column ${rowDecision}/${columnDecision} resolution order`,
      },
    );
  },
);

test.each(["bookmarks", "control", "revision"] as const)(
  "a deleted column owns its %s through saveable inverses and review decisions",
  (kind) => {
    const table: Table = {
      type: "table",
      formatting: { layout: "fixed" },
      columnWidths: [900, 900],
      rows: [rowOf(0x10000, 2), rowOf(0x10010, 2)],
    };
    const cell = table.rows.at(0)?.cells.at(0) ?? panic("Expected owned cell");
    const first = cell.content.at(0);
    const last = table.rows.at(1)?.cells.at(0)?.content.at(0);
    if (first?.type !== "paragraph" || last?.type !== "paragraph")
      panic("Expected owned paragraphs");
    switch (kind) {
      case "bookmarks":
        first.content.unshift(
          { type: "bookmarkStart", id: 400, name: "inside" },
          { type: "bookmarkEnd", id: 400 },
          { type: "bookmarkStart", id: 401, name: "down" },
        );
        last.content.push({ type: "bookmarkEnd", id: 401 });
        break;
      case "control":
        cell.contentControls = [
          { sdtType: "richText", id: 400, tag: "inside" },
        ];
        break;
      case "revision":
        first.content = [
          { type: "insertion", info: stamp(400), content: first.content },
        ];
        break;
    }
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
    const deletion = {
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      story: OP_STORIES.MAIN,
      blockId: "00010000",
      column: 0,
    } as const satisfies DocumentOp;
    const direct = applied(baseline, deletion);
    expect(tableIn(direct.document).rows.map((row) => row.cells)).toStrictEqual(
      table.rows.map((row) => row.cells.slice(1)),
    );
    const pending = applied(baseline, {
      ...deletion,
      revision: stamp(200),
      newIds: { revision: [201, 202, 203] },
    });
    const revisionIds = pending.revisions ?? panic("Expected column revisions");
    expect(
      resolved({
        document: pending.document,
        revisionIds,
        decision: REVISION_DECISIONS.ACCEPT,
      }),
    ).toStrictEqual(direct.document);
    expect(
      resolved({
        document: pending.document,
        revisionIds,
        decision: REVISION_DECISIONS.REJECT,
      }),
    ).toStrictEqual(baseline);
  },
);

test.each(["start", "end"] as const)(
  "a column cannot remove only a bookmark's %s boundary",
  (boundary) => {
    const table: Table = {
      type: "table",
      columnWidths: [900, 900],
      rows: [rowOf(0x10000, 2), rowOf(0x10010, 2)],
    };
    const firstRemoved = table.rows.at(0)?.cells.at(0)?.content.at(0);
    const lastRemoved = table.rows.at(1)?.cells.at(0)?.content.at(0);
    const survivor = table.rows.at(0)?.cells.at(1)?.content.at(0);
    if (
      firstRemoved?.type !== "paragraph" ||
      lastRemoved?.type !== "paragraph" ||
      survivor?.type !== "paragraph"
    )
      panic("Expected bookmark paragraphs");
    const start = {
      type: "bookmarkStart",
      id: 400,
      name: "_RefCrossColumn",
    } as const;
    const end = { type: "bookmarkEnd", id: 400 } as const;
    if (boundary === "start") {
      firstRemoved.content.unshift(start);
      survivor.content.push(end);
    } else {
      // The surviving start precedes the removed end in document order too.
      survivor.content.unshift(start);
      lastRemoved.content.push(end);
    }
    const document: Document = {
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
    const before = structuredClone(document);
    for (const revision of [undefined, stamp(200)]) {
      const operation = {
        type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
        story: OP_STORIES.MAIN,
        blockId: "00010000",
        column: 0,
        ...(revision === undefined
          ? {}
          : { revision, newIds: { revision: [201, 202, 203] } }),
      } as const satisfies DocumentOp;
      const result = applyDocumentOp(document, operation);
      if (result.isOk())
        panic("A crossing bookmark must refuse column removal");
      expect(result.error.reason).toBe(
        DOCUMENT_OP_REFUSAL_REASONS.DEPENDENT_RECORDS,
      );
      expect(document).toStrictEqual(before);
    }
  },
);
