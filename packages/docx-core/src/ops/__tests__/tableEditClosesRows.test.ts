/**
 * A table edit that leaves a row nothing it starts, only vertical merges
 * from above, removes the row: no reader can hold a row of continuations.
 * Directly the row goes with the edit; a suggested edit records the row's
 * deletion (a merge) or removes it on accept (a column deletion), so
 * accepting equals the direct edit and rejecting restores the original.
 */
import { describe, expect, test } from "bun:test";

import type {
  Document,
  Paragraph,
  Table,
  TableCell,
} from "../../model/document";
import { applyDocumentOp, applyDocumentOps } from "../apply";
import { contractViolation } from "../contract";
import {
  allocateEditorIntentIds,
  compileEditorIntent,
  type TableIntentOperation,
} from "../editorIntent";
import { DOCUMENT_OP_TYPES, OP_STORIES } from "../types";

const id = (value: number) => value.toString(16).toUpperCase().padStart(8, "0");
const paragraph = (value: number, text: string): Paragraph => ({
  type: "paragraph",
  paraId: id(value),
  content:
    text === "" ? [] : [{ type: "run", content: [{ type: "text", text }] }],
});
const cell = (
  value: number,
  text: string,
  vMerge?: "restart" | "continue",
): TableCell => ({
  type: "tableCell",
  content: [paragraph(value, text)],
  ...(vMerge === undefined ? {} : { formatting: { vMerge } }),
});
const documentOf = (table: Table): Document => ({
  package: {
    document: {
      content: [paragraph(1, "before"), table, paragraph(2, "after")],
    },
  },
});
const tableOf = (document: Document): Table => {
  const table = document.package.document.content.at(1);
  if (table?.type !== "table") throw new Error("The table stays in place.");
  return table;
};
const address = { story: OP_STORIES.MAIN, blockId: id(0x100) };
const REVISION = {
  author: "Reviewer",
  date: "2026-11-21T10:00:00Z",
} as const;

type Outcome = {
  direct: Document;
  tracked: Document;
  revisions: readonly number[];
};

/** The operation applied directly and suggested, through the compiler. */
const outcomes = (document: Document, operation: TableIntentOperation) => {
  const intent = { type: "table", operation } as const;
  const allocation = allocateEditorIntentIds(document, intent);
  const run = (mode: Parameters<typeof compileEditorIntent>[1]["mode"]) => {
    const compiled = compileEditorIntent(document, { intent, mode });
    if (compiled.isErr()) throw compiled.error;
    const op = compiled.value.ops.at(0);
    if (op === undefined) throw new Error("The compiler emits one operation.");
    const result = applyDocumentOp(document, op);
    if (result.isErr()) throw result.error;
    expect(contractViolation(result.value.document)).toBeUndefined();
    const undone = applyDocumentOps(
      result.value.document,
      result.value.inverse,
    );
    if (undone.isErr()) throw undone.error;
    expect(undone.value.document).toStrictEqual(document);
    return result.value;
  };
  const direct = run({ type: "editing", newIds: allocation.newIds });
  const tracked = run({
    type: "suggesting",
    revision: { ...REVISION, id: allocation.revisionId },
    newIds: allocation.newIds,
  });
  return {
    direct: direct.document,
    tracked: tracked.document,
    revisions: tracked.revisions ?? [],
  } satisfies Outcome;
};

const resolved = (
  document: Document,
  revisionIds: readonly number[],
  decision: "accept" | "reject",
): Document => {
  const result = applyDocumentOp(document, {
    type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
    story: OP_STORIES.MAIN,
    revisionIds: [...revisionIds],
    decision,
  });
  if (result.isErr()) throw result.error;
  expect(contractViolation(result.value.document)).toBeUndefined();
  return result.value.document;
};

/** Accepting every record gives the direct edit; rejecting, the original. */
const expectLaws = (
  original: Document,
  { direct, tracked, revisions }: Outcome,
) => {
  expect(revisions.length).toBeGreaterThan(0);
  expect(resolved(tracked, revisions, "accept")).toStrictEqual(direct);
  expect(resolved(tracked, revisions, "reject")).toStrictEqual(original);
};

/** One column, `owner` over three rows: merging the first two closes one. */
const ONE_COLUMN = documentOf({
  type: "table",
  columnWidths: [900],
  rows: [
    { type: "tableRow", cells: [cell(0x100, "owner")] },
    { type: "tableRow", cells: [cell(0x101, "")] },
    { type: "tableRow", cells: [cell(0x102, "last")] },
  ],
});

/** Two columns, `A1` merged down over the second row. */
const MERGED_LEFT = documentOf({
  type: "table",
  columnWidths: [900, 900],
  rows: [
    {
      type: "tableRow",
      cells: [cell(0x100, "A1", "restart"), cell(0x110, "B1")],
    },
    {
      type: "tableRow",
      cells: [cell(0x101, "", "continue"), cell(0x111, "B2")],
    },
    { type: "tableRow", cells: [cell(0x102, "A3"), cell(0x112, "B3")] },
  ],
});

describe("a table edit that leaves a row only merges from above", () => {
  test("a vertical merge removes the row directly and records its deletion when suggested", () => {
    const outcome = outcomes(ONE_COLUMN, {
      ...address,
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      top: 0,
      bottom: 2,
      left: 0,
      right: 1,
    });
    const direct = tableOf(outcome.direct);
    expect(
      direct.rows.map((row) => row.cells.map(({ content }) => content)),
    ).toEqual([[[paragraph(0x100, "owner")]], [[paragraph(0x102, "last")]]]);
    const tracked = tableOf(outcome.tracked);
    expect(tracked.rows).toHaveLength(3);
    expect(tracked.rows.at(1)?.structuralChange?.type).toBe("tableRowDeletion");
    expect(tracked.rows.at(1)?.cells.at(0)?.structuralChange?.type).toBe(
      "tableCellMerge",
    );
    expectLaws(ONE_COLUMN, outcome);
  });

  test("each record of a suggested merge resolves alone to a valid document, or refuses as any vertical merge does", () => {
    const { tracked, revisions } = outcomes(ONE_COLUMN, {
      ...address,
      type: DOCUMENT_OP_TYPES.MERGE_CELLS,
      top: 0,
      bottom: 2,
      left: 0,
      right: 1,
    });
    const owner = tableOf(tracked).rows.at(0)?.cells.at(0)?.structuralChange
      ?.info.id;
    const refused: [number, string][] = [];
    for (const revision of revisions)
      for (const decision of ["accept", "reject"] as const) {
        const result = applyDocumentOp(tracked, {
          type: DOCUMENT_OP_TYPES.RESOLVE_REVISION,
          story: OP_STORIES.MAIN,
          revisionIds: [revision],
          decision,
        });
        if (result.isErr()) refused.push([revision, decision]);
        else expect(contractViolation(result.value.document)).toBeUndefined();
      }
    // Rejecting the owner's merge record alone would leave the rows below
    // continuing nothing, for a merge that closes no row too.
    expect(refused).toEqual([[owner ?? -1, "reject"]]);
  });

  test("a column deletion beside a merge removes the row directly and on accept", () => {
    const outcome = outcomes(MERGED_LEFT, {
      ...address,
      blockId: id(0x110),
      type: DOCUMENT_OP_TYPES.DELETE_COLUMN,
      column: 1,
    });
    const direct = tableOf(outcome.direct);
    expect(
      direct.rows.map((row) => row.cells.map(({ content }) => content)),
    ).toEqual([[[paragraph(0x100, "A1")]], [[paragraph(0x102, "A3")]]]);
    expectLaws(MERGED_LEFT, outcome);
  });
});
