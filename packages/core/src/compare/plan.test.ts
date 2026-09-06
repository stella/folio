/**
 * Move detection, at the planning layer.
 *
 * The benchmark corpus relocates paragraphs verbatim, so it proves the exact
 * path and says nothing about the threshold. A relocated paragraph is usually
 * edited on the way, and where the line sits — what still counts as the same
 * paragraph somewhere else, and what is a deletion plus an unrelated
 * insertion — is a judgement the plan makes and this file pins.
 */

import { describe, expect, test } from "bun:test";

import type { FolioAIBlock, FolioAIEditSnapshot } from "../ai-edits/types";
import { planStoryCompare } from "./plan";

const MAIN_STORY = { type: "main" } as const;

const block = (id: string, text: string): FolioAIBlock => ({ id, kind: "paragraph", text });

/** One single-cell row of a table, as the snapshot would project it. */
const cell = (id: string, text: string, rowIndex: number): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
  table: {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex,
    cellIndex: 0,
    gridColumnIndex: 0,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex: 0,
  },
});

const gridCell = (
  id: string,
  text: string,
  rowIndex: number,
  cellIndex: number,
  gridColumnIndex: number,
  columnSpan = 1,
  rowSpan = 1,
): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
  table: {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex,
    cellIndex,
    gridColumnIndex,
    columnSpan,
    rowSpan,
    paragraphIndex: 0,
  },
});

const snapshotOf = (blocks: readonly FolioAIBlock[]): FolioAIEditSnapshot => ({
  blocks: [...blocks],
  anchors: Object.fromEntries(
    blocks.map((entry, index) => [
      entry.id,
      {
        id: entry.id,
        from: index * 2,
        to: index * 2 + 1,
        text: entry.text,
        normalizedText: entry.text,
        textHash: entry.text,
        hashOccurrenceCount: 1,
      },
    ]),
  ),
});

const planOf = (base: readonly FolioAIBlock[], target: readonly FolioAIBlock[]) => {
  const plan = planStoryCompare({
    story: MAIN_STORY,
    baseSnapshot: snapshotOf(base),
    targetBlocks: target,
    maxOperations: 1000,
  });
  if (plan === null) {
    throw new Error("The plan exceeded its operation budget.");
  }
  return plan;
};

const RELOCATED =
  "The Supplier shall deliver the Goods to the named place within thirty days of the order.";

/** The same clause with one word changed: still the clause that moved. */
const RELOCATED_EDITED =
  "The Supplier shall deliver the Products to the named place within thirty days of the order.";

/** A clause that shares only its opening: a different obligation entirely. */
const UNRELATED =
  "The Supplier shall not be liable for any indirect loss however it arises in contract.";

describe("table row pairing", () => {
  test("a deleted row plus edits in the rows below is one deleted row", () => {
    // Pairing rows by position instead reads every row as changed: row 0 is
    // put opposite row 1, row 1 opposite row 2, and the last row of the base
    // is deleted — a table-wide rewrite for an edit that removed one row.
    const { changes } = planOf(
      [
        cell("a", "Ordered quantity is one hundred units.", 0),
        cell("b", "Delivery is due on the first of the month.", 1),
        cell("c", "Payment falls due within thirty days.", 2),
      ],
      [
        cell("b2", "Delivery is due on the first of the quarter.", 0),
        cell("c2", "Payment falls due within sixty days.", 1),
      ],
    );

    expect(changes.map(({ kind }) => kind)).toEqual(["table-row-delete", "replace", "replace"]);
  });

  test("an unchanged table reports nothing", () => {
    const rows = [
      cell("a", "Ordered quantity is one hundred units.", 0),
      cell("b", "Delivery is due.", 1),
    ];
    expect(planOf(rows, rows).changes).toEqual([]);
  });
});

describe("table column pairing", () => {
  test("an inserted middle column is one structural edit and leaves its neighbours paired", () => {
    const base = [
      gridCell("a0", "Account", 0, 0, 0),
      gridCell("b0", "Amount", 0, 1, 1),
      gridCell("a1", "Fees", 1, 0, 0),
      gridCell("b1", "100", 1, 1, 1),
    ];
    const target = [
      gridCell("a0-target", "Account", 0, 0, 0),
      gridCell("x0", "Currency", 0, 1, 1),
      gridCell("b0-target", "Amount", 0, 2, 2),
      gridCell("a1-target", "Fees", 1, 0, 0),
      gridCell("x1", "EUR", 1, 1, 1),
      gridCell("b1-target", "100", 1, 2, 2),
    ];

    const { changes, operations } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-column-insert"]);
    expect(operations).toEqual([
      {
        id: "compare-1",
        type: "insertTableColumn",
        blockId: "b0",
        position: "before",
        cellTexts: ["Currency", "EUR"],
      },
    ]);
  });

  test("a candidate column that cuts a merged cell is not guessed from physical indexes", () => {
    const base = [
      gridCell("merged", "Account details", 0, 0, 0, 2),
      gridCell("tail", "Status", 0, 1, 2),
      gridCell("a1", "Fees", 1, 0, 0),
      gridCell("removed", "EUR", 1, 1, 1),
      gridCell("tail1", "Open", 1, 2, 2),
    ];
    const target = [
      gridCell("a0-target", "Account details", 0, 0, 0),
      gridCell("tail-target", "Status", 0, 1, 1),
      gridCell("a1-target", "Fees", 1, 0, 0),
      gridCell("tail1-target", "Open", 1, 1, 1),
    ];

    const { changes } = planOf(base, target);

    // The candidate cuts the merged first-row cell, so it is deliberately not
    // represented as a column deletion.
    expect(changes.some(({ kind }) => kind === "table-column-delete")).toBe(false);
  });

  test("a deleted middle column uses a physical-cell anchor for its grid coordinate", () => {
    const base = [
      gridCell("a0", "Account", 0, 0, 0),
      gridCell("x0", "Currency", 0, 1, 1),
      gridCell("b0", "Status", 0, 2, 2),
      gridCell("a1", "Fees", 1, 0, 0),
      gridCell("x1", "EUR", 1, 1, 1),
      gridCell("b1", "Open", 1, 2, 2),
    ];
    const target = [
      gridCell("a0-target", "Account", 0, 0, 0),
      gridCell("b0-target", "Status", 0, 1, 1),
      gridCell("a1-target", "Fees", 1, 0, 0),
      gridCell("b1-target", "Open", 1, 1, 1),
    ];

    const { changes, operations } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-column-delete"]);
    expect(operations).toEqual([{ id: "compare-1", type: "deleteTableColumn", blockId: "x0" }]);
  });

  test("repeated empty columns stay ambiguous", () => {
    const base = [gridCell("a", "", 0, 0, 0), gridCell("b", "", 0, 1, 1)];
    const target = [
      gridCell("a-target", "", 0, 0, 0),
      gridCell("x", "", 0, 1, 1),
      gridCell("b-target", "", 0, 2, 2),
    ];

    expect(planOf(base, target).changes.some(({ kind }) => kind === "table-column-insert")).toBe(
      false,
    );
  });

  test("several inserted columns retain target order at one boundary", () => {
    const base = [gridCell("a", "Account", 0, 0, 0), gridCell("b", "Status", 0, 1, 1)];
    const target = [
      gridCell("a-target", "Account", 0, 0, 0),
      gridCell("x", "Currency", 0, 1, 1),
      gridCell("y", "Region", 0, 2, 2),
      gridCell("b-target", "Status", 0, 3, 3),
    ];

    expect(planOf(base, target).operations).toEqual([
      {
        id: "compare-1",
        type: "insertTableColumn",
        blockId: "b",
        position: "before",
        cellTexts: ["Currency"],
      },
      {
        id: "compare-2",
        type: "insertTableColumn",
        blockId: "b",
        position: "before",
        cellTexts: ["Region"],
      },
    ]);
  });

  test("several deleted columns use their own physical-cell anchors", () => {
    const base = [
      gridCell("a", "Account", 0, 0, 0),
      gridCell("x", "Currency", 0, 1, 1),
      gridCell("y", "Region", 0, 2, 2),
      gridCell("b", "Status", 0, 3, 3),
    ];
    const target = [
      gridCell("a-target", "Account", 0, 0, 0),
      gridCell("b-target", "Status", 0, 1, 1),
    ];

    expect(planOf(base, target).operations).toEqual([
      { id: "compare-1", type: "deleteTableColumn", blockId: "x" },
      { id: "compare-2", type: "deleteTableColumn", blockId: "y" },
    ]);
  });
});

describe("move detection", () => {
  test("a relocated paragraph is a move even when a word changed on the way", () => {
    const { changes, operations } = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", RELOCATED_EDITED)],
    );

    expect(changes.map(({ kind }) => kind)).toEqual(["move"]);
    // Both halves of the pair carry the same link, so the applier writes
    // `w:moveFrom` and `w:moveTo` rather than two unrelated revisions.
    const moveIds = operations.flatMap((operation) =>
      "moveId" in operation && operation.moveId !== undefined ? [operation.moveId] : [],
    );
    expect(moveIds).toHaveLength(2);
    expect(new Set(moveIds).size).toBe(1);
  });

  test("a paragraph that only shares its opening is a deletion and an insertion", () => {
    const { changes, operations } = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", UNRELATED)],
    );

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
    expect(operations.some((operation) => "moveId" in operation)).toBe(false);
  });

  test("a short relocated line does not pair, so boilerplate does not read as a move", () => {
    const { changes } = planOf(
      [block("a", "Schedule 1"), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", "Schedule 1")],
    );

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
  });
});
