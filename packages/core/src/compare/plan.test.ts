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
const cell = (id: string, text: string, rowIndex: number, tableIndex = 0): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
  table: {
    outerTableIndex: tableIndex,
    tableIndex,
    rowIndex,
    cellIndex: 0,
    gridColumnIndex: 0,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex: 0,
  },
});

type GridCellGeometry = {
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex: number;
  columnSpan?: number;
  rowSpan?: number;
  paragraphIndex?: number;
};

const gridCell = (
  id: string,
  text: string,
  {
    rowIndex,
    cellIndex,
    gridColumnIndex,
    columnSpan = 1,
    rowSpan = 1,
    paragraphIndex = 0,
  }: GridCellGeometry,
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
    paragraphIndex,
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
        to: index * 2 + 2,
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
    targetSnapshot: snapshotOf(target),
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

  test("a span change that cannot align by column replaces the whole table", () => {
    const base = [
      block("before", "Before the table."),
      gridCell("a", "Shared", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b", "Shared", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      block("after", "After the table."),
    ];
    const target = [
      block("before-target", "Before the table."),
      gridCell("merged", "Shared", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        columnSpan: 2,
      }),
      block("after-target", "After the table."),
    ];

    const { changes, operations } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert"]);
    expect(operations).toEqual([
      { id: "compare-1", type: "deleteTable", blockId: "a" },
      {
        id: "compare-2",
        type: "insertTable",
        blockId: "after",
        position: "before",
        rows: [["Shared"]],
      },
    ]);
  });

  test("a compatible row insertion remains one row insertion", () => {
    const base = [
      block("before", "Before the table."),
      cell("kept", "Shared obligation", 0),
      block("after", "After the table."),
    ];
    const target = [
      block("before-target", "Before the table."),
      cell("kept-target", "Shared obligation", 0),
      cell("added", "Additional obligation", 1),
      block("after-target", "After the table."),
    ];

    const { changes } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-row-insert"]);
  });

  test("text-only rewrites in the same table shape stay cell-level", () => {
    const base = [
      block("before", "Before the table."),
      cell("old", "Original obligation", 0),
      block("after", "After the table."),
    ];
    const target = [
      block("before-target", "Before the table."),
      cell("new", "Entirely different language", 0),
      block("after-target", "After the table."),
    ];

    const { changes } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["replace"]);
  });

  test("paragraph-count changes inside one cell do not replace the table", () => {
    const base = [
      block("before", "Before the table."),
      gridCell("kept", "Shared opening", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      gridCell("removed-a", "Removed middle", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 1,
      }),
      gridCell("removed-b", "Removed ending", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 2,
      }),
      block("after", "After the table."),
    ];
    const target = [
      block("before-target", "Before the table."),
      gridCell("kept-target", "Shared opening revised", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      block("after-target", "After the table."),
    ];

    const { changes } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["replace", "delete", "delete"]);
  });
});

describe("document-terminal paragraph carrier", () => {
  test("pairs required blank carriers across preceding table deletion", () => {
    const base = [
      cell("kept", "Kept row", 0),
      block("between", ""),
      cell("removed", "Removed row", 0, 1),
      block("base-carrier", ""),
    ];
    const target = [cell("kept-target", "Kept row", 0), block("target-carrier", "")];

    const { changes, operations } = planOf(base, target);

    expect(operations).not.toContainEqual(
      expect.objectContaining({ type: "deleteBlock", blockId: "base-carrier" }),
    );
    expect(changes.map(({ kind }) => kind)).toEqual(["delete", "table-delete"]);
    expect(operations).toEqual([
      { id: "compare-1", type: "deleteBlock", blockId: "between" },
      { id: "compare-2", type: "deleteTable", blockId: "removed" },
    ]);
  });

  test("still compares paragraph properties on the reserved carrier", () => {
    const baseCarrier = block("base-carrier", "");
    const targetCarrier = { ...block("target-carrier", ""), styleId: "CustomStyle" };

    const { changes, operations } = planOf([baseCarrier], [targetCarrier]);

    expect(changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
    expect(operations).toEqual([
      {
        id: "compare-1",
        type: "setBlockParagraphProperties",
        blockId: "base-carrier",
        properties: { styleId: "CustomStyle" },
      },
    ]);
  });

  test("reserves empty heading and list carriers while comparing their properties", () => {
    const cases = [
      {
        base: { ...block("base-heading", ""), kind: "heading" as const, styleId: "Heading1" },
        target: { ...block("target-heading", ""), kind: "heading" as const, styleId: "Heading2" },
        properties: { styleId: "Heading2" },
      },
      {
        base: { ...block("base-list", ""), kind: "listItem" as const, listLevel: 0 },
        target: { ...block("target-list", ""), kind: "listItem" as const, listLevel: 1 },
        properties: { listLevel: 1 },
      },
    ];

    for (const { base, target, properties } of cases) {
      const { changes, operations } = planOf([base], [target]);
      expect(changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
      expect(operations).toEqual([
        {
          id: "compare-1",
          type: "setBlockParagraphProperties",
          blockId: base.id,
          properties,
        },
      ]);
    }
  });

  test("a carrier the merge chain reaches is removed rather than reserved", () => {
    // The reservation is a repair for a carrier nothing can reach, not a
    // preference. Where a paragraph of the same container survives in front of
    // it, the removal resolves down the chain and the reader is told the last
    // paragraph went — not that the one before it went and the last was
    // rewritten into its words, which is what pairing the two ends would say.
    const base = [
      block("alpha", RELOCATED),
      block("bravo", UNRELATED),
      block("charlie", RELOCATED_EDITED),
    ];
    const target = [block("alpha-target", RELOCATED), block("bravo-target", UNRELATED)];

    const { changes, operations } = planOf(base, target);

    expect(changes).toContainEqual(
      expect.objectContaining({ kind: "delete", baseBlockId: "charlie" }),
    );
    expect(changes).not.toContainEqual(expect.objectContaining({ baseBlockId: "bravo" }));
    expect(operations).toContainEqual(
      expect.objectContaining({ type: "mergeBlockWithNext", blockId: "bravo" }),
    );
  });

  test("insertions placed in a cell's removed run land in the cell's carrier", () => {
    // The target's first cell grew two paragraphs and its second cell is gone,
    // so the additions anchor on the paragraph the removed cell ends with —
    // the one whose mark cannot say it went. The last of them is written INTO
    // that carrier and the rest go in front of it, which is what makes the
    // mark count work out without deleting a mark the format keeps.
    const base = [
      gridCell("a0", "Alpha clause states the agreed position.", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      gridCell("b0", "Payment falls due within thirty days of invoice.", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
    ];
    const target = [
      gridCell("ta", "Alpha clause states the agreed position.", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      gridCell("tb", "Notices travel to the address named above.", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      gridCell("tc", "Governing law is that of the named place.", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
    ];

    const { operations } = planOf(base, target);

    expect(operations).toContainEqual(
      expect.objectContaining({
        type: "insertBeforeBlock",
        blockId: "b0",
        text: "Notices travel to the address named above.",
      }),
    );
    expect(operations).toContainEqual(
      expect.objectContaining({
        type: "replaceBlock",
        blockId: "b0",
        text: "Governing law is that of the named place.",
      }),
    );
    expect(operations).not.toContainEqual(
      expect.objectContaining({ type: "deleteBlock", blockId: "b0" }),
    );
  });

  test("reserves a stranded final paragraph in a header story too", () => {
    // A header ends with a paragraph for the same reason a body does, and its
    // mark is as unable to say it went. Nothing of the header's own container
    // survives in front of this carrier — a table sits there — so no merge
    // chain reaches it and the words the story ends with have to land on it.
    const base = [
      cell("kept", "Kept row", 0),
      block("between", ""),
      cell("removed", "Removed row", 0, 1),
      block("base-carrier", ""),
    ];
    const target = [cell("kept-target", "Kept row", 0), block("target-carrier", "")];
    const plan = planStoryCompare({
      story: { type: "header", relationshipId: "rId1" },
      baseSnapshot: snapshotOf(base),
      targetSnapshot: snapshotOf(target),
      maxOperations: 1000,
    });
    if (plan === null) {
      throw new Error("The plan exceeded its operation budget.");
    }

    expect(plan.changes).not.toContainEqual(
      expect.objectContaining({ kind: "delete", baseBlockId: "base-carrier" }),
    );
    expect(plan.operations).not.toContainEqual(
      expect.objectContaining({ type: "deleteBlock", blockId: "base-carrier" }),
    );
  });
});

describe("table column pairing", () => {
  test("an inserted middle column is one structural edit and leaves its neighbours paired", () => {
    const base = [
      gridCell("a0", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b0", "Amount", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("a1", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b1", "100", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
    ];
    const target = [
      gridCell("a0-target", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x0", "Currency", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("b0-target", "Amount", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      gridCell("a1-target", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x1", "EUR", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("b1-target", "100", { rowIndex: 1, cellIndex: 2, gridColumnIndex: 2 }),
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
      gridCell("merged", "Account details", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        columnSpan: 2,
      }),
      gridCell("tail", "Status", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 2 }),
      gridCell("a1", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("removed", "EUR", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("tail1", "Open", { rowIndex: 1, cellIndex: 2, gridColumnIndex: 2 }),
    ];
    const target = [
      gridCell("a0-target", "Account details", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      gridCell("tail-target", "Status", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      gridCell("a1-target", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("tail1-target", "Open", {
        rowIndex: 1,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
    ];

    const { changes } = planOf(base, target);

    // The candidate cuts the merged first-row cell, so it is deliberately not
    // represented as a column deletion.
    expect(changes.some(({ kind }) => kind === "table-column-delete")).toBe(false);
  });

  test("a deleted middle column uses a physical-cell anchor for its grid coordinate", () => {
    const base = [
      gridCell("a0", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x0", "Currency", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("b0", "Status", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      gridCell("a1", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x1", "EUR", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("b1", "Open", { rowIndex: 1, cellIndex: 2, gridColumnIndex: 2 }),
    ];
    const target = [
      gridCell("a0-target", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b0-target", "Status", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("a1-target", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b1-target", "Open", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
    ];

    const { changes, operations } = planOf(base, target);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-column-delete"]);
    expect(operations).toEqual([{ id: "compare-1", type: "deleteTableColumn", blockId: "x0" }]);
  });

  test("repeated empty columns stay ambiguous", () => {
    const base = [
      gridCell("a", "", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b", "", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
    ];
    const target = [
      gridCell("a-target", "", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x", "", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("b-target", "", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
    ];

    expect(planOf(base, target).changes.some(({ kind }) => kind === "table-column-insert")).toBe(
      false,
    );
  });

  test("several inserted columns retain target order at one boundary", () => {
    const base = [
      gridCell("a", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b", "Status", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
    ];
    const target = [
      gridCell("a-target", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x", "Currency", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("y", "Region", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      gridCell("b-target", "Status", { rowIndex: 0, cellIndex: 3, gridColumnIndex: 3 }),
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
      gridCell("a", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("x", "Currency", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      gridCell("y", "Region", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      gridCell("b", "Status", { rowIndex: 0, cellIndex: 3, gridColumnIndex: 3 }),
    ];
    const target = [
      gridCell("a-target", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      gridCell("b-target", "Status", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
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
