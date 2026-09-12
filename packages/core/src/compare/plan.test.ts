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

import { createFolioAIEditSnapshot } from "../ai-edits/snapshot";
import type { FolioAIBlock } from "../ai-edits/types";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentSnapshot,
} from "../internal/compare/resolved-docx-story-snapshot";
import { toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Document, Paragraph, Table, TableCell } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createContentComparisonWorkSession } from "./content";
import { planStoryCompare } from "./plan";

const MAIN_STORY = { type: "main" } as const;

const block = (id: string, text: string): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
  idStability: "positional",
});

/** One single-cell row of a table, as the snapshot would project it. */
const cell = (id: string, text: string, rowIndex: number, tableIndex = 0): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
  idStability: "positional",
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
  idStability: "positional",
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

const paragraphOf = ({
  id,
  kind,
  text,
  headingLevel,
  styleId,
  directAlignment,
  directSpacing,
  listLevel,
}: FolioAIBlock): Paragraph => ({
  type: "paragraph",
  paraId: id,
  content:
    text.length === 0
      ? []
      : [{ type: "run", content: [{ type: "text", text }] }],
  ...((kind === "heading" ||
    styleId !== undefined ||
    directAlignment !== undefined ||
    directSpacing !== undefined ||
    listLevel !== undefined) && {
    formatting: {
      ...(kind === "heading" && {
        outlineLevel: Math.max(0, Math.min(8, (headingLevel ?? 1) - 1)),
      }),
      ...(styleId !== undefined && { styleId }),
      ...(directAlignment !== undefined && { alignment: directAlignment }),
      ...(directSpacing !== undefined && directSpacing),
      ...(listLevel !== undefined && { numPr: { numId: 1, ilvl: listLevel } }),
    },
  }),
});

const tableOf = (blocks: readonly FolioAIBlock[]): Table => {
  const blocksByRow = new Map<number, FolioAIBlock[]>();
  for (const block of blocks) {
    const table = block.table;
    if (!table) throw new Error("table fixture block has no table location");
    const row = blocksByRow.get(table.rowIndex) ?? [];
    row.push(block);
    blocksByRow.set(table.rowIndex, row);
  }
  return {
    type: "table",
    rows: [...blocksByRow.entries()]
      .toSorted(([left], [right]) => left - right)
      .map(([, rowBlocks]) => {
        const blocksByCell = new Map<number, FolioAIBlock[]>();
        for (const block of rowBlocks) {
          const table = block.table;
          if (!table) throw new Error("table fixture block has no table location");
          const cell = blocksByCell.get(table.cellIndex) ?? [];
          cell.push(block);
          blocksByCell.set(table.cellIndex, cell);
        }
        const cells: TableCell[] = [...blocksByCell.entries()]
          .toSorted(([left], [right]) => left - right)
          .map(([, cellBlocks]) => {
            const ordered = cellBlocks.toSorted(
              (left, right) =>
                (left.table?.paragraphIndex ?? 0) - (right.table?.paragraphIndex ?? 0),
            );
            const span = ordered.at(0)?.table?.columnSpan ?? 1;
            return {
              type: "tableCell",
              ...(span > 1 && { formatting: { gridSpan: span } }),
              content: ordered.map(paragraphOf),
            };
          });
        return { type: "tableRow", cells };
      }),
  };
};

const documentOf = (blocks: readonly FolioAIBlock[]): Document => {
  const template = createEmptyDocument();
  const content: BlockContent[] = [];
  for (let index = 0; index < blocks.length; ) {
    const current = blocks[index];
    if (!current) break;
    if (!current.table) {
      content.push(paragraphOf(current));
      index += 1;
      continue;
    }
    const tableIndex = current.table.tableIndex;
    const tableBlocks: FolioAIBlock[] = [];
    while (blocks[index]?.table?.tableIndex === tableIndex) {
      tableBlocks.push(blocks[index] as FolioAIBlock);
      index += 1;
    }
    content.push(tableOf(tableBlocks));
  }
  return {
    ...template,
    package: {
      ...template.package,
      document: { ...template.package.document, content },
    },
  };
};

const resolvedDocumentSnapshotOf = (
  projectedDocument: Document,
  story = MAIN_STORY,
) => {
  const source = toProseDoc(projectedDocument);
  const snapshot = createResolvedDocxStorySnapshot({
    document:
      story.type === "main"
        ? projectedDocument
        : {
            ...projectedDocument,
            package: {
              ...projectedDocument.package,
              headers: new Map([
                [story.relationshipId, { content: projectedDocument.package.document.content }],
              ]),
            },
          },
    story,
    operationSnapshot: createFolioAIEditSnapshot(source),
  });
  if (!snapshot) throw new Error("fixture story projection missing");
  return snapshot;
};

const resolvedSnapshotOf = (
  blocks: readonly FolioAIBlock[],
  story = MAIN_STORY,
) => resolvedDocumentSnapshotOf(documentOf(blocks), story);

const planOf = (base: readonly FolioAIBlock[], target: readonly FolioAIBlock[]) => {
  const baseSnapshot = resolvedSnapshotOf(base);
  const targetSnapshot = resolvedSnapshotOf(target);
  const captured = createContentComparisonWorkSession().captureComparison({
    base: resolvedDocxContentSnapshot(baseSnapshot),
    revised: resolvedDocxContentSnapshot(targetSnapshot),
  });
  if (captured.isErr()) throw captured.error;
  const comparison = captured.value.compare();
  if (comparison.isErr()) throw comparison.error;
  const plan = planStoryCompare({
    story: MAIN_STORY,
    baseSnapshot,
    targetSnapshot,
    comparison: comparison.value,
    maxOperations: 1000,
  });
  if (plan.isErr()) throw plan.error;
  return {
    changes: plan.value.changes,
    unsupported: plan.value.unsupported,
    instructions: plan.value.program.consume(baseSnapshot),
  };
};

const RELOCATED =
  "The Supplier shall deliver the Goods to the named place within thirty days of the order.";

/** The same clause with one word changed: still the clause that moved. */
const RELOCATED_EDITED =
  "The Supplier shall deliver the Products to the named place within thirty days of the order.";

/** A clause that shares only its opening: a different obligation entirely. */
const UNRELATED =
  "The Supplier shall not be liable for any indirect loss however it arises in contract.";

test("unsupported live paragraph properties reach the typed lowering refusal", () => {
  const baseDocument = documentOf([block("clause", "Stable clause")]);
  const targetDocument = documentOf([block("clause", "Stable clause")]);
  const baseParagraph = baseDocument.package.document.content.at(0);
  const targetParagraph = targetDocument.package.document.content.at(0);
  if (baseParagraph?.type !== "paragraph" || targetParagraph?.type !== "paragraph") {
    throw new Error("paragraph fixture missing");
  }
  baseParagraph.formatting = { suppressLineNumbers: false };
  targetParagraph.formatting = { suppressLineNumbers: true };
  const baseSnapshot = resolvedDocumentSnapshotOf(baseDocument);
  const targetSnapshot = resolvedDocumentSnapshotOf(targetDocument);
  const captured = createContentComparisonWorkSession().captureComparison({
    base: resolvedDocxContentSnapshot(baseSnapshot),
    revised: resolvedDocxContentSnapshot(targetSnapshot),
  });
  if (captured.isErr()) throw captured.error;
  const comparison = captured.value.compare();
  if (comparison.isErr()) throw comparison.error;
  const planned = planStoryCompare({
    story: MAIN_STORY,
    baseSnapshot,
    targetSnapshot,
    comparison: comparison.value,
    maxOperations: 1000,
  });
  if (planned.isErr()) throw planned.error;

  expect(planned.value.unsupported).toContainEqual({
    reason: "block-semantics",
    story: MAIN_STORY,
    eventType: "modified",
    field: "blockProperties",
    baseBlockId: "clause",
    targetBlockId: "clause",
  });
});

test("a lossy live authored-run projection cannot lower through fallback formatting", () => {
  const baseSnapshot = resolvedSnapshotOf([block("clause", "Alpha")]);
  const targetDocument = documentOf([block("clause", "Alpha")]);
  const targetOperationSnapshot = createFolioAIEditSnapshot(toProseDoc(targetDocument));
  Reflect.set(targetOperationSnapshot.blocks.at(0)!, "text", "Bravo");
  const targetSnapshot = createResolvedDocxStorySnapshot({
    document: targetDocument,
    story: MAIN_STORY,
    operationSnapshot: targetOperationSnapshot,
  });
  if (!targetSnapshot) throw new Error("target story projection missing");
  const captured = createContentComparisonWorkSession().captureComparison({
    base: resolvedDocxContentSnapshot(baseSnapshot),
    revised: resolvedDocxContentSnapshot(targetSnapshot),
  });
  if (captured.isErr()) throw captured.error;
  const comparison = captured.value.compare();
  if (comparison.isErr()) throw comparison.error;
  const planned = planStoryCompare({
    story: MAIN_STORY,
    baseSnapshot,
    targetSnapshot,
    comparison: comparison.value,
    maxOperations: 1000,
  });
  if (planned.isErr()) throw planned.error;

  expect(planned.value.unsupported).toContainEqual({
    reason: "block-semantics",
    story: MAIN_STORY,
    eventType: "modified",
    field: "runs.authoredProjection",
    targetBlockId: "clause",
  });
  expect(planned.value.program.size).toBe(0);
});

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

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

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

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

    expect(operations).not.toContainEqual(
      expect.objectContaining({ type: "deleteBlock", blockId: "base-carrier" }),
    );
    expect(changes.map(({ kind }) => kind)).toEqual(["delete", "table-delete"]);
    expect(operations).toEqual([
      { id: "compare-1", type: "deleteBlock", blockId: "between" },
      { id: "compare-2", type: "deleteTable", blockId: "removed" },
    ]);
  });

  test("places a replacement terminal table after its deletable base carrier", () => {
    const base = [
      gridCell("source", "Source terminal table", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      block("base-carrier", ""),
    ];
    const target = [
      gridCell("target", "Target terminal table", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        columnSpan: 2,
      }),
    ];

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

    expect(changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert", "delete"]);
    expect(operations).toEqual([
      { id: "compare-1", type: "deleteTable", blockId: "source" },
      {
        id: "compare-2",
        type: "insertTable",
        blockId: "base-carrier",
        position: "after",
        rows: [["Target terminal table"]],
      },
      { id: "compare-3", type: "deleteBlock", blockId: "base-carrier" },
    ]);
  });

  test("still compares paragraph properties on the reserved carrier", () => {
    const baseCarrier = block("base-carrier", "");
    const targetCarrier = { ...block("target-carrier", ""), styleId: "CustomStyle" };

    const plan = planOf([baseCarrier], [targetCarrier]);
    const { changes } = plan;
    const operations = plan.instructions;

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
      const plan = planOf([base], [target]);
      const { changes } = plan;
      const operations = plan.instructions;
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

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

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

    const operations = planOf(base, target).instructions;

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
    const story = { type: "header", relationshipId: "rId1" } as const;
    const baseSnapshot = resolvedSnapshotOf(base, story);
    const targetSnapshot = resolvedSnapshotOf(target, story);
    const comparison = compareContent({
      base: resolvedDocxContentSnapshot(baseSnapshot),
      revised: resolvedDocxContentSnapshot(targetSnapshot),
    });
    if (comparison.isErr()) throw comparison.error;
    const plan = planStoryCompare({
      story,
      baseSnapshot,
      targetSnapshot,
      comparison: comparison.value,
      maxOperations: 1000,
    });
    if (plan.isErr()) throw plan.error;

    expect(plan.value.changes).not.toContainEqual(
      expect.objectContaining({ kind: "delete", baseBlockId: "base-carrier" }),
    );
    expect(plan.value.program.consume(baseSnapshot)).not.toContainEqual(
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

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

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

    const plan = planOf(base, target);
    const { changes } = plan;
    const operations = plan.instructions;

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

    expect(planOf(base, target).instructions).toEqual([
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

    expect(planOf(base, target).instructions).toEqual([
      { id: "compare-1", type: "deleteTableColumn", blockId: "x" },
      { id: "compare-2", type: "deleteTableColumn", blockId: "y" },
    ]);
  });
});

describe("move detection", () => {
  test("a relocated paragraph is a move even when a word changed on the way", () => {
    const plan = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", RELOCATED_EDITED)],
    );
    const { changes } = plan;
    const operations = plan.instructions;

    expect(changes.map(({ kind }) => kind)).toEqual(["move"]);
    // One closed instruction owns both source and destination, so execution
    // cannot lower them as unrelated deletion and insertion revisions.
    expect(operations.map(({ type }) => type)).toEqual(["moveParagraph"]);
  });

  test("a paragraph that only shares its opening is a deletion and an insertion", () => {
    const plan = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", UNRELATED)],
    );
    const { changes } = plan;
    const operations = plan.instructions;

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
