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

import type { FolioAIBlock } from "../ai-edits/types";
import {
  createResolvedDocxStorySnapshot,
  resolvedDocxContentBlocks,
  resolvedDocxContentSnapshot,
  resolvedDocxSourceOperand,
  resolvedDocxSourceOperandBlock,
  type ResolvedDocxSourceOperand,
  type ResolvedDocxStorySnapshot,
} from "../internal/compare/resolved-docx-story-snapshot";
import { headerFooterToProseDoc, toProseDoc } from "../prosemirror/conversion/toProseDoc";
import type { BlockContent, Document, Paragraph, Table, TableCell } from "../types/document";
import { createEmptyDocument } from "../utils/createDocument";
import { createContentComparisonWorkSession } from "./content";
import { planStoryCompare, type CompareStoryPlan } from "./plan";

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
  const document =
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
        };
  const conversionOptions = {
    ...(projectedDocument.package.styles !== undefined && {
      styles: projectedDocument.package.styles,
    }),
    ...(projectedDocument.package.theme !== undefined && {
      theme: projectedDocument.package.theme,
    }),
  };
  const sourceDocument =
    story.type === "main"
      ? toProseDoc(document, conversionOptions)
      : headerFooterToProseDoc(
          projectedDocument.package.document.content,
          conversionOptions,
        );
  const snapshot = createResolvedDocxStorySnapshot({
    document,
    story,
    sourceDocument,
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
    ...plan.value,
    baseSnapshot,
  };
};

type TestCompareStoryPlan = CompareStoryPlan & {
  readonly baseSnapshot: ResolvedDocxStorySnapshot;
};

/** Content transport only; table-property pairing has its own focused suite. */
const contentInstructionsOf = (plan: TestCompareStoryPlan) =>
  plan.program.consume(plan.baseSnapshot).filter(({ type }) => type !== "matchTableGeometry");

const sourceOperandOf = (
  plan: TestCompareStoryPlan,
  blockId: string,
): ResolvedDocxSourceOperand => {
  const block = resolvedDocxContentBlocks(plan.baseSnapshot).find(
    ({ identity }) => identity.id === blockId,
  );
  if (!block) throw new Error(`fixture source block ${blockId} missing`);
  return resolvedDocxSourceOperand(plan.baseSnapshot, block);
};

const sourceBlockIdOf = (
  plan: TestCompareStoryPlan,
  source: ResolvedDocxSourceOperand,
): string => resolvedDocxSourceOperandBlock(source, plan.baseSnapshot).identity.id;

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
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert"]);
    expect(instructions).toEqual([
      {
        type: "replaceTable",
        source: sourceOperandOf(plan, "a"),
        baseTableIndex: 0,
        targetTableIndex: 0,
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
    const instructions = contentInstructionsOf(plan);

    expect(
      instructions.some(
        (instruction) =>
          instruction.type === "deleteParagraph" &&
          sourceBlockIdOf(plan, instruction.source) === "base-carrier",
      ),
    ).toBe(false);
    expect(changes.map(({ kind }) => kind)).toEqual(["delete", "table-delete"]);
    expect(instructions).toEqual([
      { type: "deleteParagraph", source: sourceOperandOf(plan, "between") },
      {
        type: "deleteTable",
        source: sourceOperandOf(plan, "removed"),
        baseTableIndex: 1,
      },
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
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-delete", "table-insert", "delete"]);
    expect(instructions).toEqual([
      {
        type: "replaceTable",
        source: sourceOperandOf(plan, "source"),
        baseTableIndex: 0,
        targetTableIndex: 0,
      },
    ]);
  });

  test("still compares paragraph properties on the reserved carrier", () => {
    const baseCarrier = block("base-carrier", "");
    const targetCarrier = { ...block("target-carrier", ""), styleId: "CustomStyle" };

    const plan = planOf([baseCarrier], [targetCarrier]);
    const { changes } = plan;
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
    expect(instructions).toEqual([
      {
        type: "setParagraphProperties",
        source: sourceOperandOf(plan, "base-carrier"),
        targetProperties: expect.objectContaining({ styleId: "CustomStyle" }),
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
      const instructions = contentInstructionsOf(plan);
      expect(changes.map(({ kind }) => kind)).toEqual(["paragraph-format"]);
      expect(instructions).toEqual([
        {
          type: "setParagraphProperties",
          source: sourceOperandOf(plan, base.id),
          targetProperties: expect.objectContaining(properties),
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
    const instructions = contentInstructionsOf(plan);

    expect(changes).toContainEqual(
      expect.objectContaining({ kind: "delete", baseBlockId: "charlie" }),
    );
    expect(changes).not.toContainEqual(expect.objectContaining({ baseBlockId: "bravo" }));
    expect(instructions).toContainEqual(
      expect.objectContaining({
        type: "deleteTrailingParagraphs",
        chainStart: sourceOperandOf(plan, "bravo"),
        deleted: [sourceOperandOf(plan, "charlie")],
      }),
    );
  });

  test("insertions placed in a cell's removed run land in the cell's carrier", () => {
    // The target's first cell grew two paragraphs and its second cell is gone.
    // The additions stay after the surviving paragraph in their own cell;
    // structural deletion owns the neighboring column independently.
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
        paragraphIndex: 1,
      }),
      gridCell("tc", "Governing law is that of the named place.", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 2,
      }),
    ];

    const plan = planOf(base, target);
    const instructions = contentInstructionsOf(plan);

    expect(instructions).toContainEqual(
      expect.objectContaining({
        type: "insertParagraph",
        boundary: expect.objectContaining({
          type: "afterParagraph",
          paragraph: sourceOperandOf(plan, "a0"),
        }),
        target: expect.objectContaining({ text: "Notices travel to the address named above." }),
      }),
    );
    expect(instructions).toContainEqual(
      expect.objectContaining({
        type: "insertParagraph",
        boundary: expect.objectContaining({
          type: "afterParagraph",
          paragraph: sourceOperandOf(plan, "a0"),
        }),
        target: expect.objectContaining({ text: "Governing law is that of the named place." }),
      }),
    );
    expect(instructions).not.toContainEqual(
      expect.objectContaining({
        type: "deleteParagraph",
        source: sourceOperandOf(plan, "b0"),
      }),
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
    const captured = createContentComparisonWorkSession().captureComparison({
      base: resolvedDocxContentSnapshot(baseSnapshot),
      revised: resolvedDocxContentSnapshot(targetSnapshot),
    });
    if (captured.isErr()) throw captured.error;
    const comparison = captured.value.compare();
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
    const testPlan = { ...plan.value, baseSnapshot };
    expect(
      contentInstructionsOf(testPlan).some(
        (instruction) =>
          instruction.type === "deleteParagraph" &&
          sourceBlockIdOf(testPlan, instruction.source) === "base-carrier",
      ),
    ).toBe(false);
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
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-column-insert"]);
    expect(instructions).toEqual([
      {
        type: "insertTableColumn",
        anchor: { blockId: "b0", position: "before" },
        targetTableIndex: 0,
        targetColumnIndex: 1,
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
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["table-column-delete"]);
    expect(instructions).toEqual([
      {
        type: "deleteTableColumn",
        source: sourceOperandOf(plan, "x0"),
        baseTableIndex: 0,
        baseColumnIndex: 1,
      },
    ]);
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

    expect(contentInstructionsOf(planOf(base, target))).toEqual([
      {
        type: "insertTableColumn",
        anchor: { blockId: "b", position: "before" },
        targetTableIndex: 0,
        targetColumnIndex: 1,
        cellTexts: ["Currency"],
      },
      {
        type: "insertTableColumn",
        anchor: { blockId: "b", position: "before" },
        targetTableIndex: 0,
        targetColumnIndex: 2,
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

    const plan = planOf(base, target);
    expect(contentInstructionsOf(plan)).toEqual([
      {
        type: "deleteTableColumn",
        source: sourceOperandOf(plan, "x"),
        baseTableIndex: 0,
        baseColumnIndex: 1,
      },
      {
        type: "deleteTableColumn",
        source: sourceOperandOf(plan, "y"),
        baseTableIndex: 0,
        baseColumnIndex: 2,
      },
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
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind)).toEqual(["move"]);
    expect(instructions).toContainEqual(
      expect.objectContaining({
        type: "moveParagraph",
        source: sourceOperandOf(plan, "a"),
        target: expect.objectContaining({ text: RELOCATED_EDITED }),
      }),
    );
    expect(instructions.filter(({ type }) => type === "moveParagraph")).toHaveLength(1);
  });

  test("a paragraph that only shares its opening is a deletion and an insertion", () => {
    const plan = planOf(
      [block("a", RELOCATED), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", UNRELATED)],
    );
    const { changes } = plan;
    const instructions = contentInstructionsOf(plan);

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
    expect(instructions.some(({ type }) => type === "moveParagraph")).toBe(false);
  });

  test("a short relocated line does not pair, so boilerplate does not read as a move", () => {
    const { changes } = planOf(
      [block("a", "Schedule 1"), block("b", "An unrelated closing paragraph.")],
      [block("b2", "An unrelated closing paragraph."), block("a2", "Schedule 1")],
    );

    expect(changes.map(({ kind }) => kind).toSorted()).toEqual(["delete", "insert"]);
  });
});
