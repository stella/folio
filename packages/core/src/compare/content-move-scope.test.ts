import { describe, expect, test } from "bun:test";

import type { FolioAIBlock, FolioAIEditSnapshot } from "../ai-edits/types";
import {
  compareContent,
  createContentComparisonWorkSession,
  detectFolioContentMoves,
} from "./content";
import type { FolioContentAlignmentStep } from "./content-alignment";
import type { FolioContentBlock } from "./content-types";
import { planStoryCompare } from "./plan";
import { alignFolioBlocks } from "../version-comparison";

const block = (id: string, text: string): FolioContentBlock => ({
  id,
  kind: "paragraph",
  text,
});

const tableBlock = ({
  id,
  text,
  rowIndex,
  paragraphIndex,
  cellIndex = 0,
  containerId = "logical-cell",
}: {
  id: string;
  text: string;
  rowIndex: number;
  paragraphIndex: number;
  cellIndex?: number;
  containerId?: string;
}): FolioContentBlock => ({
  id,
  kind: "paragraph",
  text,
  containerPath: [{ kind: "cell", id: containerId }],
  table: {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex,
    cellIndex,
    gridColumnIndex: cellIndex,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex,
  },
});

const successfulComparison = (
  base: FolioContentBlock[],
  revised: FolioContentBlock[],
) => {
  const result = compareContent({ base: { blocks: base }, revised: { blocks: revised } });
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

describe("neutral move scope", () => {
  test.each([
    {
      label: "exact text",
      baseText: "alpha beta gamma delta epsilon",
      revisedText: "alpha beta gamma delta epsilon",
    },
    {
      label: "edited text",
      baseText: "alpha beta gamma delta epsilon",
      revisedText: "alpha beta gamma delta zeta",
    },
  ])("a same-slot stable-id replacement with $label is not a relocation", ({
    baseText,
    revisedText,
  }) => {
    const base = block("old-stable-id", baseText);
    const revised = block("new-stable-id", revisedText);

    const comparison = successfulComparison([base], [revised]);

    expect(comparison.events).toEqual([
      { type: "deleted", baseBlocks: [base], revisedBlocks: [] },
      { type: "inserted", baseBlocks: [], revisedBlocks: [revised] },
    ]);
  });

  test("a stable paragraph relocation survives raw row-coordinate shifts", () => {
    const baseMoved = tableBlock({
      id: "moved",
      text: "Title",
      rowIndex: 1,
      paragraphIndex: 0,
    });
    const revisedMoved = tableBlock({
      id: "moved",
      text: "Updated",
      rowIndex: 2,
      paragraphIndex: 2,
    });
    const base = [
      tableBlock({ id: "header", text: "Header row", rowIndex: 0, paragraphIndex: 0 }),
      baseMoved,
      tableBlock({
        id: "anchor-a",
        text: "First durable anchor text",
        rowIndex: 1,
        paragraphIndex: 1,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second durable anchor text",
        rowIndex: 1,
        paragraphIndex: 2,
      }),
    ];
    const revised = [
      tableBlock({
        id: "inserted-row",
        text: "Inserted unrelated row",
        rowIndex: 0,
        paragraphIndex: 0,
        containerId: "inserted-cell",
      }),
      tableBlock({ id: "header", text: "Header row", rowIndex: 1, paragraphIndex: 0 }),
      tableBlock({
        id: "anchor-a",
        text: "First durable anchor text",
        rowIndex: 2,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second durable anchor text",
        rowIndex: 2,
        paragraphIndex: 1,
      }),
      revisedMoved,
    ];

    const comparison = successfulComparison(base, revised);
    const movedFrom = comparison.events.find(({ type }) => type === "movedFrom");
    const movedTo = comparison.events.find(({ type }) => type === "movedTo");

    expect(movedFrom).toMatchObject({
      type: "movedFrom",
      baseBlocks: [{ id: "moved", table: { rowIndex: 1, cellIndex: 0 } }],
    });
    expect(movedTo).toMatchObject({
      type: "movedTo",
      baseBlockId: "moved",
      revisedBlocks: [{ id: "moved", table: { rowIndex: 2, cellIndex: 0 } }],
    });
    expect(movedTo?.type === "movedTo" ? movedTo.moveId : null).toBe(
      movedFrom?.type === "movedFrom" ? movedFrom.moveId : null,
    );
    expect(comparison.structuralChanges).toEqual([
      {
        id: 1,
        type: "table-row-insert",
        tableIndex: 0,
        rowIndex: 0,
        revisedBlockIds: ["inserted-row"],
      },
    ]);
  });

  test("incompatible cells cannot consume the edited-move comparison budget", () => {
    const incompatible = tableBlock({
      id: "incompatible",
      text: "alpha beta gamma delta epsilon",
      rowIndex: 0,
      cellIndex: 1,
      paragraphIndex: 0,
      containerId: "other-cell",
    });
    const compatible = tableBlock({
      id: "compatible",
      text: "alpha beta gamma delta epsilon",
      rowIndex: 0,
      paragraphIndex: 0,
    });
    const revised = tableBlock({
      id: "revised",
      text: "alpha beta gamma delta zeta",
      rowIndex: 0,
      paragraphIndex: 1,
    });
    const steps = [
      {
        type: "baseOnly",
        block: incompatible,
        moveScope: { bucket: 2, gap: 0 },
      },
      {
        type: "baseOnly",
        block: compatible,
        moveScope: { bucket: 1, gap: 0 },
      },
      {
        type: "revisedOnly",
        block: revised,
        moveScope: { bucket: 1, gap: 1 },
      },
    ] as const satisfies readonly FolioContentAlignmentStep<FolioContentBlock>[];
    const workSession = createContentComparisonWorkSession();
    workSession.remainingMoveComparisons = 1;

    const moves = detectFolioContentMoves({
      steps,
      consumedStepIndexes: new Set(),
      workSession,
      idStability: () => "positional",
    });

    expect(moves).toEqual([{ baseBlock: compatible, revisedBlock: revised }]);
    expect(workSession.remainingMoveComparisons).toBe(0);
  });
});

const aiBlock = (id: string, text: string): FolioAIBlock => ({
  id,
  kind: "paragraph",
  text,
});

const snapshot = (blocks: readonly FolioAIBlock[]): FolioAIEditSnapshot => ({
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
        structuralBoundaryHash: "",
        hashOccurrenceCount: 1,
      },
    ]),
  ),
});

describe("DOCX comparison compatibility", () => {
  test("independently authored same-text paragraphs remain aligned", () => {
    const base = aiBlock("10000001", "Same paragraph text");
    const revised = aiBlock("20000001", "Same paragraph text");

    expect(alignFolioBlocks([base], [revised])).toEqual([
      { type: "pair", baseBlock: base, revisedBlock: revised },
    ]);
  });

  test("the tracked-change planner does not turn a new paragraph id into a move", () => {
    const tail = aiBlock("30000001", "Durable trailing paragraph");
    const base = [aiBlock("10000001", "Same paragraph text"), tail];
    const revised = [aiBlock("20000001", "Same paragraph text"), tail];

    const plan = planStoryCompare({
      story: { type: "main" },
      baseSnapshot: snapshot(base),
      targetSnapshot: snapshot(revised),
      maxOperations: 100,
    });

    expect(plan?.changes).toEqual([]);
    expect(plan?.operations).toEqual([]);
  });
});
