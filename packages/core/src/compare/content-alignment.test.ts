import { describe, expect, test } from "bun:test";

import {
  alignFolioContentBlocks,
  alignFolioContentStructure,
  createFolioContentAlignmentWorkSession,
  longestIncreasingFolioContentPairs,
} from "./content-alignment";
import type { FolioContentBlock } from "./content-types";

const block = (
  id: string,
  text: string,
  options: Partial<FolioContentBlock> = {},
): FolioContentBlock => ({ id, kind: "paragraph", text, ...options });

type CellOptions = {
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex: number;
};

const cell = (
  id: string,
  text: string,
  { rowIndex, cellIndex, gridColumnIndex }: CellOptions,
  options: Partial<FolioContentBlock> = {},
): FolioContentBlock => ({
  id,
  kind: "paragraph",
  text,
  ...options,
  table: {
    outerTableIndex: 0,
    tableIndex: 0,
    rowIndex,
    cellIndex,
    gridColumnIndex,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex: 0,
  },
});

describe("longestIncreasingFolioContentPairs", () => {
  test("resolves crossing equal-length candidates with the historical earliest predecessor", () => {
    const pairs = [
      { baseIndex: 0, revisedIndex: 2 },
      { baseIndex: 1, revisedIndex: 1 },
      { baseIndex: 2, revisedIndex: 3 },
    ];

    const expected = [pairs[0], pairs[2]];
    expect(longestIncreasingFolioContentPairs(pairs)).toEqual(expected);
    expect(longestIncreasingFolioContentPairs(pairs)).toEqual(expected);
  });
});

describe("shared content-alignment LCS work", () => {
  const base = [
    block("base-alpha", "Alpha", { idStability: "positional" }),
    block("base-gamma", "Gamma", { idStability: "positional" }),
  ];
  const revised = [
    block("revised-alpha", "Alpha", { idStability: "positional" }),
    block("revised-epsilon", "Epsilon", { idStability: "positional" }),
    block("revised-gamma", "Gamma", { idStability: "positional" }),
  ];

  test("refuses a later matrix that exceeds the aggregate remainder without underflowing it", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 10 });

    expect(
      alignFolioContentBlocks(base, revised, { workSession }).map(({ type }) => type),
    ).toEqual(["pair", "revisedOnly", "pair"]);
    expect(workSession.remainingLcsCells).toBe(4);

    expect(
      alignFolioContentBlocks(base, revised, { workSession }).map(({ type }) => type),
    ).toEqual(["pair", "pair", "revisedOnly"]);
    expect(workSession.remainingLcsCells).toBe(4);
  });

  test("threads one aggregate allowance through body segments separated by a table", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 10 });
    const firstBase = base.map((entry) => ({ ...entry, id: `first-${entry.id}` }));
    const firstRevised = revised.map((entry) => ({ ...entry, id: `first-${entry.id}` }));
    const secondBase = [
      block("second-base-delta", "Delta", { idStability: "positional" }),
      block("second-base-zeta", "Zeta", { idStability: "positional" }),
    ];
    const secondRevised = [
      block("second-revised-delta", "Delta", { idStability: "positional" }),
      block("second-revised-eta", "Eta", { idStability: "positional" }),
      block("second-revised-zeta", "Zeta", { idStability: "positional" }),
    ];
    const baseTable = cell(
      "base-table",
      "Table boundary",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
    );
    const revisedTable = cell(
      "revised-table",
      "Table boundary",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
    );

    const steps = alignFolioContentStructure({
      baseBlocks: [...firstBase, baseTable, ...secondBase],
      revisedBlocks: [...firstRevised, revisedTable, ...secondRevised],
      workSession,
    });

    expect(steps.map(({ type }) => type)).toEqual([
      "pair",
      "revisedOnly",
      "pair",
      "pair",
      "pair",
      "pair",
      "revisedOnly",
    ]);
    expect(workSession.remainingLcsCells).toBe(4);
    const secondFallback = steps.at(-2);
    expect(secondFallback?.type).toBe("pair");
    if (secondFallback?.type !== "pair") {
      throw new Error("Expected the second body segment to use positional fallback.");
    }
    expect([secondFallback.baseBlock.text, secondFallback.revisedBlock.text]).toEqual([
      "Zeta",
      "Eta",
    ]);
  });
});

describe("container-safe structural alignment", () => {
  test("does not pair equal blocks across distinct generic container paths", () => {
    const base = [
      block("shared", "Same text", {
        containerPath: [{ kind: "section", id: "base-section" }],
      }),
    ];
    const revised = [
      block("shared", "Same text", {
        containerPath: [{ kind: "section", id: "revised-section" }],
      }),
    ];

    expect(alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised })).toEqual([
      { type: "baseOnly", block: base[0] },
      { type: "revisedOnly", block: revised[0] },
    ]);
  });

  test("does not pair equal blocks across distinct table-cell containers", () => {
    const base = [
      cell("shared", "Same text", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }, {
        containerPath: [{ kind: "cell", id: "base-cell" }],
      }),
    ];
    const revised = [
      cell("shared", "Same text", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }, {
        containerPath: [{ kind: "cell", id: "revised-cell" }],
      }),
    ];

    expect(alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised })).toEqual([
      { type: "baseOnly", block: base[0] },
      { type: "revisedOnly", block: revised[0] },
    ]);
  });
});

describe("table column structural alignment", () => {
  test("keeps an inserted column's members in document row order", () => {
    const base = [
      cell("a0", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("b0", "Amount", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      cell("a1", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      cell("b1", "100", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
    ];
    const revised = [
      cell("a0-revised", "Account", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("x0", "Currency", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      cell("b0-revised", "Amount", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      cell("a1-revised", "Fees", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      cell("x1", "EUR", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
      cell("b1-revised", "100", { rowIndex: 1, cellIndex: 2, gridColumnIndex: 2 }),
    ];

    const insertedColumn = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised })
      .filter((step) => step.type === "revisedColumn")
      .at(0);

    expect(insertedColumn?.blocks.map(({ id }) => id)).toEqual(["x0", "x1"]);
  });
});
