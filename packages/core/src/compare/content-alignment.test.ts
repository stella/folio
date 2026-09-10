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
  paragraphIndex?: number;
};

const cell = (
  id: string,
  text: string,
  { rowIndex, cellIndex, gridColumnIndex, paragraphIndex = 0 }: CellOptions,
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
    paragraphIndex,
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

  test("selects at most one candidate for each base coordinate", () => {
    const pairs = [
      { baseIndex: 0, revisedIndex: 0 },
      { baseIndex: 0, revisedIndex: 1 },
      { baseIndex: 1, revisedIndex: 1 },
    ];

    expect(longestIncreasingFolioContentPairs(pairs)).toEqual([pairs[0], pairs[2]]);
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
    const firstBase = base.map(({ id, kind, text, idStability }) => ({
      id: `first-${id}`,
      kind,
      text,
      idStability,
    }));
    const firstRevised = revised.map(({ id, kind, text, idStability }) => ({
      id: `first-${id}`,
      kind,
      text,
      idStability,
    }));
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
      { idStability: "positional" },
    );
    const revisedTable = cell(
      "revised-table",
      "Table boundary",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
      { idStability: "positional" },
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
    // The table cell's 1x1 exact-text matrix shares the same allowance.
    expect(workSession.remainingLcsCells).toBe(3);
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

  test("shares one aggregate allowance across table-cell text matrices", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 10 });
    const tableParagraph = ({
      id,
      text,
      cellIndex,
      paragraphIndex,
    }: {
      id: string;
      text: string;
      cellIndex: number;
      paragraphIndex: number;
    }): FolioContentBlock =>
      cell(
        id,
        text,
        { rowIndex: 0, cellIndex, gridColumnIndex: cellIndex, paragraphIndex },
        { idStability: "positional" },
      );
    const baseBlocks = [
      tableParagraph({
        id: "base-alpha",
        text: "Alpha",
        cellIndex: 0,
        paragraphIndex: 0,
      }),
      tableParagraph({
        id: "base-gamma",
        text: "Gamma",
        cellIndex: 0,
        paragraphIndex: 1,
      }),
      tableParagraph({
        id: "base-delta",
        text: "Delta",
        cellIndex: 1,
        paragraphIndex: 0,
      }),
      tableParagraph({
        id: "base-zeta",
        text: "Zeta",
        cellIndex: 1,
        paragraphIndex: 1,
      }),
    ];
    const revisedBlocks = [
      tableParagraph({
        id: "revised-alpha",
        text: "Alpha",
        cellIndex: 0,
        paragraphIndex: 0,
      }),
      tableParagraph({
        id: "revised-epsilon",
        text: "Epsilon",
        cellIndex: 0,
        paragraphIndex: 1,
      }),
      tableParagraph({
        id: "revised-gamma",
        text: "Gamma",
        cellIndex: 0,
        paragraphIndex: 2,
      }),
      tableParagraph({
        id: "revised-delta",
        text: "Delta",
        cellIndex: 1,
        paragraphIndex: 0,
      }),
      tableParagraph({
        id: "revised-eta",
        text: "Eta",
        cellIndex: 1,
        paragraphIndex: 1,
      }),
      tableParagraph({
        id: "revised-zeta",
        text: "Zeta",
        cellIndex: 1,
        paragraphIndex: 2,
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks, revisedBlocks, workSession });

    expect(steps.map(({ type }) => type)).toEqual([
      "pair",
      "revisedOnly",
      "pair",
      "pair",
      "pair",
      "revisedOnly",
    ]);
    expect(workSession.remainingLcsCells).toBe(4);
    const secondCellFallback = steps.at(-2);
    if (secondCellFallback?.type !== "pair") {
      throw new Error("Expected the second table cell to use positional fallback.");
    }
    expect([
      secondCellFallback.baseBlock.text,
      secondCellFallback.revisedBlock.text,
    ]).toEqual(["Zeta", "Eta"]);
  });
});

describe("residual identity continuity", () => {
  test("keeps a persisted positional block paired past a new sibling", () => {
    const base = [
      block("before", "Before anchor"),
      block("persisted", "Centered paragraph.", { idStability: "positional" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("before", "Before anchor"),
      block("inserted", "New unrelated paragraph"),
      block("persisted", "Changed paragraph."),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.id, event.revisedBlock.id]
          : [event.type, event.block.id],
      ),
    ).toEqual([
      ["pair", "before", "before"],
      ["revisedOnly", "inserted"],
      ["pair", "persisted", "persisted"],
      ["pair", "after", "after"],
    ]);
  });

  test("does not promote equality between two positional ids to identity", () => {
    const base = [
      block("0", "Before anchor"),
      block("1", "Original paragraph", { idStability: "positional" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("0", "Before anchor"),
      block("1", "Inserted paragraph", { idStability: "positional" }),
      block("2", "Original paragraph", { idStability: "positional" }),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.id, event.revisedBlock.id]
          : [event.type, event.block.id],
      ),
    ).toEqual([
      ["pair", "0", "0"],
      ["revisedOnly", "1"],
      ["pair", "1", "2"],
      ["pair", "after", "after"],
    ]);
  });

  test("lets shifted exact text outrank an asymmetric id coincidence", () => {
    const base = [
      block("before", "Before anchor"),
      block("0", "Original paragraph", { idStability: "positional" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("before", "Before anchor"),
      block("0", "New preceding paragraph"),
      block("survivor", "Original paragraph"),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.id, event.revisedBlock.id]
          : [event.type, event.block.id],
      ),
    ).toEqual([
      ["pair", "before", "before"],
      ["revisedOnly", "0"],
      ["pair", "0", "survivor"],
      ["pair", "after", "after"],
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
      { type: "baseOnly", block: base[0], moveScope: { bucket: 0, gap: 0 } },
      { type: "revisedOnly", block: revised[0], moveScope: { bucket: 0, gap: 1 } },
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
      { type: "baseOnly", block: base[0], moveScope: { bucket: 1, gap: 0 } },
      { type: "revisedOnly", block: revised[0], moveScope: { bucket: 2, gap: 0 } },
    ]);
  });

  test("multiple stable blocks cannot pair one body segment to two revised segments", () => {
    const baseTable = cell("separator", "Table separator", {
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
    });
    const revisedTable = cell("separator", "Table separator", {
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
    });
    const base = [
      block("anchor-a", "First body anchor"),
      block("relocated", "Relocated body content"),
      baseTable,
      block("anchor-b", "Second body anchor"),
    ];
    const revised = [
      block("anchor-a", "First body anchor"),
      revisedTable,
      block("anchor-b", "Second body anchor"),
      block("relocated", "Relocated body content"),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.id, step.revisedBlock.id]] : [],
      ),
    ).toEqual([
      ["anchor-a", "anchor-a"],
      ["separator", "separator"],
      ["anchor-b", "anchor-b"],
    ]);
  });
});

describe("table row and column structural alignment", () => {
  test("keeps a same-position persisted row paired through a full cell rewrite", () => {
    const base = [
      cell(
        "persisted-cell",
        "A1",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { idStability: "positional" },
      ),
      cell(
        "second-cell",
        "B1",
        { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 },
        { idStability: "positional" },
      ),
      cell(
        "third-cell",
        "C1",
        { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 },
        { idStability: "positional" },
      ),
      cell("next-row", "A2", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
    ];
    const revised = [
      cell("persisted-cell", "Entirely rewritten cell text", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("second-cell", "B1", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      cell("third-cell", "C1", { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 }),
      cell("next-row", "A2", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
      stableIdMismatch: "pair",
    });

    expect(steps.map(({ type }) => type)).toEqual(["pair", "pair", "pair", "pair"]);
    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.id, step.revisedBlock.id]] : [],
      ),
    ).toEqual([
      ["persisted-cell", "persisted-cell"],
      ["second-cell", "second-cell"],
      ["third-cell", "third-cell"],
      ["next-row", "next-row"],
    ]);
  });

  test("keeps shifted rows paired by logical row before considering stable ids", () => {
    const base = [
      cell("first-id", "First logical row phrase", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("second-id", "Second logical row phrase", {
        rowIndex: 1,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
    ];
    const revised = [
      cell("inserted-id", "Inserted unrelated row", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("second-id", "First logical row phrase", {
        rowIndex: 1,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("first-id", "Second logical row phrase", {
        rowIndex: 2,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
      stableIdMismatch: "pair",
    });
    const pairs = steps.filter((step) => step.type === "pair");

    expect(steps.map(({ type }) => type)).toEqual(["revisedRow", "pair", "pair"]);
    expect(
      pairs.map(({ baseBlock, revisedBlock }) => [
        baseBlock.table?.rowIndex,
        revisedBlock.table?.rowIndex,
        baseBlock.text,
        revisedBlock.text,
      ]),
    ).toEqual([
      [0, 1, "First logical row phrase", "First logical row phrase"],
      [1, 2, "Second logical row phrase", "Second logical row phrase"],
    ]);
  });

  test("keeps shifted columns paired by logical column before considering stable ids", () => {
    const base = [
      cell("left-id", "Account owner", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("right-id", "Amount payable", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
    ];
    const revised = [
      cell("inserted-id", "Currency code", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("right-id", "Account owner", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("left-id", "Amount payable", {
        rowIndex: 0,
        cellIndex: 2,
        gridColumnIndex: 2,
      }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
      stableIdMismatch: "pair",
    });
    const pairs = steps.filter((step) => step.type === "pair");

    expect(steps.map(({ type }) => type)).toEqual(["revisedColumn", "pair", "pair"]);
    expect(
      pairs.map(({ baseBlock, revisedBlock }) => [
        baseBlock.table?.gridColumnIndex,
        revisedBlock.table?.gridColumnIndex,
        baseBlock.text,
        revisedBlock.text,
      ]),
    ).toEqual([
      [0, 1, "Account owner", "Account owner"],
      [1, 2, "Amount payable", "Amount payable"],
    ]);
  });

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

    const insertedColumn = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    }).find((step) => step.type === "revisedColumn");

    expect(insertedColumn?.blocks.map(({ id }) => id)).toEqual(["x0", "x1"]);
  });
});
