import { describe, expect, test } from "bun:test";

import {
  alignFolioContentBlocks,
  alignFolioContentStructure,
  createFolioContentAlignmentWorkSession,
  longestIncreasingFolioContentPairs,
} from "./content-alignment";
import {
  contentBlockFixture as block,
  contentIdentity,
  tableLocationFixture,
  type ContentBlockFixtureOptions,
} from "./content-test-fixtures";
import type { FolioContentBlock, FolioContentIdentitySemantics } from "./content-types";

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
  options: ContentBlockFixtureOptions = {},
): FolioContentBlock =>
  block(id, text, {
    identityType: "persistent-hint",
    ...options,
    table: tableLocationFixture({
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex,
    }),
  });

const outerTableCell = (id: string, outerTableIndex: number): FolioContentBlock =>
  block(id, `Table ${String(outerTableIndex)}`, {
    identityType: "persistent-hint",
    table: tableLocationFixture({
      outerTableIndex,
      tableIndex: outerTableIndex,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    }),
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
    block("base-alpha", "Alpha", { identityType: "positional" }),
    block("base-gamma", "Gamma", { identityType: "positional" }),
  ];
  const revised = [
    block("revised-alpha", "Alpha", { identityType: "positional" }),
    block("revised-epsilon", "Epsilon", { identityType: "positional" }),
    block("revised-gamma", "Gamma", { identityType: "positional" }),
  ];

  test("refuses a later matrix that exceeds the aggregate remainder without underflowing it", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 10 });

    expect(alignFolioContentBlocks(base, revised, { workSession }).map(({ type }) => type)).toEqual(
      ["pair", "revisedOnly", "pair"],
    );
    expect(workSession.remainingLcsCells).toBe(4);

    expect(alignFolioContentBlocks(base, revised, { workSession }).map(({ type }) => type)).toEqual(
      ["pair", "pair", "revisedOnly"],
    );
    expect(workSession.remainingLcsCells).toBe(4);
  });

  test("threads one aggregate allowance through body segments separated by a table", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 10 });
    const firstBase = base.map(({ identity, text }) =>
      block(`first-${identity.id}`, text, { identityType: identity.type }),
    );
    const firstRevised = revised.map(({ identity, text }) =>
      block(`first-${identity.id}`, text, { identityType: identity.type }),
    );
    const secondBase = [
      block("second-base-delta", "Delta", { identityType: "positional" }),
      block("second-base-zeta", "Zeta", { identityType: "positional" }),
    ];
    const secondRevised = [
      block("second-revised-delta", "Delta", { identityType: "positional" }),
      block("second-revised-eta", "Eta", { identityType: "positional" }),
      block("second-revised-zeta", "Zeta", { identityType: "positional" }),
    ];
    const baseTable = cell(
      "base-table",
      "Table boundary",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
      { identityType: "positional" },
    );
    const revisedTable = cell(
      "revised-table",
      "Table boundary",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
      { identityType: "positional" },
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
        { identityType: "positional" },
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
    expect([secondCellFallback.baseBlock.text, secondCellFallback.revisedBlock.text]).toEqual([
      "Zeta",
      "Eta",
    ]);
  });
});

describe("residual identity continuity", () => {
  test("keeps a persisted positional block paired past a new sibling", () => {
    const base = [
      block("before", "Before anchor"),
      block("persisted", "Centered paragraph.", { identityType: "persistent-hint" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("before", "Before anchor"),
      block("inserted", "New unrelated paragraph"),
      block("persisted", "Changed paragraph.", { identityType: "persistent-hint" }),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.identity.id, event.revisedBlock.identity.id]
          : [event.type, event.block.identity.id],
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
      block("1", "Original paragraph", { identityType: "positional" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("0", "Before anchor"),
      block("1", "Inserted paragraph", { identityType: "positional" }),
      block("2", "Original paragraph", { identityType: "positional" }),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.identity.id, event.revisedBlock.identity.id]
          : [event.type, event.block.identity.id],
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
      block("0", "Original paragraph", { identityType: "positional" }),
      block("after", "After anchor"),
    ];
    const revised = [
      block("before", "Before anchor"),
      block("0", "New preceding paragraph", { identityType: "positional" }),
      block("survivor", "Original paragraph", { identityType: "positional" }),
      block("after", "After anchor"),
    ];

    expect(
      alignFolioContentBlocks(base, revised).map((event) =>
        event.type === "pair"
          ? [event.type, event.baseBlock.identity.id, event.revisedBlock.identity.id]
          : [event.type, event.block.identity.id],
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
  test.each([
    { tableSide: "base", oneSidedType: "baseTable" },
    { tableSide: "revised", oneSidedType: "revisedTable" },
  ] as const)(
    "keeps both body anchors paired when a table exists only on the $tableSide side",
    ({ tableSide, oneSidedType }) => {
      const before = block("before", "Stable paragraph before the table");
      const after = block("after", "Stable paragraph after the table");
      const table = outerTableCell("table", 0);
      const uninterrupted = [before, after];
      const separated = [before, table, after];
      const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 0 });
      const steps = alignFolioContentStructure({
        baseBlocks: tableSide === "base" ? separated : uninterrupted,
        revisedBlocks: tableSide === "revised" ? separated : uninterrupted,
        workSession,
      });

      expect(
        steps.map((step) =>
          step.type === "pair"
            ? [step.type, step.baseBlock.identity.id, step.revisedBlock.identity.id]
            : [step.type],
        ),
      ).toEqual([["pair", "before", "before"], [oneSidedType], ["pair", "after", "after"]]);
      expect(workSession.remainingLcsCells).toBe(0);
    },
  );

  test.each([
    { tableSide: "base", oneSidedType: "baseTable" },
    { tableSide: "revised", oneSidedType: "revisedTable" },
  ] as const)(
    "projects every $tableSide table boundary into one uninterrupted body run",
    ({ tableSide, oneSidedType }) => {
      const alpha = block("alpha", "Stable alpha paragraph");
      const beta = block("beta", "Stable beta paragraph");
      const gamma = block("gamma", "Stable gamma paragraph");
      const firstTable = outerTableCell("first-table", 0);
      const secondTable = outerTableCell("second-table", 1);
      const uninterrupted = [alpha, beta, gamma];
      const separated = [alpha, firstTable, beta, secondTable, gamma];
      const steps = alignFolioContentStructure({
        baseBlocks: tableSide === "base" ? separated : uninterrupted,
        revisedBlocks: tableSide === "revised" ? separated : uninterrupted,
      });

      expect(
        steps.map((step) =>
          step.type === "pair"
            ? [step.type, step.baseBlock.identity.id, step.revisedBlock.identity.id]
            : step.type === "baseTable" || step.type === "revisedTable"
              ? [step.type, step.blocks.at(0)?.identity.id]
              : [step.type],
        ),
      ).toEqual([
        ["pair", "alpha", "alpha"],
        [oneSidedType, "first-table"],
        ["pair", "beta", "beta"],
        [oneSidedType, "second-table"],
        ["pair", "gamma", "gamma"],
      ]);
    },
  );

  test("uses unique exact content when positional ids shift across a removed table", () => {
    const base = [
      block("base-before", "Exact paragraph before", { identityType: "positional" }),
      outerTableCell("removed-table", 0),
      block("base-after", "Exact paragraph after", { identityType: "positional" }),
    ];
    const revised = [
      block("revised-before", "Exact paragraph before", { identityType: "positional" }),
      block("revised-after", "Exact paragraph after", { identityType: "positional" }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(
      steps.map((step) =>
        step.type === "pair"
          ? [step.type, step.baseBlock.identity.id, step.revisedBlock.identity.id]
          : [step.type],
      ),
    ).toEqual([
      ["pair", "base-before", "revised-before"],
      ["baseTable"],
      ["pair", "base-after", "revised-after"],
    ]);
  });

  test("does not use repeated boilerplate to project a table boundary", () => {
    const repeated = "Standard terms apply";
    const base = [
      block("base-copy-before", repeated, { identityType: "positional" }),
      outerTableCell("removed-table", 0),
      block("base-copy-after", repeated, { identityType: "positional" }),
    ];
    const revised = [
      block("revised-copy-before", repeated, { identityType: "positional" }),
      block("revised-copy-after", repeated, { identityType: "positional" }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.filter((step) => step.type === "pair")).toHaveLength(1);
    expect(steps.map(({ type }) => type)).toContain("baseTable");
  });

  test("owns the predecessor and revised carrier for a terminal paragraph removal", () => {
    const alpha = block("alpha", "Alpha");
    const beta = block("beta", "Beta");
    const gamma = block("gamma", "Gamma");
    const revisedGamma = block("gamma", "Gamma");
    const revisedAlpha = block("alpha", "Alpha");
    const revisedBeta = block("beta", "Beta");

    const steps = alignFolioContentStructure({
      baseBlocks: [alpha, beta, gamma],
      revisedBlocks: [revisedGamma, revisedAlpha, revisedBeta],
    });
    const source = steps.find(
      (step) => step.type === "baseOnly" && step.block.identity.id === "gamma",
    );
    if (source?.type !== "baseOnly") throw new Error("Expected the moved terminal source.");

    expect(source.removalBoundary).toEqual({
      type: "terminalPredecessor",
      predecessor: beta,
      targetCarrier: revisedBeta,
      containerAlignment: source.moveScope.containerAlignment,
    });
  });

  test("owns an exact successor for a non-terminal paragraph removal", () => {
    const alpha = block("alpha", "Alpha");
    const beta = block("beta", "Beta");
    const gamma = block("gamma", "Gamma");
    const steps = alignFolioContentStructure({
      baseBlocks: [alpha, beta, gamma],
      revisedBlocks: [beta, alpha, gamma],
    });
    const source = steps.find(
      (step) => step.type === "baseOnly" && step.block.identity.id === "beta",
    );
    if (source?.type !== "baseOnly") throw new Error("Expected the moved non-terminal source.");

    expect(source.removalBoundary).toEqual({
      type: "successorParagraph",
      successor: gamma,
      containerAlignment: source.moveScope.containerAlignment,
    });
  });

  test("does not borrow a removal boundary across nested paths in one physical cell", () => {
    const firstPath = [{ kind: "contentControl", identity: contentIdentity("first-path") }];
    const secondPath = [{ kind: "contentControl", identity: contentIdentity("second-path") }];
    const removed = cell(
      "removed",
      "Removed",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
      { containerPath: firstPath },
    );
    const survivor = cell(
      "survivor",
      "Survivor",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0, paragraphIndex: 1 },
      { containerPath: secondPath },
    );
    const revisedSurvivor = cell(
      "survivor",
      "Survivor",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
      { containerPath: secondPath },
    );
    const replacement = cell(
      "replacement",
      "Replacement",
      { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0, paragraphIndex: 1 },
      { containerPath: firstPath },
    );

    const steps = alignFolioContentStructure({
      baseBlocks: [removed, survivor],
      revisedBlocks: [revisedSurvivor, replacement],
    });
    const deletion = steps.find(
      (step) => step.type === "baseOnly" && step.block.identity.id === "removed",
    );
    const pair = steps.find(
      (step) => step.type === "pair" && step.baseBlock.identity.id === "survivor",
    );
    if (deletion?.type !== "baseOnly" || pair?.type !== "pair") {
      throw new Error("Expected one nested-container deletion and one surviving pair.");
    }

    expect(deletion.removalBoundary).toEqual({
      type: "unanchoredContainer",
      containerAlignment: deletion.moveScope.containerAlignment,
    });
    expect(deletion.moveScope.containerAlignment).not.toBe(pair.containerAlignment);
  });

  test("anchors a revised run only to paragraphs in its paired table cell", () => {
    const base = [
      cell("a0", "Alpha", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("b0", "Delete", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
        paragraphIndex: 0,
      }),
      cell("b1", "Keep", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
        paragraphIndex: 1,
      }),
    ];
    const revised = [
      cell("a0", "Alpha", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("a1", "One", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 1,
      }),
      cell("a2", "Two", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 2,
      }),
      cell("b1", "Keep", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    const inserted = steps.filter((step) => step.type === "revisedOnly");
    expect(inserted.map(({ block }) => block.identity.id)).toEqual(["a1", "a2"]);
    for (const step of inserted) {
      expect(step.insertionBoundary.type).toBe("afterParagraph");
      if (step.insertionBoundary.type === "unanchoredContainer") {
        throw new Error("Expected the first cell to retain its own paragraph boundary.");
      }
      expect(step.insertionBoundary.paragraph.identity.id).toBe("a0");
      expect(step.insertionBoundary.paragraph.table?.cellIndex).toBe(0);
      expect(step.insertionBoundary.containerAlignment).toBe(step.moveScope.containerAlignment);
    }
    const deleted = steps.find(
      (step) => step.type === "baseOnly" && step.block.identity.id === "b0",
    );
    const kept = steps.find((step) => step.type === "pair" && step.baseBlock.identity.id === "b1");
    if (deleted?.type !== "baseOnly" || kept?.type !== "pair") {
      throw new Error("Expected the neighboring cell deletion and survivor.");
    }
    expect(deleted.moveScope.containerAlignment).toBe(kept.containerAlignment);
    expect(deleted.moveScope.containerAlignment).not.toBe(
      inserted.at(0)?.moveScope.containerAlignment,
    );
  });

  test("uses the next paragraph in the same cell for a leading insertion", () => {
    const base = [
      cell("anchor", "Anchor", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("other", "Other cell", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
    ];
    const revised = [
      cell(
        "inserted",
        "Inserted",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "authoritative" },
      ),
      cell("anchor", "Anchor", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
        paragraphIndex: 1,
      }),
      cell("other", "Other cell", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
    ];

    const step = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised }).find(
      (candidate) => candidate.type === "revisedOnly",
    );
    if (step?.type !== "revisedOnly") throw new Error("Expected a leading insertion.");
    expect(step.insertionBoundary.type).toBe("beforeParagraph");
    if (step.insertionBoundary.type === "unanchoredContainer") {
      throw new Error("Expected a cell-local paragraph boundary.");
    }
    expect(step.insertionBoundary.paragraph.identity.id).toBe("anchor");
    expect(step.insertionBoundary.paragraph.table?.cellIndex).toBe(0);
  });

  test("uses a deleted paragraph as the typed boundary for total cell replacement", () => {
    const base = [
      cell(
        "old",
        "Payment is due within thirty days",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "authoritative" },
      ),
    ];
    const revised = [
      cell(
        "new",
        "Payment is due within sixty days",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "authoritative" },
      ),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    const deletion = steps.find((step) => step.type === "baseOnly");
    const insertion = steps.find((step) => step.type === "revisedOnly");
    if (deletion?.type !== "baseOnly" || insertion?.type !== "revisedOnly") {
      throw new Error("Expected one deletion and one insertion.");
    }
    expect(deletion.moveScope.containerAlignment.type).toBe("paired");
    expect(insertion.moveScope.containerAlignment).toBe(deletion.moveScope.containerAlignment);
    expect(insertion.insertionBoundary.type).toBe("afterParagraph");
    if (insertion.insertionBoundary.type === "unanchoredContainer") {
      throw new Error("Expected the deleted cell paragraph to remain a boundary cursor.");
    }
    expect(insertion.insertionBoundary.paragraph.identity.id).toBe("old");
  });

  test("records a structural sibling as the true end of a table-terminated body", () => {
    const base = [
      block("body", "Body"),
      cell("table", "Table", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
    ];
    const revised = [
      block("body", "Body"),
      block("inserted", "Inserted", { identityType: "authoritative" }),
      cell("table", "Table", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    const bodyPair = steps.find(
      (step) => step.type === "pair" && step.baseBlock.identity.id === "body",
    );
    const insertion = steps.find(
      (step) => step.type === "revisedOnly" && step.block.identity.id === "inserted",
    );
    if (bodyPair?.type !== "pair" || insertion?.type !== "revisedOnly") {
      throw new Error("Expected a body pair and insertion.");
    }
    expect(bodyPair.containerAlignment).toBe(insertion.moveScope.containerAlignment);
    expect(bodyPair.containerAlignment.base.end).toBe("structuralSibling");
    expect(bodyPair.containerAlignment.revised.end).toBe("structuralSibling");
  });

  test("does not pair equal blocks across distinct generic container paths", () => {
    const base = [
      block("shared", "Same text", {
        containerPath: [{ kind: "section", identity: contentIdentity("base-section") }],
      }),
    ];
    const revised = [
      block("shared", "Same text", {
        containerPath: [{ kind: "section", identity: contentIdentity("revised-section") }],
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    expect(steps.map(({ type }) => type)).toEqual(["baseOnly", "revisedOnly"]);
    const baseOnly = steps.at(0);
    const revisedOnly = steps.at(1);
    if (baseOnly?.type !== "baseOnly" || revisedOnly?.type !== "revisedOnly") {
      throw new Error("Expected distinct one-sided container occurrences.");
    }
    expect(baseOnly.moveScope.containerAlignment).toMatchObject({
      type: "baseOnly",
      base: { type: "body", containerPath: base[0]?.containerPath },
      revised: null,
    });
    expect(revisedOnly.moveScope.containerAlignment).toMatchObject({
      type: "revisedOnly",
      base: null,
      revised: { type: "body", containerPath: revised[0]?.containerPath },
    });
    expect(revisedOnly.insertionBoundary).toEqual({
      type: "unanchoredContainer",
      containerAlignment: revisedOnly.moveScope.containerAlignment,
    });
  });

  test("does not pair equal blocks across distinct table-cell containers", () => {
    const base = [
      cell(
        "shared",
        "Same text",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        {
          containerPath: [{ kind: "cell", identity: contentIdentity("base-cell") }],
        },
      ),
    ];
    const revised = [
      cell(
        "shared",
        "Same text",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        {
          containerPath: [{ kind: "cell", identity: contentIdentity("revised-cell") }],
        },
      ),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    expect(steps.map(({ type }) => type)).toEqual(["baseTable", "revisedTable"]);
    expect(
      steps.map((step) =>
        step.type === "baseTable" || step.type === "revisedTable"
          ? step.blocks.map(({ identity }) => identity.id)
          : [],
      ),
    ).toEqual([["shared"], ["shared"]]);
  });

  test("does not treat distinct positional container hints as one table-cell occurrence", () => {
    const base = [
      cell(
        "shared",
        "Same text",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        {
          containerPath: [
            { kind: "tableCell", identity: contentIdentity("parent-cell-0", "positional") },
          ],
        },
      ),
    ];
    const revised = [
      cell(
        "shared",
        "Same text",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        {
          containerPath: [
            { kind: "tableCell", identity: contentIdentity("parent-cell-1", "positional") },
          ],
        },
      ),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["baseOnly", "revisedOnly"]);
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
        step.type === "pair" ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]] : [],
      ),
    ).toEqual([
      ["anchor-a", "anchor-a"],
      ["separator", "separator"],
      ["anchor-b", "anchor-b"],
    ]);
  });
});

describe("table row and column structural alignment", () => {
  test("keeps several shifted persisted rows paired without textual anchors", () => {
    const base = [
      cell(
        "removed-a",
        "A1",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "removed-b",
        "B1",
        { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell(
        "first-a",
        "A2",
        { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "first-b",
        "B2",
        { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell(
        "second-a",
        "A3",
        { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "second-b",
        "B3",
        { rowIndex: 2, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
    ];
    const revised = [
      cell("first-a", "Repeated rewritten value", {
        rowIndex: 0,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("first-b", "Repeated rewritten value", {
        rowIndex: 0,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("second-a", "Repeated rewritten value", {
        rowIndex: 1,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("second-b", "Repeated rewritten value", {
        rowIndex: 1,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    });

    expect(
      steps.flatMap((step) =>
        step.type === "baseRow" || step.type === "revisedRow"
          ? [[step.type, step.blocks.map(({ identity }) => identity.id)]]
          : [],
      ),
    ).toEqual([["baseRow", ["removed-a", "removed-b"]]]);
    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]] : [],
      ),
    ).toEqual([
      ["first-a", "first-a"],
      ["first-b", "first-b"],
      ["second-a", "second-a"],
      ["second-b", "second-b"],
    ]);
  });

  test.each([
    {
      identity: "complete",
      revisedSecondIds: ["second-base-only", "first-base-only"],
    },
    {
      identity: "partial",
      revisedSecondIds: ["second-revised-only", "first-revised-only"],
    },
  ] as const)(
    "does not cross-pair reordered rows with $identity persisted identity",
    ({ revisedSecondIds }) => {
      const base = [
        cell(
          "first-shared",
          "Original first A",
          { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
          { identityType: "positional" },
        ),
        cell(
          "first-base-only",
          "Original first B",
          { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 },
          { identityType: "positional" },
        ),
        cell(
          "second-shared",
          "Original second A",
          { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 },
          { identityType: "positional" },
        ),
        cell(
          "second-base-only",
          "Original second B",
          { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 },
          { identityType: "positional" },
        ),
      ];
      const revised = [
        cell("second-shared", "Unrelated replacement A", {
          rowIndex: 0,
          cellIndex: 0,
          gridColumnIndex: 0,
        }),
        cell(revisedSecondIds[0], "Unrelated replacement B", {
          rowIndex: 0,
          cellIndex: 1,
          gridColumnIndex: 1,
        }),
        cell("first-shared", "Different replacement A", {
          rowIndex: 1,
          cellIndex: 0,
          gridColumnIndex: 0,
        }),
        cell(revisedSecondIds[1], "Different replacement B", {
          rowIndex: 1,
          cellIndex: 1,
          gridColumnIndex: 1,
        }),
      ];

      const steps = alignFolioContentStructure({
        baseBlocks: base,
        revisedBlocks: revised,
      });

      expect(steps.filter(({ type }) => type === "pair")).toEqual([]);
      expect(
        steps
          .flatMap((step) =>
            step.type === "baseRow" || step.type === "revisedRow" ? [step.type] : [],
          )
          .toSorted(),
      ).toEqual(["baseRow", "baseRow", "revisedRow", "revisedRow"]);
    },
  );

  test("keeps an edited persisted row paired after an inserted row", () => {
    const base = [
      cell(
        "anchor-a",
        "A1",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "anchor-b",
        "B1",
        { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell(
        "persisted-a",
        "A2",
        { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "persisted-b",
        "B2",
        { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell(
        "tail-a",
        "A3",
        { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "tail-b",
        "B3",
        { rowIndex: 2, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
    ];
    const revised = [
      cell("anchor-a", "A1", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("anchor-b", "B1", { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
      cell("inserted-a", "New A", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      cell("inserted-b", "New B", { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 }),
      cell("persisted-a", "A2", { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 }),
      cell("persisted-b", "Entirely different wording", {
        rowIndex: 2,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("tail-a", "A3", { rowIndex: 3, cellIndex: 0, gridColumnIndex: 0 }),
      cell("tail-b", "B3", { rowIndex: 3, cellIndex: 1, gridColumnIndex: 1 }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    });

    expect(
      steps.flatMap((step) =>
        step.type === "baseRow" || step.type === "revisedRow"
          ? [[step.type, step.blocks.map(({ identity }) => identity.id)]]
          : [],
      ),
    ).toEqual([["revisedRow", ["inserted-a", "inserted-b"]]]);
    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]] : [],
      ),
    ).toContainEqual(["persisted-b", "persisted-b"]);
  });

  test("prefers a rewritten persisted row over a similar inserted row", () => {
    const base = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell(
        "persisted-a",
        "Payment shall be made within thirty calendar days",
        { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "persisted-b",
        "Written notice must be delivered to the address",
        { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell("after", "After", { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 }),
    ];
    const revised = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("inserted-a", "Payment shall be made within forty calendar days", {
        rowIndex: 1,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("inserted-b", "Written notice must be delivered to the address", {
        rowIndex: 1,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("persisted-a", "Completely rewritten first cell", {
        rowIndex: 2,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("persisted-b", "Entirely different second cell", {
        rowIndex: 2,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("after", "After", { rowIndex: 3, cellIndex: 0, gridColumnIndex: 0 }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    });

    expect(
      steps.flatMap((step) =>
        step.type === "baseRow" || step.type === "revisedRow"
          ? [[step.type, step.blocks.map(({ identity }) => identity.id)]]
          : [],
      ),
    ).toEqual([["revisedRow", ["inserted-a", "inserted-b"]]]);
    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]] : [],
      ),
    ).toContainEqual(["persisted-a", "persisted-a"]);
  });

  test("does not pair rows from a partial persisted id overlap", () => {
    const base = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell(
        "shared",
        "Original A",
        { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "base-only",
        "Original B",
        { rowIndex: 1, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell("after", "After", { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 }),
    ];
    const revised = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("shared", "Different X", {
        rowIndex: 1,
        cellIndex: 0,
        gridColumnIndex: 0,
      }),
      cell("revised-only", "Different Y", {
        rowIndex: 1,
        cellIndex: 1,
        gridColumnIndex: 1,
      }),
      cell("after", "After", { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    });

    expect(
      steps.flatMap((step) =>
        step.type === "baseRow" || step.type === "revisedRow"
          ? [[step.type, step.blocks.map(({ identity }) => identity.id)]]
          : [],
      ),
    ).toEqual([
      ["baseRow", ["shared", "base-only"]],
      ["revisedRow", ["shared", "revised-only"]],
    ]);
  });

  test("does not use persisted row identity beyond the bounded profile", () => {
    const oversizedRowLength = 129;
    type OversizedRowOptions = {
      rowIndex: number;
      textPrefix: string;
      identityType: FolioContentIdentitySemantics;
    };
    const row = ({
      rowIndex,
      textPrefix,
      identityType,
    }: OversizedRowOptions): FolioContentBlock[] =>
      Array.from({ length: oversizedRowLength }, (_, paragraphIndex) =>
        cell(
          `persisted-${String(paragraphIndex)}`,
          `${textPrefix}-${String(paragraphIndex)}`,
          { rowIndex, cellIndex: 0, gridColumnIndex: 0, paragraphIndex },
          { identityType },
        ),
      );
    const base = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      ...row({ rowIndex: 1, textPrefix: "original", identityType: "positional" }),
      cell("after", "After", { rowIndex: 2, cellIndex: 0, gridColumnIndex: 0 }),
    ];
    const revised = [
      cell("before", "Before", { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 }),
      cell("inserted", "Inserted", { rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
      ...row({
        rowIndex: 2,
        textPrefix: "replacement",
        identityType: "persistent-hint",
      }),
      cell("after", "After", { rowIndex: 3, cellIndex: 0, gridColumnIndex: 0 }),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
    });

    expect(
      steps.flatMap((step) =>
        step.type === "pair" && step.baseBlock.identity.id.startsWith("persisted-")
          ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]]
          : [],
      ),
    ).toEqual([]);
    expect(
      steps.flatMap((step) =>
        step.type === "baseRow" || step.type === "revisedRow" ? [step.type] : [],
      ),
    ).toEqual(["baseRow", "revisedRow", "revisedRow"]);
  });

  test("keeps a same-position persisted row paired through a full cell rewrite", () => {
    const base = [
      cell(
        "persisted-cell",
        "A1",
        { rowIndex: 0, cellIndex: 0, gridColumnIndex: 0 },
        { identityType: "positional" },
      ),
      cell(
        "second-cell",
        "B1",
        { rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 },
        { identityType: "positional" },
      ),
      cell(
        "third-cell",
        "C1",
        { rowIndex: 0, cellIndex: 2, gridColumnIndex: 2 },
        { identityType: "positional" },
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
    });

    expect(steps.map(({ type }) => type)).toEqual(["pair", "pair", "pair", "pair"]);
    expect(
      steps.flatMap((step) =>
        step.type === "pair" ? [[step.baseBlock.identity.id, step.revisedBlock.identity.id]] : [],
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

  test("does not cross stable identities to force a shifted column embedding", () => {
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
    });
    expect(steps.map(({ type }) => type)).toEqual(["tableReplacement"]);
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

    expect(insertedColumn?.blocks.map(({ identity }) => identity.id)).toEqual(["x0", "x1"]);
  });
});
