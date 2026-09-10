import { describe, expect, test } from "bun:test";

import {
  alignFolioContentStructure,
  createFolioContentAlignmentWorkSession,
} from "./content-alignment";
import type { FolioContentBlock, FolioContentIdStability } from "./content-types";

type TableBlockOptions = {
  id: string;
  text: string;
  outerTableIndex: number;
  tableIndex?: number;
  rowIndex?: number;
  paragraphIndex?: number;
  idStability?: FolioContentIdStability;
};

const tableBlock = ({
  id,
  text,
  outerTableIndex,
  tableIndex = outerTableIndex,
  rowIndex = 0,
  paragraphIndex = 0,
  idStability,
}: TableBlockOptions): FolioContentBlock => ({
  id,
  kind: "paragraph",
  text,
  idStability,
  table: {
    outerTableIndex,
    tableIndex,
    rowIndex,
    cellIndex: 0,
    gridColumnIndex: 0,
    columnSpan: 1,
    rowSpan: 1,
    paragraphIndex,
  },
});

const tableSequence = (
  entries: readonly Omit<TableBlockOptions, "outerTableIndex">[],
): FolioContentBlock[] =>
  entries.map((entry, outerTableIndex) => tableBlock({ ...entry, outerTableIndex }));

const pairIds = (
  steps: ReturnType<typeof alignFolioContentStructure>,
): [string, string][] =>
  steps.flatMap((step) =>
    step.type === "pair" ? [[step.baseBlock.id, step.revisedBlock.id]] : [],
  );

describe("bounded table sequence alignment", () => {
  test("keeps two consecutive inserted tables ahead of the original table sequence", () => {
    const base = tableSequence([
      { id: "table-a", text: "Alpha obligations" },
      { id: "table-b", text: "Beta obligations" },
    ]);
    const revised = tableSequence([
      { id: "table-x", text: "Unrelated schedule" },
      { id: "table-y", text: "Unrelated appendix" },
      { id: "table-a", text: "Alpha obligations" },
      { id: "table-b", text: "Beta obligations" },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual([
      "revisedTable",
      "revisedTable",
      "pair",
      "pair",
    ]);
    expect(pairIds(steps)).toEqual([
      ["table-a", "table-a"],
      ["table-b", "table-b"],
    ]);
  });

  test("stable table identity outranks an earlier exact boilerplate signature", () => {
    const repeated = "Standard terms apply to this schedule";
    const base = tableSequence([
      { id: "table-a", text: repeated },
      { id: "table-b", text: repeated },
    ]);
    const revised = tableSequence([
      { id: "inserted-copy", text: repeated },
      { id: "table-a", text: "Standard revised terms apply to this schedule" },
      { id: "table-b", text: repeated },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["revisedTable", "pair", "pair"]);
    expect(pairIds(steps)).toEqual([
      ["table-a", "table-a"],
      ["table-b", "table-b"],
    ]);
  });

  test("exact table signatures outrank earlier similar positional boilerplate", () => {
    const base = tableSequence([
      {
        id: "base-a",
        text: "Payment is due within thirty calendar days",
        idStability: "positional",
      },
      {
        id: "base-b",
        text: "Confidentiality obligations survive termination",
        idStability: "positional",
      },
    ]);
    const revised = tableSequence([
      {
        id: "inserted-similar",
        text: "Payment is due within forty calendar days",
        idStability: "positional",
      },
      {
        id: "revised-a",
        text: "Payment is due within thirty calendar days",
        idStability: "positional",
      },
      {
        id: "revised-b",
        text: "Confidentiality obligations survive termination",
        idStability: "positional",
      },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["revisedTable", "pair", "pair"]);
    expect(pairIds(steps)).toEqual([
      ["base-a", "revised-a"],
      ["base-b", "revised-b"],
    ]);
  });

  test("pairs an edited table only above the similarity confidence floor", () => {
    const base = tableSequence([
      {
        id: "base",
        text: "Payment is due within thirty calendar days",
        idStability: "positional",
      },
    ]);
    const revised = tableSequence([
      {
        id: "revised",
        text: "Payment is due within forty calendar days",
        idStability: "positional",
      },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["pair"]);
    expect(pairIds(steps)).toEqual([["base", "revised"]]);
  });

  test("keeps stable and exact evidence after similarity lookup budget exhaustion", () => {
    const workSession = createFolioContentAlignmentWorkSession({
      structuralTokenLookups: 0,
    });
    const positionalBase = tableSequence([
      {
        id: "positional-base",
        text: "Payment is due within thirty calendar days",
        idStability: "positional",
      },
    ]);
    const similarRevised = tableSequence([
      {
        id: "similar-revised",
        text: "Payment is due within forty calendar days",
        idStability: "positional",
      },
    ]);
    const exactRevised = tableSequence([
      {
        id: "exact-revised",
        text: "Payment is due within thirty calendar days",
        idStability: "positional",
      },
    ]);
    const stableBase = tableSequence([{ id: "stable", text: "Original schedule" }]);
    const stableRevised = tableSequence([{ id: "stable", text: "Revised appendix" }]);

    expect(
      alignFolioContentStructure({
        baseBlocks: positionalBase,
        revisedBlocks: similarRevised,
        workSession,
      }).map(({ type }) => type),
    ).toEqual(["baseTable", "revisedTable"]);
    expect(
      alignFolioContentStructure({
        baseBlocks: positionalBase,
        revisedBlocks: exactRevised,
        workSession,
      }).map(({ type }) => type),
    ).toEqual(["pair"]);
    expect(
      alignFolioContentStructure({
        baseBlocks: stableBase,
        revisedBlocks: stableRevised,
        workSession,
      }).map(({ type }) => type),
    ).toEqual(["pair"]);
    expect(workSession.remainingStructuralTokenLookups).toBe(0);
  });

  test("leaves unrelated table replacements unpaired", () => {
    const base = tableSequence([{ id: "base", text: "Payment schedule" }]);
    const revised = tableSequence([{ id: "revised", text: "Witness addresses" }]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["baseTable", "revisedTable"]);
    expect(pairIds(steps)).toEqual([]);
  });

  test("does not treat one-word boilerplate as table identity", () => {
    const base = [
      tableBlock({
        id: "base-label",
        text: "Standard",
        outerTableIndex: 0,
        paragraphIndex: 0,
        idStability: "positional",
      }),
      tableBlock({
        id: "base-content",
        text: "Payment schedule imposes several separate obligations",
        outerTableIndex: 0,
        paragraphIndex: 1,
        idStability: "positional",
      }),
    ];
    const revised = [
      tableBlock({
        id: "revised-label",
        text: "Standard",
        outerTableIndex: 0,
        paragraphIndex: 0,
        idStability: "positional",
      }),
      tableBlock({
        id: "revised-content",
        text: "Witness addresses identify wholly unrelated factual matters",
        outerTableIndex: 0,
        paragraphIndex: 1,
        idStability: "positional",
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["baseTable", "revisedTable"]);
  });

  test("pairs a sole rewritten table in a gap bounded by preserved body anchors", () => {
    const body = (id: string, text: string): FolioContentBlock => ({
      id,
      kind: "paragraph",
      text,
    });
    const baseTable = tableBlock({ id: "base", text: "Payment schedule", outerTableIndex: 0 });
    const revisedTable = tableBlock({
      id: "revised",
      text: "Witness addresses",
      outerTableIndex: 0,
    });
    const base = [body("before-base", "Before"), baseTable, body("after-base", "After")];
    const revised = [
      body("before-revised", "Before"),
      revisedTable,
      body("after-revised", "After"),
    ];

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
      stableIdMismatch: "pair",
    });

    expect(steps.map(({ type }) => type)).toEqual(["pair", "pair", "pair"]);
    expect(pairIds(steps)).toEqual([
      ["before-base", "before-revised"],
      ["base", "revised"],
      ["after-base", "after-revised"],
    ]);
  });

  test("does not pair a singleton table across a preserved body anchor", () => {
    const body = (id: string): FolioContentBlock => ({
      id,
      kind: "paragraph",
      text: "Preserved body anchor",
    });
    const table = tableBlock({
      id: "relocated-table",
      text: "Relocated schedule",
      outerTableIndex: 0,
    });

    const steps = alignFolioContentStructure({
      baseBlocks: [body("body"), table],
      revisedBlocks: [table, body("body")],
      stableIdMismatch: "pair",
    });

    expect(steps.map(({ type }) => type)).toEqual(["revisedTable", "pair", "baseTable"]);
    expect(pairIds(steps)).toEqual([["body", "body"]]);
  });

  test.each([
    {
      label: "a leading body deletion",
      baseOrder: ["leading", "table", "trailing"],
      revisedOrder: ["table", "trailing"],
      unmatchedType: "baseOnly",
    },
    {
      label: "a leading body insertion",
      baseOrder: ["table", "trailing"],
      revisedOrder: ["leading", "table", "trailing"],
      unmatchedType: "revisedOnly",
    },
  ] as const)("keeps a table paired across $label", ({ baseOrder, revisedOrder, unmatchedType }) => {
    const blocks = {
      leading: { id: "leading", kind: "paragraph", text: "Leading body paragraph" },
      table: tableBlock({ id: "table", text: "Preserved schedule", outerTableIndex: 0 }),
      trailing: { id: "trailing", kind: "paragraph", text: "Trailing body paragraph" },
    } satisfies Record<string, FolioContentBlock>;
    const materialize = (
      order: readonly ("leading" | "table" | "trailing")[],
    ): FolioContentBlock[] => order.map((id) => blocks[id]);

    const steps = alignFolioContentStructure({
      baseBlocks: materialize(baseOrder),
      revisedBlocks: materialize(revisedOrder),
    });

    expect(steps.map(({ type }) => type)).toEqual([unmatchedType, "pair", "pair"]);
    expect(pairIds(steps)).toEqual([
      ["table", "table"],
      ["trailing", "trailing"],
    ]);
  });

  test("reserves a later table identity before aligning an earlier singleton run", () => {
    const body = (id: string): FolioContentBlock => ({
      id,
      kind: "paragraph",
      text: "Preserved body anchor",
    });
    const first = tableBlock({ id: "table-one", text: "First schedule", outerTableIndex: 0 });
    const second = tableBlock({ id: "table-two", text: "Second schedule", outerTableIndex: 1 });
    const revisedSecond = tableBlock({
      id: "table-two",
      text: "Second schedule",
      outerTableIndex: 0,
    });

    const steps = alignFolioContentStructure({
      baseBlocks: [first, body("body"), second],
      revisedBlocks: [revisedSecond, body("body")],
      stableIdMismatch: "pair",
    });

    expect(pairIds(steps)).toEqual([["body", "body"]]);
    expect(steps.map(({ type }) => type)).toEqual([
      "baseTable",
      "revisedTable",
      "pair",
      "baseTable",
    ]);
  });

  test("preserved neighboring tables prevent relocated content from hijacking table mapping", () => {
    const base = [
      tableBlock({
        id: "anchor-a",
        text: "First table anchor",
        outerTableIndex: 0,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "relocated",
        text: "Relocated clause",
        outerTableIndex: 0,
        paragraphIndex: 1,
      }),
      tableBlock({ id: "anchor-b", text: "Second table anchor", outerTableIndex: 1 }),
    ];
    const revised = [
      tableBlock({
        id: "anchor-a",
        text: "First table anchor",
        outerTableIndex: 0,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "replacement",
        text: "Replacement clause",
        outerTableIndex: 0,
        paragraphIndex: 1,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second table anchor",
        outerTableIndex: 1,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "relocated",
        text: "Relocated clause",
        outerTableIndex: 1,
        paragraphIndex: 1,
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    const pairs = steps.filter((step) => step.type === "pair");

    expect(
      pairs.map(({ baseBlock, revisedBlock }) => [
        baseBlock.id,
        revisedBlock.id,
        baseBlock.table?.outerTableIndex,
        revisedBlock.table?.outerTableIndex,
      ]),
    ).toEqual([
      ["anchor-a", "anchor-a", 0, 0],
      ["anchor-b", "anchor-b", 1, 1],
    ]);
  });

  test("one shifted content match cannot define a table mapping with residue on both sides", () => {
    const base = tableSequence([
      { id: "relocated", text: "Relocated substantive clause" },
      { id: "base-only", text: "Base-only schedule" },
    ]);
    const revised = tableSequence([
      { id: "revised-only", text: "Revised-only appendix" },
      { id: "relocated", text: "Relocated substantive clause" },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(pairIds(steps)).toEqual([]);
  });

  test("falls back conservatively before a table matrix exceeds the shared budget", () => {
    const workSession = createFolioContentAlignmentWorkSession({ lcsCells: 7 });
    const base = tableSequence([
      { id: "table-a", text: "Alpha obligations" },
      { id: "table-b", text: "Beta obligations" },
    ]);
    const revised = tableSequence([
      { id: "table-x", text: "Unrelated schedule" },
      { id: "table-y", text: "Unrelated appendix" },
      { id: "table-a", text: "Alpha obligations" },
      { id: "table-b", text: "Beta obligations" },
    ]);

    const steps = alignFolioContentStructure({
      baseBlocks: base,
      revisedBlocks: revised,
      workSession,
    });

    expect(steps.map(({ type }) => type)).toEqual([
      "baseTable",
      "baseTable",
      "revisedTable",
      "revisedTable",
      "revisedTable",
      "revisedTable",
    ]);
    expect(workSession.remainingLcsCells).toBe(7);
  });

  test("does not infer a table match from a truncated similarity profile", () => {
    const sharedPrefix = "standard boilerplate ".repeat(1_000);
    const base = tableSequence([
      { id: "base", text: `${sharedPrefix}base ending`, idStability: "positional" },
    ]);
    const revised = tableSequence([
      { id: "revised", text: `${sharedPrefix}revised ending`, idStability: "positional" },
    ]);

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual(["baseTable", "revisedTable"]);
  });

  test("reports an unpaired nested-only segment as its outer table", () => {
    const nested = tableBlock({
      id: "nested",
      text: "Nested-only content",
      outerTableIndex: 0,
      tableIndex: 1,
    });

    const steps = alignFolioContentStructure({ baseBlocks: [nested], revisedBlocks: [] });

    expect(steps).toEqual([
      {
        type: "baseTable",
        blocks: [nested],
        location: { ...nested.table, tableIndex: 0 },
      },
    ]);
  });
});

describe("bounded table row sequence alignment", () => {
  test("keeps two consecutive inserted rows ahead of the original row sequence", () => {
    const base = [
      tableBlock({ id: "row-a", text: "Alpha obligations", outerTableIndex: 0, rowIndex: 0 }),
      tableBlock({ id: "row-b", text: "Beta obligations", outerTableIndex: 0, rowIndex: 1 }),
    ];
    const revised = [
      tableBlock({ id: "row-x", text: "Unrelated schedule", outerTableIndex: 0, rowIndex: 0 }),
      tableBlock({ id: "row-y", text: "Unrelated appendix", outerTableIndex: 0, rowIndex: 1 }),
      tableBlock({ id: "row-a", text: "Alpha obligations", outerTableIndex: 0, rowIndex: 2 }),
      tableBlock({ id: "row-b", text: "Beta obligations", outerTableIndex: 0, rowIndex: 3 }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(steps.map(({ type }) => type)).toEqual([
      "revisedRow",
      "revisedRow",
      "pair",
      "pair",
    ]);
    expect(pairIds(steps)).toEqual([
      ["row-a", "row-a"],
      ["row-b", "row-b"],
    ]);
  });

  test("preserved neighboring rows prevent relocated content from hijacking row mapping", () => {
    const base = [
      tableBlock({
        id: "anchor-a",
        text: "First row anchor",
        outerTableIndex: 0,
        rowIndex: 0,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "relocated",
        text: "Relocated clause",
        outerTableIndex: 0,
        rowIndex: 0,
        paragraphIndex: 1,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second row anchor",
        outerTableIndex: 0,
        rowIndex: 1,
      }),
    ];
    const revised = [
      tableBlock({
        id: "anchor-a",
        text: "First row anchor",
        outerTableIndex: 0,
        rowIndex: 0,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "replacement",
        text: "Replacement clause",
        outerTableIndex: 0,
        rowIndex: 0,
        paragraphIndex: 1,
      }),
      tableBlock({
        id: "anchor-b",
        text: "Second row anchor",
        outerTableIndex: 0,
        rowIndex: 1,
        paragraphIndex: 0,
      }),
      tableBlock({
        id: "relocated",
        text: "Relocated clause",
        outerTableIndex: 0,
        rowIndex: 1,
        paragraphIndex: 1,
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });
    const pairs = steps.filter((step) => step.type === "pair");

    expect(
      pairs.map(({ baseBlock, revisedBlock }) => [
        baseBlock.id,
        revisedBlock.id,
        baseBlock.table?.rowIndex,
        revisedBlock.table?.rowIndex,
      ]),
    ).toEqual([
      ["anchor-a", "anchor-a", 0, 0],
      ["anchor-b", "anchor-b", 1, 1],
    ]);
  });

  test("one shifted content match cannot define a row mapping with residue on both sides", () => {
    const base = [
      tableBlock({
        id: "relocated",
        text: "Relocated substantive clause",
        outerTableIndex: 0,
        rowIndex: 0,
      }),
      tableBlock({
        id: "base-only",
        text: "Base-only row",
        outerTableIndex: 0,
        rowIndex: 1,
      }),
    ];
    const revised = [
      tableBlock({
        id: "revised-only",
        text: "Revised-only row",
        outerTableIndex: 0,
        rowIndex: 0,
      }),
      tableBlock({
        id: "relocated",
        text: "Relocated substantive clause",
        outerTableIndex: 0,
        rowIndex: 1,
      }),
    ];

    const steps = alignFolioContentStructure({ baseBlocks: base, revisedBlocks: revised });

    expect(pairIds(steps)).toEqual([]);
  });
});
