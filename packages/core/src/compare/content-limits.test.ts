import { describe, expect, test } from "bun:test";

import {
  compareContent,
  createContentComparisonWorkSession,
  detectFolioContentMoves,
  FOLIO_CONTENT_COMPARISON_LIMITS,
  FolioContentComparisonLimitError,
  InvalidFolioContentComparisonError,
} from "./content";
import type { FolioContentAlignmentStep } from "./content-alignment";
import type { FolioContentBlock } from "./content-types";

const block = (
  id: string,
  overrides: Partial<FolioContentBlock> = {},
): FolioContentBlock => ({ id, kind: "paragraph", text: "", ...overrides });

const expectLimit = (
  result: ReturnType<typeof compareContent>,
  expected: {
    input: "base" | "revised" | "result";
    limit: keyof typeof FOLIO_CONTENT_COMPARISON_LIMITS;
    blockIndex?: number;
  },
): void => {
  expect(result.isErr()).toBe(true);
  if (!result.isErr()) return;
  expect(result.error).toBeInstanceOf(FolioContentComparisonLimitError);
  expect(result.error).toMatchObject(expected);
};

describe("neutral comparison resource boundaries", () => {
  test("rejects a block before tokenizing more text than one diff may retain", () => {
    const text = "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits + 1);

    expectLimit(
      compareContent({
        base: { blocks: [block("base", { text })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "blockCodeUnits", blockIndex: 0 },
    );
  });

  test("shares one cumulative text allowance across every block on a side", () => {
    const text = "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.blockCodeUnits);
    const count =
      Math.floor(FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot / text.length) +
      1;

    expectLimit(
      compareContent({
        base: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`base-${String(index)}`, { text }),
          ),
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "textCodeUnitsPerSnapshot", blockIndex: count - 1 },
    );
  });

  test("bounds empty preview runs per block and across a snapshot", () => {
    const emptyRun = { text: "" } as const;
    const tooManyInBlock = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerBlock + 1 },
      () => emptyRun,
    );
    expectLimit(
      compareContent({
        base: { blocks: [block("base", { previewRuns: tooManyInBlock })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "previewRunsPerBlock", blockIndex: 0 },
    );

    const maximumRuns = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerBlock },
      () => emptyRun,
    );
    const count =
      Math.floor(
        FOLIO_CONTENT_COMPARISON_LIMITS.previewRunsPerSnapshot /
          maximumRuns.length,
      ) + 1;
    expectLimit(
      compareContent({
        base: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`base-${String(index)}`, { previewRuns: maximumRuns }),
          ),
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "previewRunsPerSnapshot", blockIndex: count - 1 },
    );
  });

  test("bounds container depth and aggregate ancestry entries", () => {
    const entry = { kind: "c", id: "i" } as const;
    const tooDeep = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth + 1 },
      () => entry,
    );
    expectLimit(
      compareContent({
        base: { blocks: [block("base", { containerPath: tooDeep })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "containerDepth", blockIndex: 0 },
    );

    const maximumDepth = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth },
      () => entry,
    );
    const count =
      Math.floor(
        FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot /
          maximumDepth.length,
      ) + 1;
    expectLimit(
      compareContent({
        base: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`base-${String(index)}`, { containerPath: maximumDepth }),
          ),
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "containerEntriesPerSnapshot", blockIndex: count - 1 },
    );
  });

  test("bounds individual and aggregate comparison attributes", () => {
    expectLimit(
      compareContent({
        base: {
          blocks: [
            block(
              "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits + 1),
            ),
          ],
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "attributeCodeUnits", blockIndex: 0 },
    );

    const styleId = "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits);
    const count =
      Math.floor(
        FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnitsPerSnapshot /
          styleId.length,
      ) + 1;
    expectLimit(
      compareContent({
        base: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`base-${String(index)}`, { styleId }),
          ),
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "attributeCodeUnitsPerSnapshot" },
    );
  });

  test("rejects unsafe integer coordinates instead of aliasing table positions", () => {
    const result = compareContent({
      base: {
        blocks: [
          block("base", {
            table: {
              outerTableIndex: Number.MAX_SAFE_INTEGER + 1,
              tableIndex: 0,
              rowIndex: 0,
              cellIndex: 0,
              gridColumnIndex: 0,
              columnSpan: 1,
              rowSpan: 1,
              paragraphIndex: 0,
            },
          }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) return;
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      field: "blocks[0].table.outerTableIndex",
    });

    const unsafeExtent = compareContent({
      base: {
        blocks: [
          block("base", {
            table: {
              outerTableIndex: 0,
              tableIndex: 0,
              rowIndex: Number.MAX_SAFE_INTEGER,
              cellIndex: 0,
              gridColumnIndex: 0,
              columnSpan: 1,
              rowSpan: 1,
              paragraphIndex: 0,
            },
          }),
        ],
      },
      revised: { blocks: [] },
    });
    expect(unsafeExtent.isErr()).toBe(true);
    if (!unsafeExtent.isErr()) return;
    expect(unsafeExtent.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(unsafeExtent.error).toMatchObject({
      input: "base",
      field: "blocks[0].table.rowSpan",
    });
  });

  test("rejects an outer table that reappears after its document position closes", () => {
    const table = {
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    } as const;
    const result = compareContent({
      base: {
        blocks: [
          block("table-before", { table }),
          block("body"),
          block("table-after", { table }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) return;
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 2,
      field: "blocks[2].table",
    });
  });

  test("rejects a table index reused under another outer table", () => {
    const location = (outerTableIndex: number) => ({
      outerTableIndex,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    });
    const result = compareContent({
      base: {
        blocks: [
          block("nested-first", { table: location(0) }),
          block("outer-second", { table: location(1) }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) return;
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      input: "base",
      blockIndex: 1,
      field: "blocks[1].table.tableIndex",
    });

    const precedingOuter = compareContent({
      base: {
        blocks: [
          block("impossible-nesting", {
            table: { ...location(1), tableIndex: 0 },
          }),
        ],
      },
      revised: { blocks: [] },
    });
    expect(precedingOuter.isErr()).toBe(true);
    if (!precedingOuter.isErr()) return;
    expect(precedingOuter.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(precedingOuter.error).toMatchObject({
      input: "base",
      blockIndex: 0,
      field: "blocks[0].table.tableIndex",
    });
  });

  test("rejects overlapping horizontal and vertical table cells", () => {
    const cell = ({
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan = 1,
      rowSpan = 1,
    }: {
      rowIndex: number;
      cellIndex: number;
      gridColumnIndex: number;
      columnSpan?: number;
      rowSpan?: number;
    }) => ({
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan,
      rowSpan,
      paragraphIndex: 0,
    });
    for (const blocks of [
      [
        block("wide", {
          table: cell({ rowIndex: 0, cellIndex: 0, gridColumnIndex: 0, columnSpan: 2 }),
        }),
        block("inside-wide", {
          table: cell({ rowIndex: 0, cellIndex: 1, gridColumnIndex: 1 }),
        }),
      ],
      [
        block("tall", {
          table: cell({ rowIndex: 0, cellIndex: 0, gridColumnIndex: 0, rowSpan: 2 }),
        }),
        block("inside-tall", {
          table: cell({ rowIndex: 1, cellIndex: 0, gridColumnIndex: 0 }),
        }),
      ],
    ]) {
      const result = compareContent({ base: { blocks }, revised: { blocks: [] } });
      expect(result.isErr()).toBe(true);
      if (!result.isErr()) continue;
      expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
      expect(result.error).toMatchObject({
        input: "base",
        blockIndex: 1,
        field: "blocks[1].table",
      });
    }
  });

  test("accepts readonly snapshots without copying caller arrays", () => {
    const base = {
      blocks: [block("same", { text: "Same" })],
    } as const;

    const result = compareContent({ base, revised: base });

    expect(result.isOk()).toBe(true);
  });
});

describe("edited-move work accounting", () => {
  test("charges token lookups as well as candidate pairs", () => {
    const base = block("base", { text: "alpha beta gamma delta epsilon" });
    const revised = block("revised", { text: "alpha beta gamma delta zeta" });
    const steps = [
      { type: "baseOnly", block: base, moveScope: { bucket: 0, gap: 0 } },
      { type: "revisedOnly", block: revised, moveScope: { bucket: 0, gap: 1 } },
    ] as const satisfies readonly FolioContentAlignmentStep[];
    const workSession = createContentComparisonWorkSession();
    workSession.remainingMoveComparisons = 1;
    workSession.remainingMoveTokenLookups = 4;

    expect(
      detectFolioContentMoves({
        steps,
        consumedStepIndexes: new Set(),
        workSession,
        idStability: () => "positional",
      }),
    ).toEqual([]);
    expect(workSession.remainingMoveComparisons).toBe(0);
    expect(workSession.remainingMoveTokenLookups).toBe(4);

    const allowedSession = createContentComparisonWorkSession();
    allowedSession.remainingMoveComparisons = 1;
    allowedSession.remainingMoveTokenLookups = 5;
    expect(
      detectFolioContentMoves({
        steps,
        consumedStepIndexes: new Set(),
        workSession: allowedSession,
        idStability: () => "positional",
      }),
    ).toEqual([{ baseBlock: base, revisedBlock: revised }]);
    expect(allowedSession.remainingMoveTokenLookups).toBe(0);
  });
});
