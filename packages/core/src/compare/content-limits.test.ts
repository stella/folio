import { describe, expect, test } from "bun:test";

import {
  compareContent,
  createContentComparisonWorkSession,
  FOLIO_CONTENT_COMPARISON_LIMITS,
  FolioContentComparisonLimitError,
  InvalidFolioContentComparisonError,
} from "./content";
import type {
  FolioContentIdentitySemantics,
  FolioContentInputBlock,
  FolioContentInputRun,
  FolioContentPropertyInput,
  FolioContentPropertyInputValue,
  FolioContentTableLocation,
} from "./content-types";

type TableInput = Omit<
  FolioContentTableLocation,
  "outerTableIdentity" | "tableIdentity" | "rowIdentity" | "cellIdentity"
>;

type TestBlockOverrides = {
  text?: string;
  identitySemantics?: FolioContentIdentitySemantics;
  runs?: readonly FolioContentInputRun[];
  styleId?: string;
  directSpacing?: Readonly<Record<string, FolioContentPropertyInputValue>>;
  table?: TableInput;
  containerPath?: FolioContentInputBlock["containerPath"];
};

const propertySet = (
  entries: Readonly<Record<string, FolioContentPropertyInputValue | undefined>>,
): FolioContentPropertyInput =>
  Object.entries(entries)
    .filter((entry): entry is [string, FolioContentPropertyInputValue] => entry[1] !== undefined)
    .map(([key, value]) => ({ key, value }));

const block = (
  id: string,
  {
    text = "",
    identitySemantics = "authoritative",
    runs,
    styleId,
    directSpacing,
    table,
    containerPath,
  }: TestBlockOverrides = {},
): FolioContentInputBlock => ({
  identity: { type: identitySemantics, id },
  kind: "paragraph",
  text,
  ...(runs !== undefined && { runs }),
  ...((styleId !== undefined || directSpacing !== undefined) && {
    paragraphFormatting: propertySet({
      styleId,
      directSpacing:
        directSpacing === undefined
          ? undefined
          : { type: "object", entries: propertySet(directSpacing) },
    }),
  }),
  ...(table !== undefined && {
    table: {
      outerTableIdentity: {
        type: "positional",
        id: `outer-table-${String(table.outerTableIndex)}`,
      },
      tableIdentity: { type: "positional", id: `table-${String(table.tableIndex)}` },
      rowIdentity: {
        type: "positional",
        id: `table-${String(table.tableIndex)}-row-${String(table.rowIndex)}`,
      },
      cellIdentity: {
        type: "positional",
        id: `table-${String(table.tableIndex)}-row-${String(table.rowIndex)}-cell-${String(table.cellIndex)}`,
      },
      ...table,
    },
  }),
  ...(containerPath !== undefined && { containerPath }),
});

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

const richComparisonInput = () => {
  const richBlock = block("rich", {
    text: "A",
    runs: [
      {
        text: "A",
        effectiveFormatting: [{ key: "bold", value: true }],
        authoredFormatting: [{ key: "color", value: "112233" }],
      },
    ],
    directSpacing: { spaceAfter: 120 },
    containerPath: [
      {
        kind: "section",
        identity: { type: "authoritative", id: "schedule" },
      },
    ],
    table: {
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    },
  });
  Reflect.set(richBlock, "structuralBoundaries", [
    { type: "pageBreak", offset: 0, clear: "all" },
  ]);
  Reflect.set(richBlock, "blockProperties", [
    {
      key: "arrayProperty",
      value: { type: "array", items: ["value"] },
    },
    {
      key: "objectProperty",
      value: {
        type: "object",
        entries: [{ key: "nested", value: 1 }],
      },
    },
  ]);
  return {
    base: { blocks: [richBlock] },
    revised: { blocks: [] },
    granularity: "word" as const,
  };
};

const objectAtPath = (root: unknown, path: readonly PropertyKey[]): object => {
  let current = root;
  for (const key of path) {
    if ((typeof current !== "object" && typeof current !== "function") || current === null) {
      throw new Error(`Fixture path ${path.map(String).join(".")} is not an object.`);
    }
    current = Reflect.get(current, key);
  }
  if ((typeof current !== "object" && typeof current !== "function") || current === null) {
    throw new Error(`Fixture path ${path.map(String).join(".")} is not an object.`);
  }
  return current;
};

const expectInvalidComparisonInput = (input: ReturnType<typeof richComparisonInput>): void => {
  const result = compareContent(input);
  expect(result.isErr()).toBe(true);
  if (!result.isErr()) return;
  expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
};

describe("neutral comparison resource boundaries", () => {
  test("charges aggregate input usage across captured story pairs atomically", () => {
    const text = "x".repeat(1_000_000);
    const story = (prefix: string, count: number) => ({
      blocks: Array.from({ length: count }, (_unused, index) =>
        block(`${prefix}-${String(index)}`, { text }),
      ),
    });
    for (const side of ["base", "revised"] as const) {
      const workSession = createContentComparisonWorkSession();
      const empty = { blocks: [] } as const;
      const firstStory = story(`${side}-first`, 4);
      const secondStory = story(`${side}-second`, 5);
      const first = workSession.captureComparison({
        base: side === "base" ? firstStory : empty,
        revised: side === "revised" ? firstStory : empty,
      });
      expect(first.isOk()).toBe(true);
      if (first.isErr()) continue;
      expect(first.value.compare().isOk()).toBe(true);

      const second = workSession.captureComparison({
        base: side === "base" ? secondStory : empty,
        revised: side === "revised" ? secondStory : empty,
      });
      expect(second.isErr()).toBe(true);
      if (!second.isErr()) continue;
      expect(second.error).toBeInstanceOf(FolioContentComparisonLimitError);
      expect(second.error).toMatchObject({
        input: side,
        limit: "textCodeUnitsPerSnapshot",
        maximum: FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot,
        actual: 9_000_000,
      });
      const exactRemainder = story(`${side}-remainder`, 4);
      const remainder = workSession.captureComparison({
        base: side === "base" ? exactRemainder : empty,
        revised: side === "revised" ? exactRemainder : empty,
      });
      expect(remainder.isOk()).toBe(true);
      if (remainder.isOk()) expect(remainder.value.compare().isOk()).toBe(true);
    }

    const workSession = createContentComparisonWorkSession();
    const firstPair = workSession.captureComparison({
      base: story("atomic-base-first", 1),
      revised: story("atomic-revised-first", 4),
    });
    expect(firstPair.isOk()).toBe(true);
    if (firstPair.isErr()) return;
    expect(firstPair.value.compare().isOk()).toBe(true);
    const rejectedPair = workSession.captureComparison({
      base: story("atomic-base-second", 2),
      revised: story("atomic-revised-second", 5),
    });
    expect(rejectedPair.isErr()).toBe(true);
    if (rejectedPair.isErr()) {
      expect(rejectedPair.error).toMatchObject({
        input: "revised",
        limit: "textCodeUnitsPerSnapshot",
        actual: 9_000_000,
      });
    }
    const baseRemainder = workSession.captureComparison({
      base: story("atomic-base-remainder", 7),
      revised: { blocks: [] },
    });
    expect(baseRemainder.isOk()).toBe(true);
    if (baseRemainder.isOk()) expect(baseRemainder.value.compare().isOk()).toBe(true);
  });

  test("charges one aggregate change budget and poisons a failed session", () => {
    const workSession = createContentComparisonWorkSession();
    const compareInsertedStory = (prefix: string, count: number) => {
      const operation = workSession.captureComparison({
        base: { blocks: [] },
        revised: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`${prefix}-${String(index)}`),
          ),
        },
      });
      if (operation.isErr()) {
        throw operation.error;
      }
      return operation.value.compare();
    };

    expect(compareInsertedStory("first", 6_000).isOk()).toBe(true);
    const second = compareInsertedStory("second", 4_001);
    expect(second.isErr()).toBe(true);
    if (!second.isErr()) return;
    expect(second.error).toMatchObject({
      input: "result",
      limit: "changes",
      maximum: FOLIO_CONTENT_COMPARISON_LIMITS.changes,
      actual: FOLIO_CONTENT_COMPARISON_LIMITS.changes + 1,
    });
    const afterFailure = workSession.captureComparison({
      base: { blocks: [] },
      revised: { blocks: [] },
    });
    expect(afterFailure.error).toMatchObject({ reason: "session-poisoned" });
  });

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
      Math.floor(FOLIO_CONTENT_COMPARISON_LIMITS.textCodeUnitsPerSnapshot / text.length) + 1;

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
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.runsPerBlock + 1 },
      () => emptyRun,
    );
    expectLimit(
      compareContent({
        base: { blocks: [block("base", { runs: tooManyInBlock })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "runsPerBlock", blockIndex: 0 },
    );

    const maximumRuns = Array.from(
      { length: FOLIO_CONTENT_COMPARISON_LIMITS.runsPerBlock },
      () => emptyRun,
    );
    const count =
      Math.floor(FOLIO_CONTENT_COMPARISON_LIMITS.runsPerSnapshot / maximumRuns.length) + 1;
    expectLimit(
      compareContent({
        base: {
          blocks: Array.from({ length: count }, (_unused, index) =>
            block(`base-${String(index)}`, { runs: maximumRuns }),
          ),
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "runsPerSnapshot", blockIndex: count - 1 },
    );
  });

  test("bounds container depth and aggregate ancestry entries", () => {
    const entry = {
      kind: "c",
      identity: { type: "authoritative", id: "i" },
    } as const;
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
        FOLIO_CONTENT_COMPARISON_LIMITS.containerEntriesPerSnapshot / maximumDepth.length,
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
          blocks: [block("x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits + 1))],
        },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "attributeCodeUnits", blockIndex: 0 },
    );

    const styleId = "x".repeat(FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnits);
    const count =
      Math.floor(FOLIO_CONTENT_COMPARISON_LIMITS.attributeCodeUnitsPerSnapshot / styleId.length) +
      1;
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
      field: "blocks[0].table",
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
        blocks: [block("table-before", { table }), block("body"), block("table-after", { table })],
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

  test("rejects oversized arrays before traversing their elements", () => {
    const oversizedBlocks = new Array<FolioContentBlock>(
      FOLIO_CONTENT_COMPARISON_LIMITS.blocksPerSnapshot + 1,
    );
    Object.defineProperty(oversizedBlocks, 0, {
      get: () => {
        throw new Error("oversized blocks must not be traversed");
      },
    });
    expectLimit(
      compareContent({ base: { blocks: oversizedBlocks }, revised: { blocks: [] } }),
      { input: "base", limit: "blocksPerSnapshot" },
    );

    const oversizedRuns = new Array<FolioContentInputRun>(
      FOLIO_CONTENT_COMPARISON_LIMITS.runsPerBlock + 1,
    );
    Object.defineProperty(oversizedRuns, 0, {
      get: () => {
        throw new Error("oversized runs must not be traversed");
      },
    });
    expectLimit(
      compareContent({
        base: { blocks: [block("runs", { runs: oversizedRuns })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "runsPerBlock", blockIndex: 0 },
    );

    const oversizedPath = new Array<NonNullable<FolioContentInputBlock["containerPath"]>[number]>(
      FOLIO_CONTENT_COMPARISON_LIMITS.containerDepth + 1,
    );
    Object.defineProperty(oversizedPath, 0, {
      get: () => {
        throw new Error("oversized container paths must not be traversed");
      },
    });
    expectLimit(
      compareContent({
        base: { blocks: [block("container", { containerPath: oversizedPath })] },
        revised: { blocks: [] },
      }),
      { input: "base", limit: "containerDepth", blockIndex: 0 },
    );
  });

  test("capture owns semantics and can be consumed only once", () => {
    const directSpacing = { spaceAfter: 120 };
    const base = block("clause", {
      text: "Before",
      directSpacing,
      runs: [
        {
          text: "Before",
          effectiveFormatting: [{ key: "bold", value: true }],
        },
      ],
    });
    const revised = block("clause", { text: "After" });
    const session = createContentComparisonWorkSession();
    const operation = session.captureComparison({
      base: { blocks: [base] },
      revised: { blocks: [revised] },
    });
    if (operation.isErr()) throw operation.error;
    const overlappingCapture = session.captureComparison({
      base: { blocks: [] },
      revised: { blocks: [] },
    });
    expect(overlappingCapture.error).toMatchObject({ reason: "operation-active" });

    Reflect.set(base, "text", "Mutated base");
    Reflect.set(revised, "text", "Mutated revised");
    directSpacing.spaceAfter = 999;
    const compared = operation.value.compare();
    if (compared.isErr()) throw compared.error;
    const event = compared.value.events.at(0);
    if (event?.type !== "modified") throw new Error("expected a modified event");
    expect(event.relation.base.block.text).toBe("Before");
    expect(event.relation.revised.block.text).toBe("After");
    expect(event.relation.base.block.paragraphFormatting).toEqual([
      {
        key: "directSpacing",
        value: {
          type: "object",
          entries: [{ key: "spaceAfter", value: 120 }],
        },
      },
    ]);
    expect(operation.value.compare().error).toMatchObject({ reason: "operation-consumed" });
  });

  test("accepts readonly snapshots without copying caller arrays", () => {
    const base = {
      blocks: [block("same", { text: "Same" })],
    } as const;

    const result = compareContent({ base, revised: base });

    expect(result.isOk()).toBe(true);
  });

  test("rejects misspelled, hidden, symbol, and accessor fields at every record boundary", () => {
    const recordPaths = [
      [] as const,
      ["base"] as const,
      ["base", "blocks", 0] as const,
      ["base", "blocks", 0, "identity"] as const,
      ["base", "blocks", 0, "table"] as const,
      ["base", "blocks", 0, "containerPath", 0] as const,
      ["base", "blocks", 0, "runs", 0] as const,
      ["base", "blocks", 0, "structuralBoundaries", 0] as const,
      ["base", "blocks", 0, "blockProperties", 0] as const,
      ["base", "blocks", 0, "blockProperties", 0, "value"] as const,
      ["base", "blocks", 0, "blockProperties", 1, "value"] as const,
      ["base", "blocks", 0, "blockProperties", 1, "value", "entries", 0] as const,
    ];

    for (const path of recordPaths) {
      const fields = Reflect.ownKeys(objectAtPath(richComparisonInput(), path)).filter(
        (key): key is string => typeof key === "string",
      );
      for (const field of fields) {
        const input = richComparisonInput();
        const target = objectAtPath(input, path);
        const descriptor = Object.getOwnPropertyDescriptor(target, field);
        if (!descriptor) throw new Error(`Fixture field ${field} has no descriptor.`);
        Reflect.deleteProperty(target, field);
        Object.defineProperty(target, `${field}Typo`, descriptor);
        expectInvalidComparisonInput(input);
      }

      for (const extra of ["hidden", "symbol"] as const) {
        const input = richComparisonInput();
        const target = objectAtPath(input, path);
        if (extra === "hidden") {
          Object.defineProperty(target, "unexpected", { value: true, enumerable: false });
        } else {
          Object.defineProperty(target, Symbol("unexpected"), { value: true });
        }
        expectInvalidComparisonInput(input);
      }

      const input = richComparisonInput();
      const target = objectAtPath(input, path);
      const field = fields.at(0);
      if (!field) throw new Error("A record-boundary fixture must contain one field.");
      let getterCalls = 0;
      Object.defineProperty(target, field, {
        configurable: true,
        enumerable: true,
        get: () => {
          getterCalls++;
          throw new Error("comparison validation must not invoke accessors");
        },
      });
      expectInvalidComparisonInput(input);
      expect(getterCalls).toBe(0);
    }
  });

  test("rejects holes, accessors, symbols, and non-index fields on every input array", () => {
    const arrayPaths = [
      ["base", "blocks"] as const,
      ["base", "blocks", 0, "runs"] as const,
      ["base", "blocks", 0, "containerPath"] as const,
      ["base", "blocks", 0, "structuralBoundaries"] as const,
      ["base", "blocks", 0, "blockProperties"] as const,
      ["base", "blocks", 0, "blockProperties", 0, "value", "items"] as const,
      ["base", "blocks", 0, "blockProperties", 1, "value", "entries"] as const,
    ];

    for (const path of arrayPaths) {
      for (const mutation of ["hole", "accessor", "symbol", "non-index"] as const) {
        const input = richComparisonInput();
        const target = objectAtPath(input, path);
        let getterCalls = 0;
        switch (mutation) {
          case "hole":
            Reflect.deleteProperty(target, "0");
            break;
          case "accessor":
            Object.defineProperty(target, "0", {
              configurable: true,
              enumerable: true,
              get: () => {
                getterCalls++;
                throw new Error("comparison validation must not invoke array accessors");
              },
            });
            break;
          case "symbol":
            Object.defineProperty(target, Symbol("unexpected"), { value: true });
            break;
          case "non-index":
            Object.defineProperty(target, "unexpected", { value: true });
            break;
          default: {
            const unreachable: never = mutation;
            throw new Error(`Unhandled mutation ${unreachable}`);
          }
        }
        expectInvalidComparisonInput(input);
        expect(getterCalls).toBe(0);
      }
    }
  });

  test("canonicalizes property ordering and negative zero without losing explicit null", () => {
    const baseBlock = block("properties", { text: "Same" });
    Reflect.set(baseBlock, "blockProperties", [
      { key: "z", value: -0 },
      { key: "a", value: null },
    ]);
    const result = compareContent({
      base: { blocks: [baseBlock] },
      revised: { blocks: [] },
    });
    if (result.isErr()) throw result.error;
    const event = result.value.events.at(0);
    if (event?.type !== "deleted") throw new Error("expected a deleted block");
    expect(event.block.blockProperties).toEqual([
      { key: "a", value: null },
      { key: "z", value: 0 },
    ]);
    expect(Object.is(event.block.blockProperties[1]?.value, -0)).toBe(false);
  });
});

describe("edited-move work accounting", () => {
  test("charges token lookups as well as candidate pairs", () => {
    const base = [
      block("base", {
        text: "alpha beta gamma delta epsilon",
        identitySemantics: "persistent-hint",
      }),
      block("anchor-1", { text: "first durable anchor" }),
      block("anchor-2", { text: "second durable anchor" }),
      block("anchor-3", { text: "third durable anchor" }),
    ];
    const revised = [
      block("anchor-1", { text: "first durable anchor" }),
      block("anchor-2", { text: "second durable anchor" }),
      block("anchor-3", { text: "third durable anchor" }),
      block("revised", {
        text: "alpha beta gamma delta zeta",
        identitySemantics: "persistent-hint",
      }),
    ];
    const compareWithAllowance = (moveTokenLookups: number) => {
      const session = createContentComparisonWorkSession({
        moveComparisons: 1,
        moveTokenLookups,
      });
      const operation = session.captureComparison({
        base: { blocks: base },
        revised: { blocks: revised },
      });
      if (operation.isErr()) throw operation.error;
      const result = operation.value.compare();
      if (result.isErr()) throw result.error;
      return result.value.events.map(({ type }) => type);
    };

    expect(compareWithAllowance(4)).toEqual([
      "deleted",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
    ]);
    expect(compareWithAllowance(5)).toEqual([
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "movedTo",
    ]);
  });
});
