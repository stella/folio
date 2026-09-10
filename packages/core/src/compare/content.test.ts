import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import {
  compareContent,
  InvalidFolioContentComparisonError,
  type FolioContentComparison,
  type FolioContentComparisonEvent,
  type FolioContentTextSegment,
} from "./content";
import type { FolioContentBlock, FolioContentSnapshot } from "./content-types";

type TestBlockKind = "heading" | "paragraph";
type TestBlock = FolioContentBlock<TestBlockKind>;

type TestBlockOptions = Omit<Partial<TestBlock>, "id" | "kind" | "text"> & {
  id: string;
  text: string;
  kind?: TestBlockKind;
};

const contentBlock = ({
  id,
  text,
  kind = "paragraph",
  ...properties
}: TestBlockOptions): TestBlock => ({ id, kind, text, ...properties });

type TableBlockOptions = {
  id: string;
  text: string;
  rowIndex: number;
  cellIndex: number;
  gridColumnIndex?: number;
  tableIndex?: number;
  columnSpan?: number;
  rowSpan?: number;
};

const tableBlock = ({
  id,
  text,
  rowIndex,
  cellIndex,
  gridColumnIndex = cellIndex,
  tableIndex = 0,
  columnSpan = 1,
  rowSpan = 1,
}: TableBlockOptions): TestBlock =>
  contentBlock({
    id,
    text,
    table: {
      outerTableIndex: tableIndex,
      tableIndex,
      rowIndex,
      cellIndex,
      gridColumnIndex,
      columnSpan,
      rowSpan,
      paragraphIndex: 0,
    },
  });

type SuccessfulComparisonOptions = {
  base: TestBlock[];
  revised: TestBlock[];
  granularity?: "word" | "character";
};

const successfulComparison = ({
  base,
  revised,
  granularity,
}: SuccessfulComparisonOptions): FolioContentComparison<TestBlock> => {
  const result = compareContent({
    base: { blocks: base },
    revised: { blocks: revised },
    ...(granularity && { granularity }),
  });
  expect(result.isErr()).toBe(false);
  if (result.isErr()) {
    throw result.error;
  }
  return result.value;
};

const eventTypes = (
  comparison: FolioContentComparison<TestBlock>,
): FolioContentComparisonEvent<TestBlock>["type"][] =>
  comparison.events.map(({ type }) => type);

const baseProjection = (
  comparison: FolioContentComparison<TestBlock>,
): TestBlock[] => comparison.events.flatMap(({ baseBlocks }) => [...baseBlocks]);

const revisedProjection = (
  comparison: FolioContentComparison<TestBlock>,
): TestBlock[] => comparison.events.flatMap(({ revisedBlocks }) => [...revisedBlocks]);

const textBefore = (segments: readonly FolioContentTextSegment[]): string =>
  segments
    .filter(({ type }) => type !== "ins")
    .map(({ text }) => text)
    .join("");

const textAfter = (segments: readonly FolioContentTextSegment[]): string =>
  segments
    .filter(({ type }) => type !== "del")
    .map(({ text }) => text)
    .join("");

const stableAnchors = (): { base: TestBlock[]; revised: TestBlock[] } => ({
  base: [
    contentBlock({ id: "anchor-a", text: "First durable anchor text" }),
    contentBlock({ id: "anchor-b", text: "Second durable anchor text" }),
    contentBlock({ id: "anchor-c", text: "Third durable anchor text" }),
  ],
  revised: [
    contentBlock({ id: "anchor-a", text: "First durable anchor text" }),
    contentBlock({ id: "anchor-b", text: "Second durable anchor text" }),
    contentBlock({ id: "anchor-c", text: "Third durable anchor text" }),
  ],
});

describe("representation-neutral comparison stream", () => {
  test("equal content and presentation produce only unchanged events", () => {
    const base = [
      contentBlock({
        id: "heading",
        kind: "heading",
        text: "Terms",
        headingLevel: 1,
        styleId: "Heading1",
        previewRuns: [{ text: "Terms", bold: true }],
      }),
      contentBlock({
        id: "body",
        text: "Payment is due.",
        directAlignment: "center",
        directSpacing: { spaceAfter: 120 },
      }),
    ];
    const revised = [
      contentBlock({
        id: "heading",
        kind: "heading",
        text: "Terms",
        headingLevel: 1,
        styleId: "Heading1",
        previewRuns: [{ text: "Terms", bold: true }],
      }),
      contentBlock({
        id: "body",
        text: "Payment is due.",
        directAlignment: "center",
        directSpacing: { spaceAfter: 120 },
      }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(comparison).toEqual({
      events: [
        { type: "unchanged", baseBlocks: [base[0]], revisedBlocks: [revised[0]] },
        { type: "unchanged", baseBlocks: [base[1]], revisedBlocks: [revised[1]] },
      ],
      structuralChanges: [],
    });
  });

  test("empty documents produce an empty ordered stream", () => {
    expect(successfulComparison({ base: [], revised: [] })).toEqual({
      events: [],
      structuralChanges: [],
    });
  });

  test("insertions and deletions stay in their logical stream positions", () => {
    const firstBase = contentBlock({ id: "first", text: "First clause" });
    const lastBase = contentBlock({ id: "last", text: "Last clause" });
    const firstRevised = contentBlock({ id: "first", text: "First clause" });
    const inserted = contentBlock({ id: "inserted", text: "New clause" });
    const lastRevised = contentBlock({ id: "last", text: "Last clause" });

    const insertion = successfulComparison({
      base: [firstBase, lastBase],
      revised: [firstRevised, inserted, lastRevised],
    });
    expect(insertion.events).toEqual([
      { type: "unchanged", baseBlocks: [firstBase], revisedBlocks: [firstRevised] },
      { type: "inserted", baseBlocks: [], revisedBlocks: [inserted] },
      { type: "unchanged", baseBlocks: [lastBase], revisedBlocks: [lastRevised] },
    ]);

    const deletion = successfulComparison({
      base: [firstRevised, inserted, lastRevised],
      revised: [firstBase, lastBase],
    });
    expect(deletion.events).toEqual([
      { type: "unchanged", baseBlocks: [firstRevised], revisedBlocks: [firstBase] },
      { type: "deleted", baseBlocks: [inserted], revisedBlocks: [] },
      { type: "unchanged", baseBlocks: [lastRevised], revisedBlocks: [lastBase] },
    ]);
  });

  test("a replacement carries ordered segments with offsets on both sides", () => {
    const base = contentBlock({ id: "term", text: "The fee is due." });
    const revised = contentBlock({ id: "term", text: "The tax is due." });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(comparison.events).toEqual([
      {
        type: "modified",
        baseBlocks: [base],
        revisedBlocks: [revised],
        changedProperties: [],
        segments: [
          {
            type: "equal",
            text: "The",
            baseStart: 0,
            baseEnd: 3,
            revisedStart: 0,
            revisedEnd: 3,
          },
          {
            type: "del",
            text: " fee",
            baseStart: 3,
            baseEnd: 7,
            revisedStart: 3,
            revisedEnd: 3,
          },
          {
            type: "ins",
            text: " tax",
            baseStart: 7,
            baseEnd: 7,
            revisedStart: 3,
            revisedEnd: 7,
          },
          {
            type: "equal",
            text: " is due.",
            baseStart: 7,
            baseEnd: 15,
            revisedStart: 7,
            revisedEnd: 15,
          },
        ],
      },
    ]);
  });

  test("paragraph splits and merges preserve their exact block tuples", () => {
    const joinedBase = contentBlock({ id: "first", text: "Alpha Beta" });
    const firstRevised = contentBlock({ id: "first", text: "Alpha" });
    const secondRevised = contentBlock({ id: "second", text: "Beta" });

    const split = successfulComparison({
      base: [joinedBase],
      revised: [firstRevised, secondRevised],
    });
    expect(split.events).toEqual([
      {
        type: "split",
        baseBlocks: [joinedBase],
        revisedBlocks: [firstRevised, secondRevised],
        offset: 5,
        separator: " ",
      },
    ]);

    const firstBase = contentBlock({ id: "first", text: "Alpha" });
    const secondBase = contentBlock({ id: "second", text: "Beta" });
    const joinedRevised = contentBlock({ id: "first", text: "Alpha Beta" });
    const merge = successfulComparison({
      base: [firstBase, secondBase],
      revised: [joinedRevised],
    });
    expect(merge.events).toEqual([
      {
        type: "merge",
        baseBlocks: [firstBase, secondBase],
        revisedBlocks: [joinedRevised],
        separator: " ",
      },
    ]);
  });

  test("exact and edited moves share one identity between their two stream positions", () => {
    const exactBase = contentBlock({
      id: "exact-base",
      text: "This clause remains exactly here",
    });
    const editedBase = contentBlock({
      id: "edited-base",
      text: "alpha beta gamma delta epsilon",
    });
    const anchors = stableAnchors();
    const exactRevised = contentBlock({
      id: "exact-revised",
      text: "This clause remains exactly here",
    });
    const editedRevised = contentBlock({
      id: "edited-revised",
      text: "alpha beta gamma delta zeta",
    });

    const comparison = successfulComparison({
      base: [exactBase, editedBase, ...anchors.base],
      revised: [...anchors.revised, exactRevised, editedRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "movedFrom",
      "movedFrom",
      "unchanged",
      "unchanged",
      "unchanged",
      "movedTo",
      "movedTo",
    ]);
    const movedFrom = comparison.events.filter(({ type }) => type === "movedFrom");
    const movedTo = comparison.events.filter(({ type }) => type === "movedTo");
    expect(
      movedFrom.map(({ moveId, baseBlocks }) => ({
        moveId,
        baseBlockId: baseBlocks[0].id,
      })),
    ).toEqual([
      { moveId: 1, baseBlockId: "exact-base" },
      { moveId: 2, baseBlockId: "edited-base" },
    ]);
    expect(
      movedTo.map(({ moveId, baseBlockId, revisedBlocks }) => ({
        moveId,
        baseBlockId,
        revisedBlockId: revisedBlocks[0].id,
      })),
    ).toEqual([
      { moveId: 1, baseBlockId: "exact-base", revisedBlockId: "exact-revised" },
      { moveId: 2, baseBlockId: "edited-base", revisedBlockId: "edited-revised" },
    ]);
    expect(movedTo[0]?.segments).toBeUndefined();
    expect(movedTo[1]?.segments).toEqual([
      {
        type: "equal",
        text: "alpha beta gamma delta",
        baseStart: 0,
        baseEnd: 22,
        revisedStart: 0,
        revisedEnd: 22,
      },
      {
        type: "del",
        text: " epsilon",
        baseStart: 22,
        baseEnd: 30,
        revisedStart: 22,
        revisedEnd: 22,
      },
      {
        type: "ins",
        text: " zeta",
        baseStart: 30,
        baseEnd: 30,
        revisedStart: 22,
        revisedEnd: 27,
      },
    ]);
  });

  test("repeated move candidates pair FIFO", () => {
    const repeatedText = "standard terms apply equally here";
    const firstBase = contentBlock({ id: "first-base", text: repeatedText });
    const secondBase = contentBlock({ id: "second-base", text: repeatedText });
    const firstRevised = contentBlock({ id: "first-revised", text: repeatedText });
    const secondRevised = contentBlock({ id: "second-revised", text: repeatedText });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [firstBase, secondBase, ...anchors.base],
      revised: [...anchors.revised, firstRevised, secondRevised],
    });
    const movedTo = comparison.events.filter(({ type }) => type === "movedTo");

    expect(movedTo.map(({ baseBlockId }) => baseBlockId)).toEqual([
      "first-base",
      "second-base",
    ]);
    expect(movedTo.map(({ revisedBlocks }) => revisedBlocks[0].id)).toEqual([
      "first-revised",
      "second-revised",
    ]);
  });

  test("short relocated boilerplate is not reported as a move", () => {
    const shortBase = contentBlock({ id: "short-base", text: "standard terms" });
    const shortRevised = contentBlock({ id: "short-revised", text: "standard terms" });
    const anchors = stableAnchors();

    const comparison = successfulComparison({
      base: [shortBase, ...anchors.base],
      revised: [...anchors.revised, shortRevised],
    });

    expect(eventTypes(comparison)).toEqual([
      "deleted",
      "unchanged",
      "unchanged",
      "unchanged",
      "inserted",
    ]);
  });

  test("paragraph-only formatting changes carry the revised property values", () => {
    const base = contentBlock({
      id: "clause",
      text: "Payment is due.",
      styleId: "Body",
      listLevel: 0,
      directAlignment: "left",
      directSpacing: { spaceBefore: 0, lineSpacingRule: "auto" },
    });
    const revised = contentBlock({
      id: "clause",
      text: "Payment is due.",
      styleId: "Clause",
      listLevel: 1,
      directAlignment: "center",
      directSpacing: { spaceAfter: 120, lineSpacingRule: "exact" },
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(comparison.events).toEqual([
      {
        type: "formatting",
        baseBlocks: [base],
        revisedBlocks: [revised],
        formatting: {
          paragraph: {
            styleId: "Clause",
            listLevel: 1,
            alignment: "center",
            spacing: { spaceAfter: 120, lineSpacingRule: "exact" },
          },
          ranges: [],
        },
      },
    ]);
  });

  test("inline-only formatting changes carry UTF-16 range offsets", () => {
    const base = contentBlock({
      id: "clause",
      text: "A😀B",
      previewRuns: [{ text: "A" }, { text: "😀B", bold: true }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "A😀B",
      previewRuns: [{ text: "A" }, { text: "😀B", italic: true }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(comparison.events).toEqual([
      {
        type: "formatting",
        baseBlocks: [base],
        revisedBlocks: [revised],
        formatting: {
          ranges: [
            {
              startOffset: 1,
              endOffset: 4,
              formatting: { bold: false, italic: true },
            },
          ],
        },
      },
    ]);
  });

  test("text segment offsets use UTF-16 boundaries compatible with string slicing", () => {
    const base = contentBlock({ id: "unicode", text: "A😀B" });
    const revised = contentBlock({ id: "unicode", text: "A😀XB" });

    const comparison = successfulComparison({
      base: [base],
      revised: [revised],
      granularity: "character",
    });
    const event = comparison.events[0];

    expect(event?.type).toBe("modified");
    if (event?.type !== "modified") {
      return;
    }
    expect(event.segments).toEqual([
      {
        type: "equal",
        text: "A😀",
        baseStart: 0,
        baseEnd: 3,
        revisedStart: 0,
        revisedEnd: 3,
      },
      {
        type: "ins",
        text: "X",
        baseStart: 3,
        baseEnd: 3,
        revisedStart: 3,
        revisedEnd: 4,
      },
      {
        type: "equal",
        text: "B",
        baseStart: 3,
        baseEnd: 4,
        revisedStart: 4,
        revisedEnd: 5,
      },
    ]);
    for (const segment of event.segments) {
      if (segment.type !== "ins") {
        expect(base.text.slice(segment.baseStart, segment.baseEnd)).toBe(segment.text);
      }
      if (segment.type !== "del") {
        expect(revised.text.slice(segment.revisedStart, segment.revisedEnd)).toBe(segment.text);
      }
    }
  });
});

describe("container-aware comparison", () => {
  test("stable identities never pair blocks across table cells", () => {
    const base = [
      tableBlock({ id: "left", text: "Alpha", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "right", text: "Beta", rowIndex: 0, cellIndex: 1 }),
    ];
    const revised = [
      tableBlock({ id: "right", text: "Beta", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "left", text: "Alpha", rowIndex: 0, cellIndex: 1 }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(eventTypes(comparison)).toEqual(["modified", "modified"]);
    expect(
      comparison.events.map(({ baseBlocks, revisedBlocks }) => [
        baseBlocks[0]?.table?.cellIndex,
        revisedBlocks[0]?.table?.cellIndex,
      ]),
    ).toEqual([
      [0, 0],
      [1, 1],
    ]);
    expect(
      comparison.events.map(({ baseBlocks, revisedBlocks }) => [
        baseBlocks[0]?.id,
        revisedBlocks[0]?.id,
      ]),
    ).toEqual([
      ["left", "right"],
      ["right", "left"],
    ]);
    expect(comparison.structuralChanges).toEqual([]);
  });

  test("an inserted table row is one structural change with row-major events", () => {
    const base = [
      tableBlock({ id: "header-a", text: "Header A", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "header-b", text: "Header B", rowIndex: 0, cellIndex: 1 }),
      tableBlock({ id: "existing-a", text: "Existing A", rowIndex: 1, cellIndex: 0 }),
      tableBlock({ id: "existing-b", text: "Existing B", rowIndex: 1, cellIndex: 1 }),
    ];
    const insertedA = tableBlock({
      id: "inserted-a",
      text: "Added X",
      rowIndex: 1,
      cellIndex: 0,
    });
    const insertedB = tableBlock({
      id: "inserted-b",
      text: "Added Y",
      rowIndex: 1,
      cellIndex: 1,
    });
    const revised = [
      tableBlock({ id: "header-a", text: "Header A", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "header-b", text: "Header B", rowIndex: 0, cellIndex: 1 }),
      insertedA,
      insertedB,
      tableBlock({ id: "existing-a", text: "Existing A", rowIndex: 2, cellIndex: 0 }),
      tableBlock({ id: "existing-b", text: "Existing B", rowIndex: 2, cellIndex: 1 }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(comparison.structuralChanges).toEqual([
      {
        id: 1,
        type: "table-row-insert",
        tableIndex: 0,
        rowIndex: 1,
        revisedBlockIds: ["inserted-a", "inserted-b"],
      },
    ]);
    expect(eventTypes(comparison)).toEqual([
      "unchanged",
      "unchanged",
      "inserted",
      "inserted",
      "unchanged",
      "unchanged",
    ]);
    expect(
      comparison.events.map(({ revisedBlocks }) => revisedBlocks[0]?.id),
    ).toEqual(revised.map(({ id }) => id));
    expect(
      comparison.events
        .filter(({ type }) => type === "inserted")
        .map(({ structuralChangeId }) => structuralChangeId),
    ).toEqual([1, 1]);
  });

  test("an inserted table column stays grouped while events remain row-major", () => {
    const base = [
      tableBlock({ id: "a0", text: "A0", rowIndex: 0, cellIndex: 0 }),
      tableBlock({ id: "a1", text: "A1", rowIndex: 0, cellIndex: 1 }),
      tableBlock({ id: "b0", text: "B0", rowIndex: 1, cellIndex: 0 }),
      tableBlock({ id: "b1", text: "B1", rowIndex: 1, cellIndex: 1 }),
    ];
    const insertedA = tableBlock({
      id: "ax",
      text: "AX",
      rowIndex: 0,
      cellIndex: 1,
      gridColumnIndex: 1,
    });
    const insertedB = tableBlock({
      id: "bx",
      text: "BX",
      rowIndex: 1,
      cellIndex: 1,
      gridColumnIndex: 1,
    });
    const revised = [
      tableBlock({ id: "a0", text: "A0", rowIndex: 0, cellIndex: 0 }),
      insertedA,
      tableBlock({
        id: "a1",
        text: "A1",
        rowIndex: 0,
        cellIndex: 2,
        gridColumnIndex: 2,
      }),
      tableBlock({ id: "b0", text: "B0", rowIndex: 1, cellIndex: 0 }),
      insertedB,
      tableBlock({
        id: "b1",
        text: "B1",
        rowIndex: 1,
        cellIndex: 2,
        gridColumnIndex: 2,
      }),
    ];

    const comparison = successfulComparison({ base, revised });

    expect(comparison.structuralChanges).toEqual([
      {
        id: 1,
        type: "table-column-insert",
        tableIndex: 0,
        columnIndex: 1,
        revisedBlockIds: ["ax", "bx"],
      },
    ]);
    expect(eventTypes(comparison)).toEqual([
      "unchanged",
      "inserted",
      "unchanged",
      "unchanged",
      "inserted",
      "unchanged",
    ]);
    expect(revisedProjection(comparison)).toEqual(revised);
    expect(baseProjection(comparison)).toEqual(base);
    expect(
      comparison.events
        .filter(({ type }) => type === "inserted")
        .map(({ structuralChangeId }) => structuralChangeId),
    ).toEqual([1, 1]);
  });

  test("a generic container-path change cannot become an unchanged pair", () => {
    const base = contentBlock({
      id: "clause",
      text: "Clause",
      containerPath: [{ kind: "section", id: "schedule-a" }],
    });
    const revised = contentBlock({
      id: "clause",
      text: "Clause",
      containerPath: [{ kind: "section", id: "schedule-b" }],
    });

    const comparison = successfulComparison({ base: [base], revised: [revised] });

    expect(comparison.events).toEqual([
      { type: "deleted", baseBlocks: [base], revisedBlocks: [] },
      { type: "inserted", baseBlocks: [], revisedBlocks: [revised] },
    ]);
  });
});

describe("identity semantics and input boundaries", () => {
  test("stable IDs survive a shift while positional IDs follow their content", () => {
    const stableBase = [
      contentBlock({ id: "a", text: "Repeated" }),
      contentBlock({ id: "b", text: "Repeated" }),
    ];
    const stableInserted = contentBlock({ id: "new", text: "Repeated" });
    const stableRevised = [
      stableInserted,
      contentBlock({ id: "a", text: "Repeated" }),
      contentBlock({ id: "b", text: "Repeated" }),
    ];
    const stable = successfulComparison({ base: stableBase, revised: stableRevised });

    expect(
      stable.events.map(({ type, baseBlocks, revisedBlocks }) => ({
        type,
        baseId: baseBlocks[0]?.id,
        revisedId: revisedBlocks[0]?.id,
      })),
    ).toEqual([
      { type: "inserted", baseId: undefined, revisedId: "new" },
      { type: "unchanged", baseId: "a", revisedId: "a" },
      { type: "unchanged", baseId: "b", revisedId: "b" },
    ]);

    const positionalBase = [
      contentBlock({ id: "0", text: "Alpha", idStability: "positional" }),
      contentBlock({ id: "1", text: "Beta", idStability: "positional" }),
    ];
    const positionalRevised = [
      contentBlock({ id: "0", text: "Inserted", idStability: "positional" }),
      contentBlock({ id: "1", text: "Alpha", idStability: "positional" }),
      contentBlock({ id: "2", text: "Beta", idStability: "positional" }),
    ];
    const positional = successfulComparison({
      base: positionalBase,
      revised: positionalRevised,
    });

    expect(
      positional.events.map(({ type, baseBlocks, revisedBlocks }) => ({
        type,
        baseId: baseBlocks[0]?.id,
        revisedId: revisedBlocks[0]?.id,
      })),
    ).toEqual([
      { type: "inserted", baseId: undefined, revisedId: "0" },
      { type: "unchanged", baseId: "0", revisedId: "1" },
      { type: "unchanged", baseId: "1", revisedId: "2" },
    ]);
  });

  test("duplicate IDs return a typed error with the exact side and block index", () => {
    const result = compareContent({
      base: {
        blocks: [
          contentBlock({ id: "duplicate", text: "First" }),
          contentBlock({ id: "duplicate", text: "Second" }),
        ],
      },
      revised: { blocks: [] },
    });

    expect(result.isErr()).toBe(true);
    if (!result.isErr()) {
      return;
    }
    expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
    expect(result.error).toMatchObject({
      side: "base",
      blockIndex: 1,
      field: "blocks[1].id",
    });
  });

  test("malformed optional projections return field-specific errors", () => {
    const cases: {
      block: TestBlock;
      field: string;
    }[] = [
      {
        block: contentBlock({
          id: "runs",
          text: "Whole text",
          previewRuns: [{ text: "Partial" }],
        }),
        field: "blocks[0].previewRuns",
      },
      {
        block: contentBlock({
          id: "container",
          text: "Text",
          containerPath: [{ kind: "", id: "section" }],
        }),
        field: "blocks[0].containerPath",
      },
      {
        block: tableBlock({
          id: "table",
          text: "Text",
          rowIndex: 0,
          cellIndex: 0,
          columnSpan: 0,
        }),
        field: "blocks[0].table.columnSpan",
      },
    ];

    for (const { block, field } of cases) {
      const result = compareContent({
        base: { blocks: [block] },
        revised: { blocks: [] },
      });
      expect(result.isErr()).toBe(true);
      if (!result.isErr()) {
        continue;
      }
      expect(result.error).toBeInstanceOf(InvalidFolioContentComparisonError);
      expect(result.error).toMatchObject({ side: "base", blockIndex: 0, field });
    }
  });

  test("comparison is deterministic and does not mutate frozen inputs", () => {
    const base: FolioContentSnapshot<TestBlockKind> = {
      blocks: [
        contentBlock({
          id: "a",
          text: "Alpha beta",
          previewRuns: [{ text: "Alpha ", bold: true }, { text: "beta" }],
          containerPath: [{ kind: "section", id: "main" }],
        }),
      ],
    };
    const revised: FolioContentSnapshot<TestBlockKind> = {
      blocks: [
        contentBlock({
          id: "a",
          text: "Alpha gamma",
          previewRuns: [{ text: "Alpha " }, { text: "gamma", italic: true }],
          containerPath: [{ kind: "section", id: "main" }],
        }),
      ],
    };
    const before = JSON.stringify({ base, revised });
    for (const snapshot of [base, revised]) {
      for (const block of snapshot.blocks) {
        block.previewRuns?.forEach((run) => {
          if (run.directFormatting) Object.freeze(run.directFormatting);
          Object.freeze(run);
        });
        block.containerPath?.forEach(Object.freeze);
        if (block.previewRuns) Object.freeze(block.previewRuns);
        if (block.containerPath) Object.freeze(block.containerPath);
        Object.freeze(block);
      }
      Object.freeze(snapshot.blocks);
      Object.freeze(snapshot);
    }

    const first = compareContent({ base, revised });
    const second = compareContent({ base, revised });

    expect(first.isErr()).toBe(false);
    expect(second.isErr()).toBe(false);
    if (first.isErr() || second.isErr()) {
      return;
    }
    expect(first.value).toEqual(second.value);
    expect(JSON.stringify({ base, revised })).toBe(before);
  });
});

const generatedText = fc
  .array(fc.constantFrom("alpha", " ", "beta", "\n", "§", "😀", "č", "م", "e\u0301"), {
    maxLength: 8,
  })
  .map((parts) => parts.join(""));

const generatedBlocks = fc.array(
  fc.record({
    kind: fc.constantFrom<TestBlockKind>("paragraph", "heading"),
    text: generatedText,
  }),
  { maxLength: 14 },
);

describe("comparison projection invariants", () => {
  test("every generated stream reconstructs both ordered inputs exactly", () => {
    fc.assert(
      fc.property(generatedBlocks, generatedBlocks, (baseValues, revisedValues) => {
        const base = baseValues.map(({ kind, text }, index) =>
          contentBlock({
            id: `base-${String(index)}`,
            kind,
            text,
            idStability: "positional",
          }),
        );
        const revised = revisedValues.map(({ kind, text }, index) =>
          contentBlock({
            id: `revised-${String(index)}`,
            kind,
            text,
            idStability: "positional",
          }),
        );
        const comparison = successfulComparison({ base, revised });

        expect(baseProjection(comparison)).toEqual(base);
        expect(revisedProjection(comparison)).toEqual(revised);
        for (const event of comparison.events) {
          const expectedCardinality = {
            unchanged: [1, 1],
            modified: [1, 1],
            formatting: [1, 1],
            inserted: [0, 1],
            deleted: [1, 0],
            movedFrom: [1, 0],
            movedTo: [0, 1],
            split: [1, 2],
            merge: [2, 1],
          } as const satisfies Record<
            FolioContentComparisonEvent<TestBlock>["type"],
            readonly [number, number]
          >;
          expect([event.baseBlocks.length, event.revisedBlocks.length]).toEqual(
            expectedCardinality[event.type],
          );
          if (event.type === "modified") {
            expect(textBefore(event.segments)).toBe(event.baseBlocks[0].text);
            expect(textAfter(event.segments)).toBe(event.revisedBlocks[0].text);
          }
          if (event.type === "movedTo" && event.segments) {
            const source = base.find(({ id }) => id === event.baseBlockId);
            expect(source).toBeDefined();
            expect(textBefore(event.segments)).toBe(source?.text);
            expect(textAfter(event.segments)).toBe(event.revisedBlocks[0].text);
          }
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });
});
