/**
 * The structural guard the serialize stage runs before a package is written.
 *
 * The round-trip verdict is covered where it is produced (`probes.test.ts`,
 * `compare.property.test.ts`); what is pinned here is the one finding that is
 * fatal whatever the caller asked for, because the package it describes does
 * not open.
 */

import { describe, expect, test } from "bun:test";

import type { FolioAIBlock, FolioAIBlockTableLocation } from "../ai-edits/types";
import { getCompareSkipDisposition } from "./compare";
import { classifyProjectionMismatch, revisedFinalParagraphMarks } from "./verification";

const revision = { id: 1, author: "compare", date: "2024-03-01T00:00:00.000Z" };

const paragraph = (text: string, pPrMark?: { kind: string }) => ({
  type: "paragraph",
  content: [{ type: "run", content: [{ type: "text", text }] }],
  ...(pPrMark && { pPrMark: { kind: pPrMark.kind, info: revision } }),
});

const cell = (content: unknown[]) => ({ type: "tableCell", content });

const table = (cells: unknown[][]) => ({
  type: "table",
  rows: [{ type: "tableRow", cells: cells.map((content) => cell(content)) }],
});

test("a singular run-formatting ownership conflict is an unwritable comparison slice", () => {
  expect(getCompareSkipDisposition("pendingRunPropertyChange")).toBe("unwritable");
});

describe("revisedFinalParagraphMarks", () => {
  test("a body whose last paragraph mark is deleted is named", () => {
    expect(
      revisedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("last", { kind: "del" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "del" }]);
  });

  test("a relocation's source break counts, because it resolves the same way", () => {
    expect(
      revisedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("last", { kind: "moveFrom" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "moveFrom" }]);
  });

  test("a mark deleted on a paragraph that something follows is not a finding", () => {
    expect(
      revisedFinalParagraphMarks({
        document: {
          content: [paragraph("first", { kind: "del" }), paragraph("last")],
        },
      }),
    ).toEqual([]);
  });

  test("an inserted final mark is a finding too: nothing follows it to close over", () => {
    // The break was ADDED, and rejecting an added break closes the paragraph
    // it ends back over the NEXT one. A container's last paragraph has none,
    // so the mark states an edit no reader can carry out in either direction:
    // it survives accepting everything and rejecting everything alike.
    expect(
      revisedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("added", { kind: "ins" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "ins" }]);
  });

  test("a relocation's destination break counts for the same reason", () => {
    expect(
      revisedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("moved", { kind: "moveTo" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "moveTo" }]);
  });

  test("a cell, a header and a note are containers too", () => {
    const found = revisedFinalParagraphMarks({
      document: {
        content: [table([[paragraph("cell", { kind: "del" })]]), paragraph("last")],
      },
      headers: new Map([["rId2", { content: [paragraph("header", { kind: "del" })] }]]),
      footnotes: [{ id: 2, content: [paragraph("note", { kind: "del" })] }],
    });
    expect(found.map(({ container }) => container).toSorted()).toEqual([
      "package.document.content[0].rows[0].cells[0].content",
      "package.footnotes[0].content",
      "package.headers.rId2.content",
    ]);
  });

  test("a mark the base arrived with is not this comparison's to report", () => {
    // A base can carry one on a paragraph of a part no story mounts, so
    // resolving it to its accepted view does not reach it. `since` is one past
    // the highest id the base already used, so what is left is what this
    // comparison wrote.
    expect(
      revisedFinalParagraphMarks(
        { document: { content: [paragraph("first"), paragraph("last", { kind: "del" })] } },
        { since: revision.id + 1 },
      ),
    ).toEqual([]);
  });

  test("a package with nothing on any final mark reports nothing", () => {
    expect(
      revisedFinalParagraphMarks({
        document: { content: [table([[paragraph("cell")]]), paragraph("last")] },
      }),
    ).toEqual([]);
  });
});

describe("classifyProjectionMismatch", () => {
  const projectedBlock = (
    tableLocation: FolioAIBlockTableLocation | undefined,
    text: string,
  ): FolioAIBlock => ({
    id: "projected-block",
    kind: "paragraph",
    text,
    ...(tableLocation ? { table: tableLocation } : {}),
  });

  test("names table-cell containers from the complete projected coordinate", () => {
    const tableContainer = {
      outerTableIndex: 0,
      tableIndex: 0,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    } satisfies FolioAIBlockTableLocation;
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [projectedBlock(tableContainer, "same text")],
        expected: [projectedBlock(undefined, "same text")],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "container",
      story: { type: "main" },
      detail:
        "a block sits in a cell where it is expected in a body, at block 0/1 (1 blocks against 1)",
    });
  });

  test("normalizes hidden-row coordinate gaps while retaining cell geometry", () => {
    const container = (
      outerTableIndex: number,
      tableIndex: number,
      rowIndex: number,
      cellIndex: number,
      paragraphIndex: number,
    ): FolioAIBlockTableLocation => ({
      outerTableIndex,
      tableIndex,
      rowIndex,
      cellIndex,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex,
    });
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [
          projectedBlock(container(0, 1, 0, 0, 0), "first"),
          projectedBlock(container(0, 1, 2, 0, 0), "second"),
        ],
        expected: [
          projectedBlock(container(4, 9, 5, 7, 3), "first"),
          projectedBlock(container(4, 9, 9, 7, 8), "second"),
        ],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "invisible-structure",
      story: { type: "main" },
      detail: "every block matches once table coordinates count visible blocks (2 blocks)",
    });
  });

  test("distinguishes a nested table moved to another outer table", () => {
    const location = (outerTableIndex: number, tableIndex: number): FolioAIBlockTableLocation => ({
      outerTableIndex,
      tableIndex,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    });
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [projectedBlock(location(0, 0), "outer"), projectedBlock(location(0, 2), "nested")],
        expected: [
          projectedBlock(location(0, 0), "outer"),
          projectedBlock(location(1, 2), "nested"),
        ],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "container",
      story: { type: "main" },
      detail:
        "a block sits in a cell where it is expected in a cell, at block 1/2 (2 blocks against 2)",
    });
  });

  test.each([
    {
      label: "count",
      actual: [{ type: "pageBreak" as const, offset: 4 }],
      expected: [
        { type: "pageBreak" as const, offset: 4 },
        { type: "pageBreak" as const, offset: 4 },
      ],
    },
    {
      label: "offset",
      actual: [{ type: "pageBreak" as const, offset: 3 }],
      expected: [{ type: "pageBreak" as const, offset: 4 }],
    },
    {
      label: "order",
      actual: [
        { type: "pageBreak" as const, offset: 4, clear: "left" as const },
        { type: "pageBreak" as const, offset: 4, clear: "right" as const },
      ],
      expected: [
        { type: "pageBreak" as const, offset: 4, clear: "right" as const },
        { type: "pageBreak" as const, offset: 4, clear: "left" as const },
      ],
    },
    {
      label: "preserved clear",
      actual: [{ type: "pageBreak" as const, offset: 4, clear: "left" as const }],
      expected: [{ type: "pageBreak" as const, offset: 4, clear: "right" as const }],
    },
  ])("reports an inline page-break $label mismatch", ({ actual, expected }) => {
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [{ ...projectedBlock(undefined, "same text"), structuralBoundaries: actual }],
        expected: [{ ...projectedBlock(undefined, "same text"), structuralBoundaries: expected }],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "inline-structure",
      story: { type: "main" },
      detail:
        "a block's zero-width inline structure does not match at block 0/1 (1 blocks against 1)",
    });
  });

  test("treats absent, empty, and equal inline-structure projections as equivalent", () => {
    const leftBoundary = { type: "pageBreak" as const, offset: 4, clear: "left" as const };
    const equivalentPairs = [
      [undefined, undefined],
      [undefined, []],
      [[], undefined],
      [[], []],
      [[leftBoundary], [{ ...leftBoundary }]],
    ] as const;

    for (const [actual, expected] of equivalentPairs) {
      expect(
        classifyProjectionMismatch({
          invariant: "accept-reproduces-target",
          story: { type: "main" },
          actual: [{ ...projectedBlock(undefined, "same text"), structuralBoundaries: actual }],
          expected: [{ ...projectedBlock(undefined, "same text"), structuralBoundaries: expected }],
        }),
      ).toBeNull();
    }
  });

  test("reports a direct alignment mismatch separately from text and style", () => {
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [
          { ...projectedBlock(undefined, "same text"), styleId: "Body", directAlignment: "left" },
        ],
        expected: [
          {
            ...projectedBlock(undefined, "same text"),
            styleId: "Body",
            directAlignment: "right",
          },
        ],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "alignment",
      story: { type: "main" },
      detail: "the direct paragraph alignment did not move at block 0/1 (1 blocks against 1)",
    });
  });

  test("reports a direct spacing mismatch separately from text and style", () => {
    expect(
      classifyProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [{ ...projectedBlock(undefined, "same text"), directSpacing: { spaceAfter: 0 } }],
        expected: [
          { ...projectedBlock(undefined, "same text"), directSpacing: { spaceAfter: 240 } },
        ],
      }),
    ).toEqual({
      invariant: "accept-reproduces-target",
      cause: "spacing",
      story: { type: "main" },
      detail: "the direct paragraph spacing did not move at block 0/1 (1 blocks against 1)",
    });
  });
});
