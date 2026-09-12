/**
 * The structural guard the serialize stage runs before a package is written.
 *
 * The round-trip verdict is covered where it is produced (`probes.test.ts`,
 * `compare.property.test.ts`); what is pinned here is the one finding that is
 * fatal whatever the caller asked for, because the package it describes does
 * not open.
 */

import { describe, expect, test } from "bun:test";

import type { FolioContentBlock } from "./content-types";
import { classifyContentProjectionMismatch, revisedFinalParagraphMarks } from "./verification";

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

describe("complete canonical projection verification", () => {
  const canonicalBlock = (): FolioContentBlock => ({
    identity: { type: "positional", id: "paragraph-1" },
    kind: "paragraph",
    text: "Alpha",
    blockProperties: [{ key: "docx.anchor", value: "anchor-1" }],
    paragraphFormatting: {
      authored: [{ key: "alignment", value: "left" }],
      effective: [{ key: "alignment", value: "left" }],
    },
    runs: [
      {
        text: "Alpha",
        authoredFormatting: [{ key: "bold", value: true }],
        effectiveFormatting: [{ key: "bold", value: true }],
      },
    ],
    structuralBoundaries: [{ type: "pageBreak", offset: 2 }],
    table: {
      outerTableIdentity: { type: "positional", id: "outer-table-0" },
      tableIdentity: { type: "positional", id: "table-1" },
      rowIdentity: { type: "positional", id: "row-0" },
      cellIdentity: { type: "positional", id: "cell-0" },
      outerTableIndex: 0,
      tableIndex: 1,
      rowIndex: 0,
      cellIndex: 0,
      gridColumnIndex: 0,
      columnSpan: 1,
      rowSpan: 1,
      paragraphIndex: 0,
    },
    containerPath: [
      { kind: "blockSdt", identity: { type: "positional", id: "0" } },
      { kind: "textBox", identity: { type: "positional", id: "0.0" } },
    ],
  });

  test.each([
    {
      name: "kind",
      cause: "unsupported",
      mutate: (block: FolioContentBlock) => Reflect.set(block, "kind", "heading"),
    },
    {
      name: "anchor property",
      cause: "unsupported",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block, "blockProperties", [{ key: "docx.anchor", value: "anchor-2" }]),
    },
    {
      name: "authored paragraph presentation",
      cause: "alignment",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block.paragraphFormatting, "authored", [{ key: "alignment", value: "right" }]),
    },
    {
      name: "effective paragraph presentation",
      cause: "unsupported",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block.paragraphFormatting, "effective", [{ key: "alignment", value: "right" }]),
    },
    {
      name: "authored run presentation",
      cause: "inline-formatting",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block.runs[0]!, "authoredFormatting", [{ key: "bold", value: false }]),
    },
    {
      name: "table location",
      cause: "container",
      mutate: (block: FolioContentBlock) => Reflect.set(block.table!, "cellIndex", 1),
    },
    {
      name: "nested table ownership",
      cause: "container",
      mutate: (block: FolioContentBlock) => Reflect.set(block.table!, "outerTableIndex", 2),
    },
    {
      name: "same-kind sibling container topology",
      cause: "container",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block.containerPath[1]!.identity, "id", "0.1"),
    },
    {
      name: "structural boundary",
      cause: "inline-structure",
      mutate: (block: FolioContentBlock) =>
        Reflect.set(block.structuralBoundaries[0]!, "offset", 3),
    },
  ])("rejects a same-id/text $name mutation at the exact story scope", ({ cause, mutate }) => {
    const expected = canonicalBlock();
    const actual = structuredClone(expected);
    const story = { type: "header", relationshipId: "rId7" } as const;
    mutate(actual);
    expect(actual.identity).toEqual(expected.identity);
    expect(actual.text).toBe(expected.text);
    const failure = classifyContentProjectionMismatch({
      invariant: "accept-reproduces-target",
      story,
      actual: [actual],
      expected: [expected],
    });
    expect(failure).toMatchObject({ cause });
    expect(failure?.scope).toEqual({ type: "story", story });
    expect(failure).not.toHaveProperty("story");
  });

  test("treats package-local block and table ids as transport identity", () => {
    const expected = canonicalBlock();
    const actual = structuredClone(expected);
    Reflect.set(actual.identity, "id", "other-paragraph-id");
    Reflect.set(actual.table?.cellIdentity ?? {}, "id", "other-cell-id");
    expect(
      classifyContentProjectionMismatch({
        invariant: "accept-reproduces-target",
        story: { type: "main" },
        actual: [actual],
        expected: [expected],
      }),
    ).toBeNull();
  });
});
