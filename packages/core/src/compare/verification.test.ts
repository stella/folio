/**
 * The structural guard the serialize stage runs before a package is written.
 *
 * The round-trip verdict is covered where it is produced (`probes.test.ts`,
 * `compare.property.test.ts`); what is pinned here is the one finding that is
 * fatal whatever the caller asked for, because the package it describes does
 * not open.
 */

import { describe, expect, test } from "bun:test";

import { deletedFinalParagraphMarks } from "./verification";

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

describe("deletedFinalParagraphMarks", () => {
  test("a body whose last paragraph mark is deleted is named", () => {
    expect(
      deletedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("last", { kind: "del" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "del" }]);
  });

  test("a relocation's source break counts, because it resolves the same way", () => {
    expect(
      deletedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("last", { kind: "moveFrom" })] },
      }),
    ).toEqual([{ container: "package.document.content", paragraphIndex: 1, kind: "moveFrom" }]);
  });

  test("a mark deleted on a paragraph that something follows is not a finding", () => {
    expect(
      deletedFinalParagraphMarks({
        document: {
          content: [paragraph("first", { kind: "del" }), paragraph("last")],
        },
      }),
    ).toEqual([]);
  });

  test("an inserted final mark is not a finding: rejecting it removes the paragraph it added", () => {
    expect(
      deletedFinalParagraphMarks({
        document: { content: [paragraph("first"), paragraph("added", { kind: "ins" })] },
      }),
    ).toEqual([]);
  });

  test("a cell, a header and a note are containers too", () => {
    const found = deletedFinalParagraphMarks({
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

  test("a package with nothing on any final mark reports nothing", () => {
    expect(
      deletedFinalParagraphMarks({
        document: { content: [table([[paragraph("cell")]]), paragraph("last")] },
      }),
    ).toEqual([]);
  });
});
