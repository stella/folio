import { expect, test } from "bun:test";

import { diffRecords, type DocumentRecord, type LineRecord } from "./engineLayoutDiff";

const line = (first: string, last: string, width: number | null = 100): LineRecord => ({
  first,
  last,
  width,
});

const record = (pages: LineRecord[][]): DocumentRecord => ({
  fixture: "doc.docx",
  fonts: [],
  pages: pages.map((lines) => ({ rendered: true, lines })),
});

test("identical records have no differences", () => {
  const a = record([[line("A", "b"), line("c", "d")]]);
  expect(diffRecords(a, structuredClone(a))).toEqual([]);
});

test("page and line count differences are reported with 1-based positions", () => {
  const expected = record([[line("A", "b"), line("c", "d")], [line("e", "f")]]);
  const actual = record([[line("A", "b")]]);
  expect(diffRecords(expected, actual)).toEqual([
    { kind: "page-count", document: "doc.docx", expected: 2, actual: 1 },
    { kind: "line-count", document: "doc.docx", page: 1, expected: 2, actual: 1 },
  ]);
});

test("a moved break reports the word and the width delta", () => {
  const expected = record([[line("A", "b", 400)]]);
  const actual = record([[line("A", "z", 390)]]);
  expect(diffRecords(expected, actual)).toEqual([
    {
      kind: "last-word",
      document: "doc.docx",
      page: 1,
      line: 1,
      expected: "b",
      actual: "z",
      widthDelta: -10,
    },
  ]);
});

test("width drift is reported only above the noise floor and only when words agree", () => {
  const expected = record([[line("A", "b", 400), line("c", "d", 400)]]);
  const actual = record([[line("A", "b", 400.2), line("c", "d", 402)]]);
  const found = diffRecords(expected, actual);
  expect(found).toHaveLength(1);
  expect(found[0]).toMatchObject({ kind: "width", line: 2, widthDelta: 2 });
});
