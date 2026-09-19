import { describe, expect, test } from "bun:test";

import {
  CorpusSelectionError,
  corpusFileId,
  parseOnlySelection,
  selectOnly,
} from "./lib/corpus-selection";

const ENTRIES = [
  { sourceId: "apache-poi", relativePath: "test-data/document/55733.docx" },
  { sourceId: "apache-poi", relativePath: "test-data/document/moves.docx" },
  { sourceId: "docx4j", relativePath: "sample-docs/comments.docx" },
];

const select = (patterns: readonly string[]) =>
  selectOnly({ entries: ENTRIES, patterns, idOf: corpusFileId }).map(corpusFileId);

describe("parseOnlySelection", () => {
  test("absent means the whole corpus", () => {
    expect(parseOnlySelection([])).toBeUndefined();
  });

  test("splits a comma-separated list and collapses repeats", () => {
    expect(parseOnlySelection(["a.docx,b.docx", " a.docx "])).toEqual(["a.docx", "b.docx"]);
  });

  test("refuses a flag with nothing in it", () => {
    expect(() => parseOnlySelection([" , "])).toThrow(CorpusSelectionError);
  });
});

describe("selectOnly", () => {
  test("selects by full file id", () => {
    expect(select(["docx4j/sample-docs/comments.docx"])).toEqual([
      "docx4j/sample-docs/comments.docx",
    ]);
  });

  test("selects by any part of the id, so a whole source works", () => {
    expect(select(["apache-poi"])).toEqual([
      "apache-poi/test-data/document/55733.docx",
      "apache-poi/test-data/document/moves.docx",
    ]);
  });

  test("returns each file once when two patterns select it", () => {
    expect(select(["apache-poi/test-data/document/moves.docx", "moves"])).toEqual([
      "apache-poi/test-data/document/moves.docx",
    ]);
  });

  test("refuses a pattern that matches nothing, so a typo is not an empty run", () => {
    expect(() => select(["apache-poi", "no-such-file.docx"])).toThrow(CorpusSelectionError);
  });
});
