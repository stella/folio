import { describe, expect, test } from "bun:test";

import { selectFailure } from "./corpus-minimize";
import { EXTENDED_CORPUS_INVARIANTS } from "./lib/corpus-invariants/contract";
import { type CorpusFailure, NO_FRAME } from "./lib/corpus-signature";

const failure = (message: string): CorpusFailure => ({
  invariant: EXTENDED_CORPUS_INVARIANTS.reserialize,
  message,
  frame: NO_FRAME,
});

const CANDIDATES = [
  failure("package.document.content[].bold: true became false"),
  failure("package.document.content[].italic: true became false"),
  failure("package.document.content[].runs[]: length changed"),
];

/**
 * A file now reports every difference it exhibits, so shrinking towards "the
 * first failure of the chosen invariant" is a selection nobody made.
 */
describe("selectFailure", () => {
  test("without a selector it takes the file's first failure", () => {
    expect(selectFailure(CANDIDATES, undefined)).toBe(CANDIDATES[0] as CorpusFailure);
  });

  test("a selector picks the one signature that contains it", () => {
    expect(selectFailure(CANDIDATES, "italic")).toBe(CANDIDATES[1] as CorpusFailure);
    expect(selectFailure(CANDIDATES, "length changed")).toBe(CANDIDATES[2] as CorpusFailure);
  });

  /**
   * Refused rather than resolved by position: the wrong choice costs hundreds
   * of evaluations before it shows, and the candidates are worth reading.
   */
  test("an ambiguous or unmatched selector names the candidates", () => {
    expect(() => selectFailure(CANDIDATES, "true became false")).toThrow(/matches 2/u);
    expect(() => selectFailure(CANDIDATES, "underline")).toThrow(/matches none/u);
    expect(() => selectFailure(CANDIDATES, "underline")).toThrow(/length changed/u);
  });
});
