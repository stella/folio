/**
 * The pure half of a direct-mode replacement: which characters change.
 *
 * Whatever the strings, applying the planned changes must give exactly the
 * replacement, the changes must be disjoint and in order, each must start and
 * end where the strings differ, and none may cut into a field result.
 */

import { describe, expect, setDefaultTimeout, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig, propertyTestTimeout } from "../../../../test/property-testing";

import {
  applyTextChanges,
  type AtomicTextSpan,
  commonPrefixLength,
  commonSuffixLength,
  keepAtomicSpans,
  planTextChanges,
  shortestTokenDiff,
  type TextChange,
  widenChangesToAtomicSpans,
} from "./minimal-replacement";
import { tokenizeWords } from "./word-diff";

setDefaultTimeout(propertyTestTimeout(10_000));

/** Letters, spaces, punctuation and a surrogate pair, so tokens and code points both matter. */
const fragment = fc.constantFrom(
  "the",
  "Seller",
  "shall",
  " ",
  "  ",
  ",",
  ".",
  "(a)",
  "3.6",
  "\t",
  "\n",
  "😀",
  "é",
  "x",
);
const text = fc.array(fragment, { maxLength: 14 }).map((parts) => parts.join(""));

const assertWellFormed = (source: string, replacement: string, changes: readonly TextChange[]) => {
  expect(applyTextChanges(source, changes)).toBe(replacement);
  let previousEnd = 0;
  for (const change of changes) {
    expect(change.start).toBeGreaterThanOrEqual(previousEnd);
    expect(change.start).toBeLessThanOrEqual(change.end);
    expect(change.start < change.end || change.text.length > 0).toBe(true);
    previousEnd = change.end;
  }
};

const splitsSurrogatePair = (value: string, offset: number): boolean => {
  const before = value.charCodeAt(offset - 1);
  const after = value.charCodeAt(offset);
  return before >= 0xd8_00 && before <= 0xdb_ff && after >= 0xdc_00 && after <= 0xdf_ff;
};

const WORD_CHARACTER = /[\p{L}\p{N}\p{M}\p{Pc}]/u;

/** At most one run of word characters: a change inside one word. */
const isOneWord = (value: string): boolean =>
  (value.match(/[\p{L}\p{N}\p{M}\p{Pc}]+/gu) ?? []).length <= 1;

const isWordBoundary = (value: string, index: number): boolean => {
  if (index <= 0 || index >= value.length) {
    return true;
  }
  const before = String.fromCodePoint(value.codePointAt(index - 1) ?? 0);
  const after = String.fromCodePoint(value.codePointAt(index) ?? 0);
  return !WORD_CHARACTER.test(before) || !WORD_CHARACTER.test(after);
};

/** Token LCS by dynamic programming, the length the shortest edit script must match. */
const lcsLength = (before: readonly string[], after: readonly string[]): number => {
  let previous: number[] = Array.from({ length: after.length + 1 }, () => 0);
  for (const token of before) {
    const current = [0];
    for (let column = 0; column < after.length; column++) {
      current.push(
        token === after[column]
          ? (previous[column] ?? 0) + 1
          : Math.max(previous[column + 1] ?? 0, current[column] ?? 0),
      );
    }
    previous = current;
  }
  return previous.at(-1) ?? 0;
};

describe("planTextChanges", () => {
  test("appending a character changes only the end", () => {
    expect(planTextChanges("Bold start italic end", "Bold start italic end‸")).toEqual([
      { start: 21, end: 21, text: "‸" },
    ]);
  });

  test("a changed word keeps the letters it shares", () => {
    expect(planTextChanges("pay John and 3.6 now", "pay Jane and 3.7 now")).toEqual([
      { start: 5, end: 8, text: "ane" },
      { start: 15, end: 16, text: "7" },
    ]);
  });

  test("a change across words keeps only whole words it shares", () => {
    // `agrees` and `performs` share a final `s`; keeping it would leave the
    // last letter of `performs` in the formatting of `agrees`.
    expect(planTextChanges("Supplier agrees", "Provider performs")).toEqual([
      { start: 0, end: 15, text: "Provider performs" },
    ]);
    // Within one word the shared letters stay: they belong to the same word.
    expect(planTextChanges("the Supplier agrees", "the Provider agrees")).toEqual([
      { start: 4, end: 10, text: "Provid" },
    ]);
    expect(planTextChanges("3.6the", "3.6")).toEqual([{ start: 3, end: 6, text: "" }]);
  });

  test("a change within one word keeps the letters it shares", () => {
    expect(planTextChanges("Supplier", "Suppliers")).toEqual([{ start: 8, end: 8, text: "s" }]);
  });

  test("a short match between two changes is kept", () => {
    expect(planTextChanges("a b c", "x b y")).toEqual([
      { start: 0, end: 1, text: "x" },
      { start: 4, end: 5, text: "y" },
    ]);
  });

  test("reconstructs the replacement, in order, trimmed, never splitting a pair", () => {
    fc.assert(
      fc.property(text, text, (source, replacement) => {
        const changes = planTextChanges(source, replacement);
        assertWellFormed(source, replacement, changes);
        for (const change of changes) {
          const removed = source.slice(change.start, change.end);
          const prefix = commonPrefixLength(removed, change.text);
          const suffix = commonSuffixLength(removed, change.text);
          if (isOneWord(removed) && isOneWord(change.text)) {
            // Trimmed: nothing the removed and inserted text share at an end is rewritten.
            expect(prefix).toBe(0);
            expect(suffix).toBe(0);
          } else {
            // Across words, only what ends on a word boundary of both is kept.
            for (let length = 1; length <= prefix; length++) {
              expect(isWordBoundary(removed, length) && isWordBoundary(change.text, length)).toBe(
                false,
              );
            }
            for (let length = 1; length <= Math.min(suffix, removed.length - prefix); length++) {
              expect(
                isWordBoundary(removed, removed.length - length) &&
                  isWordBoundary(change.text, change.text.length - length),
              ).toBe(false);
            }
          }
          expect(splitsSurrogatePair(source, change.start)).toBe(false);
          expect(splitsSurrogatePair(source, change.end)).toBe(false);
        }
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });

  test("identical strings need no change", () => {
    fc.assert(
      fc.property(text, (source) => {
        expect(planTextChanges(source, source)).toEqual([]);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });
});

describe("shortestTokenDiff", () => {
  test("keeps exactly a longest common subsequence", () => {
    // One-letter tokens, so the kept text's length is the kept token count.
    const tokens = fc.array(fc.constantFrom("a", "b", "c"), { maxLength: 24 });
    fc.assert(
      fc.property(tokens, tokens, (before, after) => {
        const segments = shortestTokenDiff(before, after) ?? [];
        const joined = (skip: "ins" | "del") =>
          segments
            .filter((segment) => segment.type !== skip)
            .map((segment) => segment.text)
            .join("");
        expect(joined("ins")).toBe(before.join(""));
        expect(joined("del")).toBe(after.join(""));
        const kept = segments
          .filter((segment) => segment.type === "equal")
          .reduce((sum, segment) => sum + segment.text.length, 0);
        expect(kept).toBe(lcsLength(before, after));
      }),
      propertyConfig({ numRuns: 400 }),
    );
  });

  test("reconstructs both strings over word tokens", () => {
    fc.assert(
      fc.property(text, text, (source, replacement) => {
        const segments = shortestTokenDiff(tokenizeWords(source), tokenizeWords(replacement)) ?? [];
        expect(
          segments
            .filter((segment) => segment.type !== "ins")
            .map((s) => s.text)
            .join(""),
        ).toBe(source);
        expect(
          segments
            .filter((segment) => segment.type !== "del")
            .map((s) => s.text)
            .join(""),
        ).toBe(replacement);
      }),
      propertyConfig({ numRuns: 300 }),
    );
  });

  test("gives up past the edit budget", () => {
    expect(shortestTokenDiff(["a", "b", "c"], ["x", "y", "z"], 2)).toBeNull();
    expect(shortestTokenDiff(["a", "b"], ["a", "b"], 0)).toEqual([{ type: "equal", text: "ab" }]);
  });
});

describe("widenChangesToAtomicSpans", () => {
  const spansIn = (source: string) =>
    fc
      .array(
        fc.tuple(fc.nat({ max: Math.max(source.length, 1) }), fc.integer({ min: 2, max: 5 })),
        { maxLength: 3 },
      )
      .map((pairs) => {
        const spans: AtomicTextSpan[] = [];
        let cursor = 0;
        for (const [gap, length] of pairs.toSorted(([a], [b]) => a - b)) {
          const offset = Math.max(cursor, gap);
          if (offset + length > source.length) {
            break;
          }
          spans.push({ offset, length });
          cursor = offset + length;
        }
        return spans;
      });

  test("never cuts into a span and still reconstructs the replacement", () => {
    fc.assert(
      fc.property(
        text.chain((source) => fc.tuple(fc.constant(source), text, spansIn(source))),
        ([source, replacement, spans]) => {
          const changes = widenChangesToAtomicSpans(
            source,
            planTextChanges(source, replacement),
            spans,
          );
          expect(applyTextChanges(source, changes)).toBe(replacement);
          for (const change of changes) {
            for (const { offset, length } of spans) {
              const cuts = (position: number) => position > offset && position < offset + length;
              expect({ change, cutsStart: cuts(change.start), cutsEnd: cuts(change.end) }).toEqual({
                change,
                cutsStart: false,
                cutsEnd: false,
              });
            }
          }
          for (let index = 1; index < changes.length; index++) {
            expect(changes[index]!.start).toBeGreaterThanOrEqual(changes[index - 1]!.end);
          }
        },
      ),
      propertyConfig({ numRuns: 400 }),
    );
  });

  test("a change inside a field result replaces the whole result", () => {
    expect(
      widenChangesToAtomicSpans(
        "see 3.6 above",
        [{ start: 6, end: 7, text: "7" }],
        [{ offset: 4, length: 3 }],
      ),
    ).toEqual([{ start: 4, end: 7, text: "3.7" }]);
  });

  test("changes the widening makes overlap become one", () => {
    expect(
      widenChangesToAtomicSpans(
        "ab123cd",
        [
          { start: 2, end: 3, text: "9" },
          { start: 4, end: 5, text: "8" },
        ],
        [{ offset: 2, length: 3 }],
      ),
    ).toEqual([{ start: 2, end: 5, text: "928" }]);
  });
});

describe("keepAtomicSpans", () => {
  test("a word rewritten into a field keeps the field and marks only what changed", () => {
    // Tab, a field showing `3.6`, then `Seller `. The word diff rewrites the
    // whole `\t3.6Seller ` word; the direct plan changes only the tab. Cut
    // around the kept field, the text after it rewrites `Seller ` to itself,
    // which is no change at all.
    const source = "\t3.6Seller ";
    const replacement = " and3.6Seller ";
    expect(
      keepAtomicSpans(
        source,
        replacement,
        [{ start: 0, end: 11, text: replacement }],
        planTextChanges(source, replacement),
        [{ offset: 1, length: 3 }],
      ),
    ).toEqual([{ start: 0, end: 1, text: " and" }]);
  });

  test("a word prefixed in front of a field is an insertion before the field", () => {
    expect(
      keepAtomicSpans(
        "3.6",
        "X3.6",
        [{ start: 0, end: 3, text: "X3.6" }],
        planTextChanges("3.6", "X3.6"),
        [{ offset: 0, length: 3 }],
      ),
    ).toEqual([{ start: 0, end: 0, text: "X" }]);
  });
});
