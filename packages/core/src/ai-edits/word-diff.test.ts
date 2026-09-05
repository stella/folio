/**
 * What `diffWordSegments` must produce, and what it must refuse to produce.
 *
 * The reconstruction and ordering rules are properties: they hold for every
 * pair of strings, and stating them as examples would only pin the pairs
 * someone happened to think of. The readability rules — the separator-only
 * rejection, the isolated-match rule, the fragmentation floor — get named
 * examples, because each one exists to make a specific redline read better and
 * the example IS the argument.
 */

import { describe, expect, test } from "bun:test";
import fc from "fast-check";

import { propertyConfig } from "../../../../test/property-testing";
import { diffWordSegments, type WordDiffSegment } from "./word-diff";

const rebuildBefore = (segments: readonly WordDiffSegment[]): string =>
  segments
    .filter(({ type }) => type !== "ins")
    .map(({ text }) => text)
    .join("");

const rebuildAfter = (segments: readonly WordDiffSegment[]): string =>
  segments
    .filter(({ type }) => type !== "del")
    .map(({ text }) => text)
    .join("");

/** Prose-shaped strings: the input the diff is for, rather than random bytes. */
const sentence = fc
  .array(fc.constantFrom("shall", "must", "the", "Goods", "Products", "thirty", "days", ",", "."), {
    minLength: 0,
    maxLength: 14,
  })
  .map((words) => words.join(" "));

describe("diffWordSegments", () => {
  test("both strings reconstruct from the segments, at either granularity", () => {
    fc.assert(
      fc.property(
        sentence,
        sentence,
        fc.constantFrom("word" as const, "character" as const),
        (before, after, granularity) => {
          const segments = diffWordSegments(before, after, { granularity });
          expect(rebuildBefore(segments)).toBe(before);
          expect(rebuildAfter(segments)).toBe(after);
        },
      ),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("identical strings produce one equal segment and no change", () => {
    fc.assert(
      fc.property(sentence, (text) => {
        const segments = diffWordSegments(text, text);
        expect(segments).toEqual(text.length === 0 ? [] : [{ type: "equal", text }]);
      }),
      propertyConfig({ numRuns: 100 }),
    );
  });

  test("adjacent segments never share a status, and a deletion never follows its insertion", () => {
    fc.assert(
      fc.property(sentence, sentence, (before, after) => {
        const types = diffWordSegments(before, after).map(({ type }) => type);
        for (const [index, type] of types.entries()) {
          expect(type).not.toBe(types[index - 1]);
          if (type === "del") {
            expect(types[index - 1]).not.toBe("ins");
          }
        }
      }),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("marks only the changed word in a small replacement", () => {
    expect(diffWordSegments("The quick fox jumps.", "The slow fox jumps.")).toEqual([
      { type: "equal", text: "The" },
      { type: "del", text: " quick" },
      { type: "ins", text: " slow" },
      { type: "equal", text: " fox jumps." },
    ]);
  });

  test("replaces a rewritten sentence whole rather than interleaving its coincidences", () => {
    // Both sides share "The", " days" and " receipt of" — enough for an LCS to
    // shred the sentence into six alternating fragments, and not enough for a
    // reader to see anything but a rewrite.
    const before =
      "The Supplier shall deliver the Goods within thirty days after receipt of the Purchase Order.";
    const after =
      "The Vendor must provide all Products no later than twenty business days following receipt of a valid order.";

    expect(diffWordSegments(before, after)).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
  });

  test("does not match on punctuation alone", () => {
    // The standalone comma is the only common token. Matching it would strike
    // out two halves of a clause separately and claim the comma survived both.
    const segments = diffWordSegments(
      "Payment falls due , without any deduction",
      "Interest accrues daily , at the statutory rate",
    );

    expect(segments.filter(({ type }) => type === "equal")).toEqual([]);
  });

  test("keeps a match that opens or closes the string, drops the island between two changes", () => {
    // Two changed words with an untouched run before, between and after them.
    expect(
      diffWordSegments(
        "The buyer shall pay the seller within thirty days.",
        "The buyer must pay the seller within sixty days.",
      ),
    ).toEqual([
      { type: "equal", text: "The buyer" },
      { type: "del", text: " shall" },
      { type: "ins", text: " must" },
      { type: "equal", text: " pay the seller within" },
      { type: "del", text: " thirty" },
      { type: "ins", text: " sixty" },
      { type: "equal", text: " days." },
    ]);

    // "goods" is the sole survivor between two rewritten halves.
    const island = diffWordSegments(
      "Risk in the goods passes to the buyer",
      "Title over goods vests in the purchaser",
    );
    expect(island.some(({ type, text }) => type === "equal" && text.trim() === "goods")).toBe(
      false,
    );
  });

  test("character granularity marks the changed letters inside one word", () => {
    expect(diffWordSegments("clause 14.2", "clause 14.3", { granularity: "character" })).toEqual([
      { type: "equal", text: "clause 14." },
      { type: "del", text: "2" },
      { type: "ins", text: "3" },
    ]);
  });

  test("word granularity replaces the whole token the character diff would split", () => {
    expect(diffWordSegments("clause 14.2", "clause 14.3")).toEqual([
      { type: "equal", text: "clause" },
      { type: "del", text: " 14.2" },
      { type: "ins", text: " 14.3" },
    ]);
  });

  test("case normalization reports a re-capitalized word as unchanged", () => {
    expect(
      diffWordSegments("The Supplier shall pay", "The supplier shall pay", {
        normalization: { case: true },
      }),
    ).toEqual([{ type: "equal", text: "The Supplier shall pay" }]);
    // Without it, the difference is a change like any other.
    expect(
      diffWordSegments("The Supplier shall pay", "The supplier shall pay").some(
        ({ type }) => type !== "equal",
      ),
    ).toBe(true);
  });

  test("whitespace normalization ignores a collapsed run of spaces", () => {
    expect(
      diffWordSegments("the goods  and the services", "the goods and the services", {
        normalization: { whitespace: true },
      }),
    ).toEqual([{ type: "equal", text: "the goods  and the services" }]);
  });

  test("returns empty segments for two empty strings", () => {
    expect(diffWordSegments("", "")).toEqual([]);
  });

  test("falls back to a single whole-string del+ins pair once the token-count product exceeds the cell budget", () => {
    // Regression guard for the word-diff DoS fix: diffWordSegments used to
    // allocate an unbounded (m+1)*(n+1) DP table for two attacker-controlled
    // strings inside one `modified` block pair. 2,001 distinct words each
    // tokenize to 2,001 tokens; 2,001 * 2,001 = 4,004,001 cells, just over
    // the 4,000,000-cell budget, so the DP must be skipped entirely.
    const before = Array.from({ length: 2001 }, (_unused, index) => `before${String(index)}`).join(
      " ",
    );
    const after = Array.from({ length: 2001 }, (_unused, index) => `after${String(index)}`).join(
      " ",
    );

    expect(diffWordSegments(before, after)).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
  });
});
