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
import { createWordDiffSession, diffWordSegments, type WordDiffSegment } from "./text-diff";

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

const unicodeText = fc
  .array(fc.constantFrom("a", " ", "\n", "😀", "👩‍⚖️", "§", "č", "م", "क", "e\u0301"), {
    minLength: 0,
    maxLength: 40,
  })
  .map((units) => units.join(""));

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

  test("arbitrary Unicode reconstructs exactly and aligns deterministically", () => {
    fc.assert(
      fc.property(
        unicodeText,
        unicodeText,
        fc.constantFrom("word" as const, "character" as const),
        (before, after, granularity) => {
          const first = diffWordSegments(before, after, { granularity });
          const second = diffWordSegments(before, after, { granularity });
          expect(first).toEqual(second);
          expect(rebuildBefore(first)).toBe(before);
          expect(rebuildAfter(first)).toBe(after);
        },
      ),
      propertyConfig({ numRuns: 200 }),
    );
  });

  test("a token-subsequence edit never invents the opposite change direction", () => {
    const tokenSequence = fc.array(fc.tuple(fc.constantFrom("a", "b", "c"), fc.boolean()), {
      maxLength: 9,
    });
    fc.assert(
      fc.property(tokenSequence, (entries) => {
        const whole = entries.map(([token]) => token).join("");
        const subsequence = entries
          .filter(([, retained]) => retained)
          .map(([token]) => token)
          .join("");

        const insertion = diffWordSegments(subsequence, whole, { granularity: "character" });
        expect(insertion.some(({ type }) => type === "del")).toBe(false);
        expect(rebuildBefore(insertion)).toBe(subsequence);
        expect(rebuildAfter(insertion)).toBe(whole);

        const deletion = diffWordSegments(whole, subsequence, { granularity: "character" });
        expect(deletion.some(({ type }) => type === "ins")).toBe(false);
        expect(rebuildBefore(deletion)).toBe(whole);
        expect(rebuildAfter(deletion)).toBe(subsequence);

        const wholeWords = ["anchor", ...entries.map(([token]) => token)].join(" ");
        const subsequenceWords = [
          "anchor",
          ...entries.filter(([, retained]) => retained).map(([token]) => token),
        ].join(" ");
        const wordInsertion = diffWordSegments(subsequenceWords, wholeWords);
        expect(wordInsertion.some(({ type }) => type === "del")).toBe(false);
        expect(rebuildBefore(wordInsertion)).toBe(subsequenceWords);
        expect(rebuildAfter(wordInsertion)).toBe(wholeWords);
        const wordDeletion = diffWordSegments(wholeWords, subsequenceWords);
        expect(wordDeletion.some(({ type }) => type === "ins")).toBe(false);
        expect(rebuildBefore(wordDeletion)).toBe(wholeWords);
        expect(rebuildAfter(wordDeletion)).toBe(subsequenceWords);
      }),
      propertyConfig({ numRuns: 500 }),
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

  test("anchors a unique term instead of matching more repeated boilerplate", () => {
    expect(
      diffWordSegments("the the the the the SIGNATURE PAGE", "the the SIGNATURE PAGE the the the"),
    ).toEqual([
      { type: "equal", text: "the the" },
      { type: "del", text: " the the the" },
      { type: "equal", text: " SIGNATURE PAGE" },
      { type: "ins", text: " the the the" },
    ]);
  });

  test("preserves the historical LCS tie break when no unique anchor is selected", () => {
    const cases = [
      {
        before: "aabb",
        after: "bbaa",
        granularity: "character" as const,
        expected: [
          { type: "del" as const, text: "aa" },
          { type: "equal" as const, text: "bb" },
          { type: "ins" as const, text: "aa" },
        ],
      },
      {
        before: "abab",
        after: "baba",
        granularity: "character" as const,
        expected: [
          { type: "del" as const, text: "a" },
          { type: "equal" as const, text: "bab" },
          { type: "ins" as const, text: "a" },
        ],
      },
      {
        before: "the the the",
        after: "the the",
        granularity: "word" as const,
        expected: [
          { type: "equal" as const, text: "the" },
          { type: "del" as const, text: " the" },
          { type: "equal" as const, text: " the" },
        ],
      },
    ];

    for (const { before, after, granularity, expected } of cases) {
      expect(diffWordSegments(before, after, { granularity })).toEqual(expected);
    }
  });

  test("preserves historical ties when a weak selected anchor would be discarded", () => {
    expect(diffWordSegments("ab", "aaba", { granularity: "character" })).toEqual([
      { type: "ins", text: "a" },
      { type: "equal", text: "ab" },
      { type: "ins", text: "a" },
    ]);
    expect(diffWordSegments("the the clause", "the the the clause the")).toEqual([
      { type: "equal", text: "the" },
      { type: "ins", text: " the" },
      { type: "equal", text: " the clause" },
      { type: "ins", text: " the" },
    ]);
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

  test("factors a usable common affix before applying the residual cell budget", () => {
    const shared = Array.from({ length: 2001 }, (_unused, index) => `clause${String(index)}`).join(
      " ",
    );
    const before = `${shared} former`;
    const after = `${shared} revised`;

    expect(diffWordSegments(before, after)).toEqual([
      { type: "equal", text: shared },
      { type: "del", text: " former" },
      { type: "ins", text: " revised" },
    ]);
  });

  test("uses unique anchors when the whole input exceeds the residual cell budget", () => {
    const shared = Array.from({ length: 2000 }, (_unused, index) => `clause${String(index)}`).join(
      " ",
    );
    const before = `before ${shared} former`;
    const after = `after ${shared} revised`;

    expect(diffWordSegments(before, after)).toEqual([
      { type: "del", text: "before" },
      { type: "ins", text: "after" },
      { type: "equal", text: ` ${shared}` },
      { type: "del", text: " former" },
      { type: "ins", text: " revised" },
    ]);
  });

  test("bounds unique-anchor discovery for very large residuals at either granularity", () => {
    const wordAnchors = Array.from(
      { length: 8193 },
      (_unused, index) => `anchor${String(index)}`,
    ).join(" ");
    const wordBefore = `before ${wordAnchors} former`;
    const wordAfter = `after ${wordAnchors} revised`;
    expect(diffWordSegments(wordBefore, wordAfter)).toEqual([
      { type: "del", text: wordBefore },
      { type: "ins", text: wordAfter },
    ]);

    const characterAnchors = Array.from({ length: 8193 }, (_unused, index) =>
      String.fromCodePoint(0x10_000 + index),
    ).join("");
    const characterBefore = `a${characterAnchors}b`;
    const characterAfter = `c${characterAnchors}d`;
    expect(diffWordSegments(characterBefore, characterAfter, { granularity: "character" })).toEqual(
      [
        { type: "del", text: characterBefore },
        { type: "ins", text: characterAfter },
      ],
    );
  });

  test("does not normalize every token after an oversized residual is refused", () => {
    let caseReads = 0;
    const normalization = {
      get case() {
        caseReads++;
        return true;
      },
    };
    const before = "ab".repeat(20_000);
    const after = "ba".repeat(20_000);

    expect(diffWordSegments(before, after, { granularity: "character", normalization })).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
    expect(caseReads).toBeLessThanOrEqual(10);
  });

  test("applies the token-storage gate before a cell-cheap asymmetric residual", () => {
    let caseReads = 0;
    const normalization = {
      get case() {
        caseReads++;
        return true;
      },
    };
    const before = "z";
    const after = "ab".repeat(10_000);

    expect(diffWordSegments(before, after, { granularity: "character", normalization })).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
    // Subsequence classification may inspect the long side once. A whole
    // comparison-key array would inspect every token a second time even
    // though the combined residual is already above its storage ceiling.
    expect(caseReads).toBeLessThanOrEqual(after.length + 10);
  });

  test("applies the code-unit gate before retaining normalized comparison keys", () => {
    let caseReads = 0;
    const normalization = {
      get case() {
        caseReads++;
        return true;
      },
    };
    const before = `${"a".repeat(600_000)}x`;
    const after = `${"b".repeat(600_000)}y`;

    expect(diffWordSegments(before, after, { normalization })).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
    // Prefix, suffix, and monotone classification compare the pair once each.
    // Building either residual or whole-input key arrays would read again.
    expect(caseReads).toBeLessThanOrEqual(6);
  });

  test("does not duplicate a capped residual's keys for the whole affixed input", () => {
    let caseReads = 0;
    const normalization = {
      get case() {
        caseReads++;
        return true;
      },
    };
    const before = "z";
    const after = `z${"a".repeat(16_384)}`;

    expect(diffWordSegments(before, after, { granularity: "character", normalization })).toEqual([
      { type: "equal", text: before },
      { type: "ins", text: after.slice(1) },
    ]);
    expect(caseReads).toBeLessThanOrEqual(16_400);
  });

  test("shares one dense-cell allowance across a comparison session", () => {
    const alternating = (first: string, second: string): string =>
      Array.from({ length: 2000 }, (_unused, index) => (index % 2 === 0 ? first : second)).join(
        " ",
      );
    const before = alternating("a", "b");
    const after = alternating("b", "a");
    const session = createWordDiffSession();

    const first = session.diff(before, after);
    expect(first.some(({ type }) => type === "equal")).toBe(true);
    expect(session.diff(before, after)).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
    // A standalone call owns a fresh allowance.
    expect(diffWordSegments(before, after)).toEqual(first);
  });
});
