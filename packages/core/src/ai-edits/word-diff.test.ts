/**
 * Unit tests for diffWordSegments: correctness of the normal O(n*m) LCS path,
 * plus a regression guard for the MAX_WORD_DIFF_CELLS bounded fallback.
 */

import { describe, expect, test } from "bun:test";

import { diffWordSegments } from "./word-diff";

describe("diffWordSegments", () => {
  test("reconstructs both strings from equal/del/ins segments for a small replacement", () => {
    const before = "The quick fox jumps.";
    const after = "The slow fox jumps.";

    const segments = diffWordSegments(before, after);

    const reconstructedBefore = segments
      .filter((s) => s.type !== "ins")
      .map((s) => s.text)
      .join("");
    const reconstructedAfter = segments
      .filter((s) => s.type !== "del")
      .map((s) => s.text)
      .join("");
    expect(reconstructedBefore).toBe(before);
    expect(reconstructedAfter).toBe(after);
    // "quick" -> "slow" is the only divergence; everything else is shared.
    expect(segments.some((s) => s.type === "del" && s.text.includes("quick"))).toBe(true);
    expect(segments.some((s) => s.type === "ins" && s.text.includes("slow"))).toBe(true);
  });

  test("does not use whitespace-only matches to interleave rewritten clauses", () => {
    const before =
      "The Supplier shall deliver the Goods within thirty days after receipt of the Purchase Order.";
    const after =
      "The Vendor must provide all Products no later than twenty business days following receipt of a valid order.";

    const segments = diffWordSegments(before, after);

    expect(segments).toEqual([
      { type: "equal", text: "The" },
      {
        type: "del",
        text: " Supplier shall deliver the Goods within thirty",
      },
      {
        type: "ins",
        text: " Vendor must provide all Products no later than twenty business",
      },
      { type: "equal", text: " days" },
      { type: "del", text: " after" },
      { type: "ins", text: " following" },
      { type: "equal", text: " receipt of" },
      { type: "del", text: " the Purchase Order." },
      { type: "ins", text: " a valid order." },
    ]);
    expect(
      segments.filter(({ type }) => type === "equal").every(({ text }) => /\S/u.test(text)),
    ).toBe(true);
    expect(
      segments
        .filter(({ type }) => type !== "ins")
        .map(({ text }) => text)
        .join(""),
    ).toBe(before);
    expect(
      segments
        .filter(({ type }) => type !== "del")
        .map(({ text }) => text)
        .join(""),
    ).toBe(after);
  });

  test("returns empty segments for two empty strings", () => {
    expect(diffWordSegments("", "")).toEqual([]);
  });

  test("falls back to a single whole-string del+ins pair once the token-count product exceeds MAX_WORD_DIFF_CELLS", () => {
    // Regression guard for the word-diff DoS fix: diffWordSegments used to
    // allocate an unbounded (m+1)*(n+1) DP table for two attacker-controlled
    // strings inside one `modified` block pair. 2,001 distinct words each
    // tokenize to 2,001 tokens; 2,001 * 2,001 = 4,004,001 cells, just over
    // the 4,000,000-cell budget, so the DP must be skipped entirely.
    const before = Array.from({ length: 2001 }, (_, i) => `before${i}`).join(" ");
    const after = Array.from({ length: 2001 }, (_, i) => `after${i}`).join(" ");

    const segments = diffWordSegments(before, after);

    expect(segments).toEqual([
      { type: "del", text: before },
      { type: "ins", text: after },
    ]);
  });
});
