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

  test("returns empty segments for two empty strings", () => {
    expect(diffWordSegments("", "")).toEqual([]);
  });

  test("falls back to a single whole-string del+ins pair once the token-count product exceeds MAX_WORD_DIFF_CELLS", () => {
    // Regression guard for the word-diff DoS fix: diffWordSegments used to
    // allocate an unbounded (m+1)*(n+1) DP table for two attacker-controlled
    // strings inside one `modified` block pair. 2,001 distinct words each
    // tokenize (tokenize() also emits the whitespace runs as their own
    // tokens) to 4,001 tokens; 4,001 * 4,001 = 16,008,001 cells, well over
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
