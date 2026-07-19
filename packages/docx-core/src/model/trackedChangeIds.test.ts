import { describe, expect, test } from "bun:test";

import { MAX_REVISION_ID, normalizeRevisionId } from "./content";

describe("normalizeRevisionId (eigenpal #1093)", () => {
  test("leaves an in-range id unchanged", () => {
    expect(normalizeRevisionId(0)).toBe(0);
    expect(normalizeRevisionId(5)).toBe(5);
    expect(normalizeRevisionId(MAX_REVISION_ID)).toBe(MAX_REVISION_ID);
  });

  test("collapses a malformed id to 0", () => {
    expect(normalizeRevisionId(-3)).toBe(0);
    expect(normalizeRevisionId(1.5)).toBe(0);
    expect(normalizeRevisionId(Number.NaN)).toBe(0);
    expect(normalizeRevisionId(Number.POSITIVE_INFINITY)).toBe(0);
  });

  test("folds an out-of-range id into the int32 range", () => {
    expect(normalizeRevisionId(MAX_REVISION_ID + 6)).toBeLessThanOrEqual(MAX_REVISION_ID);
    expect(normalizeRevisionId(MAX_REVISION_ID + 6)).toBeGreaterThanOrEqual(0);
  });

  test("an out-of-range id folds to the same value as its in-range residue", () => {
    // 2^31 === MAX + 1, so 2^31 + 5 folds to 5.
    expect(normalizeRevisionId(2 ** 31 + 5)).toBe(normalizeRevisionId(5));
    expect(normalizeRevisionId(2 ** 31 + 5)).toBe(5);
  });
});
