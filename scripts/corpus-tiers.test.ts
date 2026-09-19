import { describe, expect, test } from "bun:test";

import type { CorpusLock } from "./lib/corpus-manifest";
import {
  CorpusTierError,
  parseTierSelection,
  selectTiers,
  tierScopedLockDigest,
} from "./lib/corpus-tiers";

const lock = (
  sources: readonly { id: string; tier: 1 | 2 | 3; paths: readonly string[] }[],
): CorpusLock => ({
  schemaVersion: 1,
  manifestDigest: "digest",
  fileCount: sources.reduce((total, source) => total + source.paths.length, 0),
  totalBytes: 0,
  sources: sources.map(({ id, tier, paths }) => ({
    id,
    commit: "0".repeat(40),
    tier,
    files: paths.map((path, index) => ({ path, sha256: `${id}-${index}`, bytes: 10 })),
  })),
});

const TWO_TIERS = lock([
  { id: "permissive", tier: 1, paths: ["a.docx", "b.docx"] },
  { id: "copyleft", tier: 2, paths: ["c.docx"] },
]);

describe("parseTierSelection", () => {
  test("defaults to tier 1 alone, which is what CI runs", () => {
    expect(parseTierSelection(undefined)).toEqual([1]);
  });

  test("accepts a list and sorts it", () => {
    expect(parseTierSelection("2,1")).toEqual([1, 2]);
  });

  test("collapses a repeated tier", () => {
    expect(parseTierSelection("1,1")).toEqual([1]);
  });

  test("refuses a tier that does not exist", () => {
    expect(() => parseTierSelection("4")).toThrow(CorpusTierError);
    expect(() => parseTierSelection("")).toThrow(CorpusTierError);
  });
});

describe("selectTiers", () => {
  test("keeps only the selected tiers and recounts", () => {
    const selected = selectTiers(TWO_TIERS, [1]);
    expect(selected.sources.map((source) => source.id)).toEqual(["permissive"]);
    expect(selected.fileCount).toBe(2);
  });

  test("selecting every tier is the whole lock", () => {
    expect(selectTiers(TWO_TIERS, [1, 2]).fileCount).toBe(3);
  });
});

describe("tierScopedLockDigest", () => {
  /**
   * The point of the whole tier mechanism: a tier-1 baseline must survive the
   * addition of a tier-2 source, or every copyleft source added would force a
   * re-measurement of findings it cannot affect.
   */
  test("adding a tier-2 source leaves the tier-1 digest unchanged", () => {
    const before = lock([{ id: "permissive", tier: 1, paths: ["a.docx", "b.docx"] }]);
    expect(tierScopedLockDigest(TWO_TIERS, [1])).toBe(tierScopedLockDigest(before, [1]));
  });

  test("two tier selections never share a digest", () => {
    expect(tierScopedLockDigest(TWO_TIERS, [1])).not.toBe(tierScopedLockDigest(TWO_TIERS, [1, 2]));
  });

  test("a tier selection contributing no files still has its own digest", () => {
    const onlyTierOne = lock([{ id: "permissive", tier: 1, paths: ["a.docx"] }]);
    expect(tierScopedLockDigest(onlyTierOne, [1])).not.toBe(
      tierScopedLockDigest(onlyTierOne, [1, 2]),
    );
  });

  test("a file added to a selected tier changes the digest", () => {
    const grown = lock([
      { id: "permissive", tier: 1, paths: ["a.docx", "b.docx", "c.docx"] },
      { id: "copyleft", tier: 2, paths: ["c.docx"] },
    ]);
    expect(tierScopedLockDigest(grown, [1])).not.toBe(tierScopedLockDigest(TWO_TIERS, [1]));
  });
});
