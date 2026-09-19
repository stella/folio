import { describe, expect, test } from "bun:test";

import corpusManifest from "../corpus/sources.json";
import { CORPUS_TIERS, validateCorpusManifest } from "./lib/corpus-manifest";

const cloneManifest = (): Record<string, unknown> =>
  structuredClone(corpusManifest) as unknown as Record<string, unknown>;

const firstSource = (manifest: Record<string, unknown>): Record<string, unknown> => {
  const sources = manifest["sources"];
  if (!Array.isArray(sources)) {
    throw new Error("the committed manifest has no sources");
  }
  return sources[0] as Record<string, unknown>;
};

describe("the manifest's licence tier", () => {
  test("the committed manifest declares one per source", () => {
    expect(validateCorpusManifest(corpusManifest)).toEqual([]);
    for (const source of corpusManifest.sources) {
      expect(Object.values(CORPUS_TIERS)).toContain(source.tier);
      expect(source.tierReason.length).toBeGreaterThan(0);
    }
  });

  test("a source without a tier is refused", () => {
    const manifest = cloneManifest();
    const source = firstSource(manifest);
    delete source["tier"];
    expect(validateCorpusManifest(manifest)).toContain("sources[0].tier: expected 1, 2 or 3");
  });

  test("a tier outside the three is refused", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["tier"] = 4;
    expect(validateCorpusManifest(manifest)).toContain("sources[0].tier: expected 1, 2 or 3");
  });

  /**
   * The reason is the audit trail. A tier asserted with no stated reasoning is
   * a number nobody can check, which is the failure mode the field exists to
   * prevent.
   */
  test("a tier with no stated reasoning is refused", () => {
    const manifest = cloneManifest();
    firstSource(manifest)["tierReason"] = "";
    expect(validateCorpusManifest(manifest)).toContain(
      "sources[0].tierReason: expected a non-empty string",
    );
  });

  test("CI's default tier is populated, so a tier-1 run is never empty", () => {
    const permissive = corpusManifest.sources.filter(
      (source) => source.tier === CORPUS_TIERS.permissive,
    );
    expect(permissive.length).toBeGreaterThan(0);
  });
});
